"""
Gestor de incidencias ambientales.

Toma los informes generados en salida/, los clasifica segun la cantidad de
alertas detectadas, separa las alertas por indicador y deja un inventario
en gestion_ambiental/inventario_ambiental.json.

El gestor es idempotente: se puede ejecutar las veces que sea necesario y el
resultado final es el mismo. Los logs por indicador se reconstruyen en cada
corrida, la bitacora se reescribe completa en cada corrida y el inventario se
construye leyendo el estado real de las carpetas de destino, no solo lo movido
en la corrida actual.

Decisiones tomadas para que el entregable sea REPRODUCIBLE:

  - La bitacora logs/gestion_ambiental.log es POR CORRIDA: la primera linea de
    cada ejecucion trunca el archivo. Antes se abria en modo "a" y el archivo
    crecia, de modo que el tamano declarado en el README solo existia despues
    de la primera corrida. Ver el comentario de _BITACORAS_INICIADAS.
  - Los archivos auxiliares que el propio entregable deja en salida/
    (resumen_ambiental.txt y LEEME.txt) no se cuentan como anomalia. Antes
    LEEME.txt disparaba un falso positivo que obligaba a toda reproduccion a
    cerrar con al menos una anomalia. Ver ARCHIVOS_AUXILIARES_SALIDA.
  - Los marcadores .gitkeep no se cuentan como informes en el inventario.
    Ver ARCHIVOS_MARCADORES.

Para reproducir el entregable desde cero hay que ejecutar antes
scripts/limpiar.py, porque el arbol viaja con la salida de la corrida ya
documentada. Ver la seccion 2 del README.

Politica de control de errores:

  - Ningun error detiene el procesamiento: se registra en la bitacora y se
    continua con el resto del trabajo.
  - Ningun error se pierde: todo error capturado se cuenta como anomalia, se
    publica en el inventario y se refleja en la linea final de CONTROL DE
    ERRORES. La bitacora nunca declara una corrida limpia si hubo fallas.
  - La reparticion de alertas es atomica: el origen se lee completo antes de
    tocar los logs de destino, de modo que un origen ilegible o corrupto no
    puede dejar los logs por indicador vacios.
  - Un origen que se lee bien pero no aporta ninguna alerta tampoco pisa el
    historial: si el estado previo tenia alertas, se conserva y la perdida
    evitada queda declarada como anomalia.
  - Ninguna ruta de destino se asume del tipo correcto: antes de listar una
    carpeta o abrir un log se verifica que sea carpeta o archivo, y el
    inventario se construye protegido para que la corrida siempre cierre.
"""

import json
import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

# Rutas relativas a la raiz del proyecto para poder ejecutar desde cualquier cwd.
RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
DIR_SALIDA = RAIZ_PROYECTO / "salida"
DIR_ALERTAS = RAIZ_PROYECTO / "alertas"
DIR_GESTION = RAIZ_PROYECTO / "gestion_ambiental"
DIR_LOGS = RAIZ_PROYECTO / "logs"
DIR_INDICADORES = DIR_GESTION / "alertas_por_indicador"

# Categorias de clasificacion de informes.
CATEGORIAS = ("sin_alertas", "con_alertas", "criticas")

# Nombre de carpeta destino para cada indicador reconocido en el log de alertas.
# El log escribe "Temperatura", "Humedad", "PM2.5" y "Ruido", por eso se aceptan
# las variantes en minusculas de cada uno.
INDICADORES = {
    "temperatura": "temperatura",
    "humedad": "humedad",
    "pm2.5": "pm25",
    "pm25": "pm25",
    "ruido": "ruido",
}

# Orden fijo de indicadores para que el inventario sea reproducible.
CARPETAS_INDICADORES = ("temperatura", "humedad", "pm25", "ruido")

# informe_CODIGO_AAAAMMDD.txt
PATRON_INFORME = re.compile(r"^informe_[A-Za-z0-9]+_\d{8}(_v\d+)?\.txt$")

# Linea "Alertas detectadas: N" dentro del informe.
PATRON_ALERTAS_DETECTADAS = re.compile(r"^\s*alertas\s+detectadas\s*:\s*(\d+)\s*$", re.IGNORECASE)

NOMBRE_RESUMEN_ORIGEN = "resumen_ambiental.txt"
NOMBRE_RESUMEN = "resumen_resguardado.txt"
NOMBRE_BITACORA = "gestion_ambiental.log"
# Bitacora de emergencia: se usa solo si logs/ no esta disponible como carpeta.
NOMBRE_BITACORA_ALTERNATIVA = "gestion_ambiental_fallback.log"
# Caracter que deja Python al reemplazar un byte que no es UTF-8 valido (U+FFFD).
CARACTER_REEMPLAZO = chr(0xFFFD)
MAX_EJEMPLOS = 5

# ---------------------------------------------------------------------------
# Lista blanca de archivos auxiliares de salida/
# ---------------------------------------------------------------------------
# En salida/ conviven dos cosas distintas: los informes que el gestor debe
# clasificar y archivos auxiliares que estan ahi A PROPOSITO y que NO son
# informes:
#
#   - resumen_ambiental.txt : la pauta pide explicitamente CONSERVARLO en
#     salida/ (ademas de copiarlo como resumen_resguardado.txt), asi que jamas
#     debe moverse ni clasificarse.
#   - LEEME.txt : nota del equipo que le explica al corrector que los 20
#     informes individuales existieron en salida/ y donde quedaron despues de
#     que el gestor los movio.
#
# Ninguno de los dos calza con PATRON_INFORME. Sin esta lista blanca el gestor
# los trataba como "archivo que no corresponde al patron" y los contaba como
# anomalia. Era un FALSO POSITIVO con consecuencia directa sobre la
# reproducibilidad: cualquier persona que reprodujera el entregable cerraba con
# anomalias_registradas >= 1 y la bitacora nunca podia declarar corrida limpia,
# aunque no hubiera pasado nada malo.
#
# La lista es EXPLICITA y CERRADA: solo se ignoran estos nombres exactos.
# Cualquier otro archivo que no calce con el patron -- por ejemplo un informe
# mal nombrado como "informe_STG01.txt" o un "notas.txt" olvidado -- SIGUE
# registrandose como anomalia en la bitacora y en el inventario. La proteccion
# contra archivos inesperados no se debilita, solo deja de disparar sobre los
# dos archivos que el propio entregable pone ahi.
ARCHIVOS_AUXILIARES_SALIDA = frozenset({NOMBRE_RESUMEN_ORIGEN, "LEEME.txt"})

# Marcadores de carpeta versionados en Git (Git no versiona directorios vacios).
# gestion_ambiental/sin_alertas/ viaja con un .gitkeep porque el arbol
# obligatorio de la pauta exige esa carpeta y, con datos que cumplen la pauta,
# ningun informe termina con 0 alertas: sin el marcador la carpeta simplemente
# no existiria en el entregable. Un marcador NO es un informe, por eso se
# excluye del inventario: si se contara, sin_alertas apareceria con 1 informe
# inexistente y el total dejaria de cuadrar en 20.
ARCHIVOS_MARCADORES = frozenset({".gitkeep"})


def inicializar_directorios():
    """
    Crea la estructura de carpetas de trabajo si todavia no existe.

    Devuelve la lista de problemas encontrados. Si una ruta no se puede crear
    (por ejemplo, porque ya existe como archivo en vez de directorio) el
    problema se reporta como error de proceso y la corrida NO se detiene.
    """
    errores = []
    directorios = [DIR_LOGS, DIR_GESTION]
    directorios.extend(DIR_GESTION / categoria for categoria in CATEGORIAS)
    directorios.extend(DIR_INDICADORES / carpeta for carpeta in CARPETAS_INDICADORES)
    for directorio in directorios:
        try:
            directorio.mkdir(parents=True, exist_ok=True)
        except Exception as error:
            errores.append(f"no se pudo preparar {ruta_relativa(directorio)}: {error}")
    return errores


# Bitacoras a las que YA se escribio durante ESTA corrida del proceso.
#
# DECISION DOCUMENTADA: la bitacora es POR CORRIDA, no acumulativa.
#
# Antes el archivo se abria siempre en modo "a". Eso cumplia el requisito 12
# (generar logs/gestion_ambiental.log) pero rompia la reproducibilidad del
# entregable: la bitacora crecia en cada ejecucion, de modo que despues de la
# primera corrida ya era imposible volver a obtener el archivo que el README
# describe con stat (28 lineas, 2442 bytes). Un entregable que promete un
# tamano de bitacora que solo existe una vez no es reproducible.
#
# Se evaluaron dos salidas y se eligio TRUNCAR al inicio de cada corrida:
#
#   - Truncar: logs/gestion_ambiental.log queda siempre con la bitacora
#     COMPLETA de la ultima corrida. Cumple "mantener una bitacora de gestion"
#     y el requisito 12, y ademas es lo que la pauta manda inspeccionar con
#     "stat logs/gestion_ambiental.log": un archivo unico, con un tamano
#     estable y explicable.
#   - Rotar: dejaria gestion_ambiental.log.1, .log.2, ... El archivo principal
#     si seria reproducible, pero el arbol entregado se ensuciaria con
#     historicos que nadie pidio y cuya cantidad depende de cuantas veces se
#     ejecuto el gestor: la carpeta logs/ dejaria de ser reproducible aunque el
#     archivo lo fuera. Se descarto.
#
# El requisito 13 (controlar al menos un error sin detener el procesamiento)
# NO depende del historico acumulado y sigue siendo demostrable dentro de UNA
# sola corrida: cada corrida registra sus propias anomalias linea por linea,
# sigue procesando despues de cada una y cierra SIEMPRE con la linea
# "CONTROL DE ERRORES: ...", que declara cuantas anomalias hubo o que la
# corrida termino limpia. La evidencia de las 4 anomalias inyectadas
# (evidencias/debian/06-control-de-errores.txt) es justamente la bitacora de
# una unica corrida.
#
# El truncado se hace en la PRIMERA escritura de cada corrida y no al arrancar
# main(), para que valga tambien si el gestor se importa como modulo y para que
# la bitacora alternativa se trate igual que la principal.
_BITACORAS_INICIADAS = set()


def registrar_bitacora(mensaje):
    """
    Escribe una linea con marca de tiempo en logs/gestion_ambiental.log.

    La primera escritura de cada corrida TRUNCA el archivo y las siguientes
    agregan, de modo que la bitacora contiene siempre exactamente una corrida
    completa y el entregable es reproducible (ver _BITACORAS_INICIADAS).

    Nunca lanza excepcion. Si la bitacora principal no se puede escribir se
    intenta una bitacora alternativa dentro de gestion_ambiental/ y, como
    ultimo recurso, se emite el mensaje por la salida de error estandar. Asi
    un problema de carpetas nunca deja la corrida sin rastro.
    """
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linea = f"[{fecha}] {mensaje}\n"
    candidatas = (DIR_LOGS / NOMBRE_BITACORA, DIR_GESTION / NOMBRE_BITACORA_ALTERNATIVA)
    for ruta in candidatas:
        try:
            ruta.parent.mkdir(parents=True, exist_ok=True)
            # "w" solo en la primera linea de la corrida; "a" en las siguientes.
            modo = "a" if ruta in _BITACORAS_INICIADAS else "w"
            # newline con salto LF explicito: la bitacora se escribe con
            # saltos LF en TODOS los sistemas. Sin esto Python traduce el salto
            # al de la plataforma y en Windows el archivo pesaria 28 bytes mas
            # que en Debian, con lo que el tamano declarado en el README
            # (2442 bytes) dejaria de ser reproducible fuera de Linux.
            with open(ruta, modo, encoding="utf-8", newline="\n") as f:
                f.write(linea)
            _BITACORAS_INICIADAS.add(ruta)
            return
        except Exception:
            continue
    print(f"BITACORA NO DISPONIBLE -> {linea}", end="", file=sys.stderr)


def ruta_bitacora_efectiva():
    """Ruta de la bitacora que realmente quedo escrita en esta corrida."""
    principal = DIR_LOGS / NOMBRE_BITACORA
    if principal.exists():
        return principal
    return DIR_GESTION / NOMBRE_BITACORA_ALTERNATIVA


def obtener_nombre_seguro(ruta_destino):
    """Devuelve una ruta libre agregando sufijos _v1, _v2, ... si hace falta."""
    ruta_destino = Path(ruta_destino)
    if not ruta_destino.exists():
        return ruta_destino

    contador = 1
    while True:
        candidata = ruta_destino.with_name(f"{ruta_destino.stem}_v{contador}{ruta_destino.suffix}")
        if not candidata.exists():
            return candidata
        contador += 1


def ruta_relativa(ruta):
    """Ruta en formato portable respecto de la raiz del proyecto."""
    try:
        return Path(ruta).resolve().relative_to(RAIZ_PROYECTO).as_posix()
    except ValueError:
        return Path(ruta).as_posix()


def ruta_log_indicador(carpeta):
    """Ruta del log acumulado de un indicador."""
    return DIR_INDICADORES / carpeta / f"alertas_{carpeta}.log"


def leer_cantidad_alertas(ruta_informe):
    """
    Busca la linea "Alertas detectadas: N" recorriendo el informe linea por linea.

    Devuelve el numero de alertas o None si el informe no declara el dato.
    No se asume cero en silencio: quien llama decide como registrar la anomalia.

    Se lee con errors="replace" para que un byte danado no aborte la lectura
    completa del informe.
    """
    with open(ruta_informe, "r", encoding="utf-8", errors="replace") as f:
        for linea in f:
            coincidencia = PATRON_ALERTAS_DETECTADAS.match(linea)
            if coincidencia:
                return int(coincidencia.group(1))
    return None


def clasificar(alertas):
    """Traduce la cantidad de alertas a la carpeta de destino."""
    if alertas == 0:
        return "sin_alertas"
    if 1 <= alertas <= 3:
        return "con_alertas"
    return "criticas"


def resguardar_resumen():
    """
    Copia salida/resumen_ambiental.txt a gestion_ambiental/resumen_resguardado.txt.

    Se usa siempre el mismo nombre de destino para que repetir la corrida no
    acumule copias _v1, _v2, ... del mismo resumen. Devuelve la lista de errores
    capturados para que la corrida no pueda declararse limpia si algo fallo.
    """
    errores = []
    origen = DIR_SALIDA / NOMBRE_RESUMEN_ORIGEN
    destino = DIR_GESTION / NOMBRE_RESUMEN
    try:
        shutil.copy(origen, destino)
        registrar_bitacora(f"Resumen resguardado en {ruta_relativa(destino)}")
    except FileNotFoundError:
        # No es un error fatal: puede que el resumen ya se haya resguardado antes.
        if destino.exists():
            registrar_bitacora(
                f"ADVERTENCIA: no se encontro {ruta_relativa(origen)}. "
                f"Se conserva el resguardo previo {ruta_relativa(destino)}. Se continua."
            )
        else:
            registrar_bitacora(
                f"ADVERTENCIA: no se encontro {ruta_relativa(origen)} y no hay resguardo previo. Se continua."
            )
    except Exception as error:
        errores.append(f"copia del resumen: {error}")
        registrar_bitacora(f"ERROR copiando el resumen: {error}. Se continua.")
    return errores


def procesar_informes():
    """
    Clasifica y mueve los informes que esten en salida/.

    Devuelve un diccionario con:
      - "no_clasificados": informes que no declaran la cantidad de alertas.
        Se dejan en salida/ a proposito, para que queden visibles y no
        contaminen una categoria equivocada.
      - "ignorados": archivos presentes en salida/ que no calzan con el patron
        de informe. Antes se descartaban en silencio; ahora quedan registrados.
      - "errores": fallas capturadas durante el procesamiento.
      - "movidos": cantidad de informes movidos en esta corrida.
    """
    resultado = {"no_clasificados": [], "ignorados": [], "errores": [], "movidos": 0}

    if not DIR_SALIDA.exists():
        registrar_bitacora(
            f"ADVERTENCIA: la carpeta {ruta_relativa(DIR_SALIDA)} no existe. "
            "No hay informes nuevos que mover. Se continua."
        )
        return resultado

    try:
        contenido = sorted(DIR_SALIDA.iterdir())
    except Exception as error:
        resultado["errores"].append(f"listado de {ruta_relativa(DIR_SALIDA)}: {error}")
        registrar_bitacora(
            f"ERROR listando {ruta_relativa(DIR_SALIDA)}: {error}. Se continua sin mover informes."
        )
        return resultado

    for ruta_origen in contenido:
        if not ruta_origen.is_file():
            # Una carpeta (o cualquier entrada que no sea archivo) dentro de salida/
            # tampoco se descarta en silencio: queda registrada como anomalia igual
            # que un archivo mal nombrado.
            resultado["ignorados"].append(ruta_origen.name)
            registrar_bitacora(
                f"ANOMALIA: la entrada {ruta_origen.name} dentro de {ruta_relativa(DIR_SALIDA)} no es un "
                "archivo (es una carpeta u otro tipo de entrada). No se clasifica, queda en salida/ "
                "para revision manual. Se continua."
            )
            continue

        if not PATRON_INFORME.match(ruta_origen.name):
            # Los archivos auxiliares de salida/ estan ahi a proposito y no son
            # informes: el resumen se resguarda aparte y LEEME.txt es la nota
            # para el corrector. No se clasifican y NO son anomalia (ver el
            # comentario de ARCHIVOS_AUXILIARES_SALIDA). Cualquier OTRO archivo
            # que no calce con el patron si sigue contando como anomalia.
            if ruta_origen.name in ARCHIVOS_AUXILIARES_SALIDA:
                continue
            resultado["ignorados"].append(ruta_origen.name)
            registrar_bitacora(
                f"ANOMALIA: el archivo {ruta_origen.name} esta en {ruta_relativa(DIR_SALIDA)} pero no "
                "corresponde al patron informe_CODIGO_AAAAMMDD.txt. No se clasifica, queda en salida/ "
                "para revision manual. Se continua."
            )
            continue

        try:
            alertas = leer_cantidad_alertas(ruta_origen)

            if alertas is None:
                # Control de errores: se registra la anomalia y el proceso sigue.
                resultado["no_clasificados"].append(ruta_origen.name)
                registrar_bitacora(
                    f"ANOMALIA: el informe {ruta_origen.name} no declara la linea "
                    "'Alertas detectadas: N'. No se clasifica, queda en salida/ para revision manual. Se continua."
                )
                continue

            categoria = clasificar(alertas)
            destino = obtener_nombre_seguro(DIR_GESTION / categoria / ruta_origen.name)
            shutil.move(str(ruta_origen), str(destino))
            resultado["movidos"] += 1
            registrar_bitacora(f"Informe movido: {ruta_origen.name} -> {categoria} ({alertas} alertas)")

        except Exception as error:
            # Un informe con problemas no detiene el resto del procesamiento,
            # pero el error queda contabilizado y no se pierde.
            resultado["errores"].append(f"informe {ruta_origen.name}: {error}")
            registrar_bitacora(f"ERROR procesando el informe {ruta_origen.name}: {error}. Se continua.")

    registrar_bitacora(f"Informes movidos en esta corrida: {resultado['movidos']}")
    return resultado


def _descartar_temporales(temporales):
    """Borra los archivos temporales de indicadores que quedaron a medio camino."""
    for temporal in temporales.values():
        try:
            if temporal.exists():
                temporal.unlink()
        except Exception:
            # Un temporal que no se puede borrar no justifica detener la corrida.
            pass


def procesar_alertas():
    """
    Reparte alertas/alertas_detectadas.log en un log por indicador.

    La operacion es ATOMICA respecto de los logs definitivos:

      1. Se lee el origen COMPLETO antes de tocar cualquier destino. La lectura
         usa errors="replace" para que un byte que no sea UTF-8 valido no aborte
         el archivo entero.
      2. Se escribe el resultado en archivos temporales.
      3. Solo cuando todo salio bien, los temporales reemplazan a los logs
         definitivos.

    De este modo un origen ilegible jamas puede dejar los cuatro logs vacios ni
    el inventario en cero: si algo falla, se conserva el estado previo.
    """
    resultado = {
        "invalidas": [],
        "desconocidas": [],
        "lineas_corruptas": [],
        "errores": [],
        "advertencias_origen": [],
        "distribuidas": 0,
    }
    ruta_alertas_log = DIR_ALERTAS / "alertas_detectadas.log"

    if not ruta_alertas_log.exists():
        # No se truncan los logs por indicador: sin origen, se conserva el estado previo.
        registrar_bitacora(
            f"ADVERTENCIA: no se encontro {ruta_relativa(ruta_alertas_log)}. "
            "Se conservan los logs por indicador ya existentes. Se continua."
        )
        return resultado

    # PASO 1: lectura completa del origen ANTES de tocar ningun destino.
    try:
        with open(ruta_alertas_log, "r", encoding="utf-8", errors="replace") as origen:
            lineas_origen = origen.readlines()
    except Exception as error:
        resultado["errores"].append(f"lectura de {ruta_relativa(ruta_alertas_log)}: {error}")
        registrar_bitacora(
            f"ERROR leyendo {ruta_relativa(ruta_alertas_log)}: {error}. "
            "No se modifico ningun log por indicador: se conserva el estado previo. Se continua."
        )
        return resultado

    # PASO 2: se clasifica en memoria y se escribe en archivos temporales.
    temporales = {}
    manejadores = {}
    escritura_completa = False
    try:
        for carpeta in CARPETAS_INDICADORES:
            definitivo = ruta_log_indicador(carpeta)
            definitivo.parent.mkdir(parents=True, exist_ok=True)
            temporal = definitivo.with_name(definitivo.name + ".tmp")
            temporales[carpeta] = temporal
            # Salto LF explicito por la misma razon que en la bitacora: los logs
            # por indicador deben quedar byte a byte iguales en Debian y en Windows.
            manejadores[carpeta] = open(
                temporal, "w", encoding="utf-8", newline="\n"
            )

        for numero, linea in enumerate(lineas_origen, start=1):
            linea = linea.strip()
            if not linea:
                continue

            if CARACTER_REEMPLAZO in linea:
                # La linea se recupera igual, pero la corrupcion queda declarada.
                resultado["lineas_corruptas"].append({"linea": numero, "contenido": linea})
                registrar_bitacora(
                    f"ANOMALIA: la linea {numero} de {ruta_relativa(ruta_alertas_log)} contiene bytes "
                    "que no son UTF-8 valido. Se reemplazaron por el caracter de sustitucion y la "
                    "alerta se conserva. Se continua."
                )

            partes = linea.split(";")
            if len(partes) < 4 or not partes[3].strip():
                resultado["invalidas"].append({"linea": numero, "contenido": linea})
                registrar_bitacora(
                    f"ANOMALIA: alerta con formato invalido en la linea {numero} "
                    f"de {ruta_relativa(ruta_alertas_log)}: '{linea}'. Se descarta y se continua."
                )
                continue

            indicador = partes[3].strip().lower()
            carpeta = INDICADORES.get(indicador)
            if carpeta is None:
                resultado["desconocidas"].append(
                    {"linea": numero, "indicador": partes[3].strip(), "contenido": linea}
                )
                registrar_bitacora(
                    f"ANOMALIA: indicador desconocido '{partes[3].strip()}' en la linea {numero} "
                    f"de {ruta_relativa(ruta_alertas_log)}. Se descarta y se continua."
                )
                continue

            manejadores[carpeta].write(linea + "\n")
            resultado["distribuidas"] += 1

        escritura_completa = True

    except Exception as error:
        resultado["errores"].append(f"reparticion de alertas por indicador: {error}")
        registrar_bitacora(
            f"ERROR repartiendo las alertas por indicador: {error}. "
            "No se reemplazo ningun log definitivo: se conserva el estado previo. Se continua."
        )
    finally:
        for manejador in manejadores.values():
            try:
                manejador.close()
            except Exception:
                escritura_completa = False

    if not escritura_completa:
        _descartar_temporales(temporales)
        registrar_bitacora("Alertas distribuidas por indicador: 0 (se mantuvieron los logs anteriores)")
        resultado["distribuidas"] = 0
        return resultado

    # PASO 2.5: guarda explicita contra la destruccion silenciosa del historial.
    #
    # Una lectura EXITOSA que no produce ninguna alerta distribuible tambien es
    # peligrosa: si se dejara continuar, os.replace pisaria los cuatro logs
    # definitivos con temporales vacios y el inventario caeria a cero sin que
    # nadie se entere. El caso mas probable es correr el gestor antes que los
    # procesadores. Un origen legitimamente vacio con estado previo tambien
    # vacio NO es anomalia y sigue el camino normal.
    total_previo = total_alertas_previas()

    if resultado["distribuidas"] == 0 and total_previo > 0:
        _descartar_temporales(temporales)
        detalle = (
            f"el origen {ruta_relativa(ruta_alertas_log)} no aporta ninguna alerta distribuible, "
            f"pero el estado previo tenia {total_previo} alertas en los logs por indicador"
        )
        resultado["advertencias_origen"].append(detalle)
        registrar_bitacora("*** ANOMALIA GRAVE: EL ORIGEN DE ALERTAS NO APORTA NINGUNA ALERTA ***")
        registrar_bitacora(
            f"ANOMALIA: {detalle}. Se CONSERVA el estado previo: no se reemplazo ninguno de los "
            "logs por indicador. Verifique que secuencial.py o concurrente.py se hayan ejecutado "
            "antes que el gestor. Se continua."
        )
        registrar_bitacora("Alertas distribuidas por indicador: 0 (se mantuvieron los logs anteriores)")
        return resultado

    if 0 < resultado["distribuidas"] < total_previo:
        # Perdida parcial: si se reconstruye igual, al menos queda dicho cuanto se perdio.
        detalle = (
            f"el origen {ruta_relativa(ruta_alertas_log)} aporta {resultado['distribuidas']} alertas "
            f"y el estado previo tenia {total_previo}"
        )
        resultado["advertencias_origen"].append(detalle)
        registrar_bitacora(
            f"ADVERTENCIA: {detalle}. Se pierden {total_previo - resultado['distribuidas']} alertas "
            "del historial por indicador al reconstruir los logs. La anomalia queda registrada. Se continua."
        )

    # PASO 3: recien ahora se reemplazan los logs definitivos.
    reemplazos_fallidos = []
    for carpeta, temporal in temporales.items():
        definitivo = ruta_log_indicador(carpeta)
        try:
            os.replace(temporal, definitivo)
        except Exception as error:
            reemplazos_fallidos.append(carpeta)
            resultado["errores"].append(f"reemplazo del log de {carpeta}: {error}")
            registrar_bitacora(
                f"ERROR reemplazando el log de {carpeta}: {error}. "
                "Se conserva el log anterior de ese indicador. Se continua."
            )

    if reemplazos_fallidos:
        _descartar_temporales({c: temporales[c] for c in reemplazos_fallidos})
    else:
        registrar_bitacora("Logs por indicador reconstruidos de forma atomica")

    registrar_bitacora(f"Alertas distribuidas por indicador: {resultado['distribuidas']}")
    return resultado


def listar_informes(categoria):
    """
    Nombres de los informes que hay realmente en una carpeta de destino.

    Devuelve (nombres, errores). No basta con comprobar .exists(): si la ruta
    existe pero fue reemplazada por un ARCHIVO, iterdir() lanza NotADirectoryError
    y tumba la corrida entera. Aqui se verifica el TIPO de la ruta y cualquier
    fallo se registra como error de proceso en vez de propagarse.
    """
    carpeta = DIR_GESTION / categoria
    if not carpeta.exists():
        return [], []

    if not carpeta.is_dir():
        detalle = (
            f"{ruta_relativa(carpeta)} existe pero NO es una carpeta: no se puede inventariar "
            f"la categoria '{categoria}'"
        )
        registrar_bitacora(
            f"ERROR de estructura: {detalle}. Se inventaria la categoria como vacia y se continua."
        )
        return [], [detalle]

    try:
        # Los marcadores de carpeta (.gitkeep) no son informes y no se inventarian:
        # existen solo para que Git pueda versionar una carpeta vacia exigida por la
        # pauta. Ver el comentario de ARCHIVOS_MARCADORES.
        return sorted(
            ruta.name
            for ruta in carpeta.iterdir()
            if ruta.is_file() and ruta.name not in ARCHIVOS_MARCADORES
        ), []
    except Exception as error:
        detalle = f"listado de {ruta_relativa(carpeta)}: {error}"
        registrar_bitacora(
            f"ERROR listando {ruta_relativa(carpeta)}: {error}. "
            "Se inventaria la categoria como vacia y se continua."
        )
        return [], [detalle]


def contar_alertas_indicador(carpeta):
    """
    Cuenta las alertas realmente escritas en el log de un indicador.

    Devuelve (cantidad, errores). Igual que arriba, .exists() no alcanza: si el
    log fue reemplazado por un DIRECTORIO, open() lanza PermissionError (Windows)
    o IsADirectoryError (POSIX). Se verifica el TIPO y el fallo se registra.
    """
    ruta = ruta_log_indicador(carpeta)
    if not ruta.exists():
        return 0, []

    if not ruta.is_file():
        detalle = (
            f"{ruta_relativa(ruta)} existe pero NO es un archivo: no se pueden contar "
            f"las alertas del indicador '{carpeta}'"
        )
        registrar_bitacora(
            f"ERROR de estructura: {detalle}. Se cuenta 0 para ese indicador y se continua."
        )
        return 0, [detalle]

    try:
        with open(ruta, "r", encoding="utf-8", errors="replace") as f:
            return sum(1 for linea in f if linea.strip()), []
    except Exception as error:
        detalle = f"lectura de {ruta_relativa(ruta)}: {error}"
        registrar_bitacora(
            f"ERROR leyendo {ruta_relativa(ruta)}: {error}. "
            "Se cuenta 0 para ese indicador y se continua."
        )
        return 0, [detalle]


def total_alertas_previas():
    """
    Alertas que hay AHORA MISMO en los cuatro logs por indicador.

    Se usa como estado previo para decidir si el origen destruiria historial.
    Un indicador ilegible cuenta 0: la guarda debe poder decidir igual.
    """
    total = 0
    for carpeta in CARPETAS_INDICADORES:
        cantidad, _ = contar_alertas_indicador(carpeta)
        total += cantidad
    return total


def construir_inventario(
    anomalias_alertas, resultado_informes, errores_directorios, errores_resumen, estado_disponible=True
):
    """
    Arma el inventario a partir del ESTADO REAL de las carpetas de destino.

    No depende de lo movido en la corrida actual, por eso repetir la ejecucion
    no vacia el inventario.

    Toda anomalia cuenta: alertas invalidas, indicadores desconocidos, informes
    sin dato, archivos ignorados, lineas con bytes corruptos, advertencias sobre
    el origen de alertas y errores de proceso capturados por los except.

    Con estado_disponible=False no se recorre el disco: es el inventario de
    emergencia que se usa si la lectura del estado real falla por completo, para
    que la corrida igual cierre con inventario y con CONTROL DE ERRORES.
    """
    errores_estructura = []
    informes = {}
    alertas_por_indicador = {}

    if estado_disponible:
        for categoria in CATEGORIAS:
            nombres, errores = listar_informes(categoria)
            informes[categoria] = nombres
            errores_estructura.extend(errores)
        for carpeta in CARPETAS_INDICADORES:
            cantidad, errores = contar_alertas_indicador(carpeta)
            alertas_por_indicador[carpeta] = cantidad
            errores_estructura.extend(errores)
    else:
        informes = {categoria: [] for categoria in CATEGORIAS}
        alertas_por_indicador = {carpeta: 0 for carpeta in CARPETAS_INDICADORES}

    por_categoria = {categoria: len(informes[categoria]) for categoria in CATEGORIAS}

    invalidas = anomalias_alertas.get("invalidas", [])
    desconocidas = anomalias_alertas.get("desconocidas", [])
    corruptas = anomalias_alertas.get("lineas_corruptas", [])
    advertencias_origen = anomalias_alertas.get("advertencias_origen", [])
    no_clasificados = resultado_informes.get("no_clasificados", [])
    ignorados = resultado_informes.get("ignorados", [])

    # Los errores capturados por los except tambien son anomalias: si no se
    # cuentan, la bitacora puede declarar exito falso sobre una corrida rota.
    errores_de_proceso = (
        list(errores_directorios)
        + list(errores_resumen)
        + list(resultado_informes.get("errores", []))
        + list(anomalias_alertas.get("errores", []))
        + errores_estructura
    )

    total_anomalias = (
        len(invalidas)
        + len(desconocidas)
        + len(no_clasificados)
        + len(ignorados)
        + len(corruptas)
        + len(advertencias_origen)
        + len(errores_de_proceso)
    )

    ruta_resumen = DIR_GESTION / NOMBRE_RESUMEN

    inventario = {
        "resumen": {
            "total_informes": sum(por_categoria.values()),
            "informes_por_categoria": por_categoria,
            "total_alertas": sum(alertas_por_indicador.values()),
            "anomalias_registradas": total_anomalias,
            "errores_de_proceso": len(errores_de_proceso),
        },
        "informes": {
            categoria: {"cantidad": por_categoria[categoria], "archivos": informes[categoria]}
            for categoria in CATEGORIAS
        },
        "alertas_por_indicador": {
            carpeta: {
                "cantidad": alertas_por_indicador[carpeta],
                "log": ruta_relativa(ruta_log_indicador(carpeta)),
            }
            for carpeta in CARPETAS_INDICADORES
        },
        "total_alertas": sum(alertas_por_indicador.values()),
        "anomalias": {
            "total": total_anomalias,
            "alertas_invalidas": {
                "cantidad": len(invalidas),
                "ejemplos": invalidas[:MAX_EJEMPLOS],
            },
            "indicadores_desconocidos": {
                "cantidad": len(desconocidas),
                "ejemplos": desconocidas[:MAX_EJEMPLOS],
            },
            "informes_sin_dato_de_alertas": {
                "cantidad": len(no_clasificados),
                "archivos": sorted(no_clasificados),
            },
            "archivos_ignorados_en_salida": {
                "cantidad": len(ignorados),
                "archivos": sorted(ignorados),
            },
            "lineas_con_bytes_corruptos": {
                "cantidad": len(corruptas),
                "ejemplos": corruptas[:MAX_EJEMPLOS],
            },
            "advertencias_de_origen": {
                "cantidad": len(advertencias_origen),
                "detalles": advertencias_origen[:MAX_EJEMPLOS],
            },
            "errores_de_proceso": {
                "cantidad": len(errores_de_proceso),
                "detalles": errores_de_proceso[:MAX_EJEMPLOS],
            },
        },
        "resumen_resguardado": {
            "existe": ruta_resumen.exists(),
            "ruta": ruta_relativa(ruta_resumen),
        },
        "bitacora": ruta_relativa(ruta_bitacora_efectiva()),
        # Campo aparte, informativo: cambia en cada corrida y no forma parte del inventario reproducible.
        "marca_de_ejecucion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return inventario


def main():
    # La preparacion de carpetas tampoco puede tumbar el gestor: si una ruta
    # existe como archivo en vez de carpeta, se registra y se continua.
    errores_directorios = inicializar_directorios()
    registrar_bitacora("--- INICIO DE GESTION DE INCIDENCIAS ---")
    for detalle in errores_directorios:
        registrar_bitacora(
            f"ERROR preparando la estructura de carpetas: {detalle}. "
            "Se continua con lo que si esta disponible."
        )

    errores_resumen = resguardar_resumen()
    resultado_informes = procesar_informes()
    anomalias_alertas = procesar_alertas()

    # La construccion del inventario tampoco puede tumbar el gestor: recorre el
    # disco y una ruta con el tipo equivocado no debe dejar la corrida sin
    # inventario, sin CONTROL DE ERRORES y sin FIN DE GESTION.
    try:
        inventario = construir_inventario(
            anomalias_alertas, resultado_informes, errores_directorios, errores_resumen
        )
    except Exception as error:
        registrar_bitacora(
            f"ERROR construyendo el inventario a partir del estado real: {error}. "
            "Se genera un inventario de emergencia y se continua hasta cerrar la corrida."
        )
        inventario = construir_inventario(
            anomalias_alertas,
            resultado_informes,
            list(errores_directorios) + [f"construccion del inventario: {error}"],
            errores_resumen,
            estado_disponible=False,
        )

    ruta_inventario = DIR_GESTION / "inventario_ambiental.json"
    error_inventario = None
    try:
        # Salto LF explicito: el inventario tambien debe ser byte a byte igual
        # entre sistemas.
        with open(ruta_inventario, "w", encoding="utf-8", newline="\n") as f:
            json.dump(inventario, f, indent=4, ensure_ascii=False)
            f.write("\n")
        registrar_bitacora(
            f"Inventario generado en {ruta_relativa(ruta_inventario)}: "
            f"{inventario['resumen']['total_informes']} informes, "
            f"{inventario['resumen']['total_alertas']} alertas."
        )
    except Exception as error:
        error_inventario = f"escritura del inventario: {error}"
        registrar_bitacora(f"ERROR generando el inventario JSON: {error}. Se continua.")

    # Cierre explicito del control de errores exigido por la pauta.
    # La cuenta incluye los errores capturados por los except: la bitacora no
    # puede declarar una corrida limpia si algo se rompio durante el proceso.
    anomalias = inventario["anomalias"]
    total_anomalias = anomalias["total"]
    errores_de_proceso = anomalias["errores_de_proceso"]["cantidad"]
    if error_inventario:
        total_anomalias += 1
        errores_de_proceso += 1

    if total_anomalias:
        registrar_bitacora(
            f"CONTROL DE ERRORES: se detectaron {total_anomalias} anomalias "
            f"({anomalias['alertas_invalidas']['cantidad']} alertas con formato invalido, "
            f"{anomalias['indicadores_desconocidos']['cantidad']} con indicador desconocido, "
            f"{anomalias['lineas_con_bytes_corruptos']['cantidad']} lineas con bytes corruptos, "
            f"{anomalias['informes_sin_dato_de_alertas']['cantidad']} informes sin dato de alertas, "
            f"{anomalias['archivos_ignorados_en_salida']['cantidad']} archivos ignorados en salida, "
            f"{anomalias['advertencias_de_origen']['cantidad']} advertencias sobre el origen de alertas, "
            f"{errores_de_proceso} errores de proceso capturados). "
            "Todas quedaron registradas y el procesamiento CONTINUO hasta generar el inventario."
        )
    else:
        registrar_bitacora(
            "CONTROL DE ERRORES: no se detectaron anomalias en esta corrida. "
            "El procesamiento termino completo."
        )

    registrar_bitacora("--- FIN DE GESTION DE INCIDENCIAS ---")


if __name__ == "__main__":
    main()
