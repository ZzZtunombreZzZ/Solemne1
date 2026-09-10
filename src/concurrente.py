"""
Version CONCURRENTE del procesador de mediciones ambientales.

Reparte los archivos entre varios trabajadores threading.Thread mediante una
queue.Queue y protege los acumuladores globales con threading.Lock.
Implementa exactamente el mismo contrato de validacion que src/secuencial.py
para que ambas versiones entreguen metricas identicas.
"""

import json
import math
import queue
import re
import sys
import threading
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# Todas las rutas se resuelven contra la raiz del proyecto, no contra el
# directorio actual, para poder ejecutar el script desde cualquier lugar.
RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
DIR_ENTRADA = RAIZ_PROYECTO / "entrada"
DIR_SALIDA = RAIZ_PROYECTO / "salida"
DIR_ALERTAS = RAIZ_PROYECTO / "alertas"

# Campos obligatorios de cada medicion.
CAMPOS_REQUERIDOS = ("timestamp", "estacion", "temperatura", "humedad", "pm25", "ruido_db")

# Campos que deben ser numeros reales de Python (int o float, nunca bool).
CAMPOS_NUMERICOS = ("temperatura", "humedad", "pm25", "ruido_db")

# Umbrales de alerta (identicos en ambas versiones).
UMBRALES = {
    "temperatura": 35.0,
    "humedad": 20.0,
    "pm25": 35.0,
    "ruido_db": 75.0
}

FORMATO_TIMESTAMP = "%Y-%m-%dT%H:%M:%S"

# El nombre de entrada debe ser estacion_CODIGO_AAAAMMDD.jsonl. El patron se
# ancla contra Path.stem (la extension ya viene recortada) y prohibe guiones
# bajos dentro del codigo y de la fecha. No se usa replace(".jsonl", "")
# porque reemplaza TODAS las ocurrencias: "estacion_X.jsonlY_20240101.jsonl"
# se normalizaria a codigo "XY" y colisionaria con "estacion_XY_20240101.jsonl".
PATRON_NOMBRE = re.compile(r"^estacion_(?P<codigo>[^_]+)_(?P<fecha>[^_]+)$")

# Lectura tolerante de la entrada, identica en ambas versiones:
#  - "utf-8-sig" consume el BOM opcional para no invalidar la primera linea.
#  - errors="replace" evita que un byte no UTF-8 mate el hilo: la linea
#    danada queda con caracteres de reemplazo y cae como medicion invalida.
CODIFICACION_ENTRADA = "utf-8-sig"
ERRORES_ENTRADA = "replace"

NUM_TRABAJADORES = 3

# Distribucion del trabajo y exclusion mutua.
cola_archivos = queue.Queue()
mutex_globales = threading.Lock()
mutex_alertas = threading.Lock()
mutex_informes = threading.Lock()
mutex_stderr = threading.Lock()
mutex_errores = threading.Lock()

# Buffers de alertas acumulados en memoria, uno por archivo de entrada.
# Protegido por mutex_alertas. El volcado a alertas/alertas_detectadas.log
# ocurre una sola vez al final, en orden de nombre de archivo, para que el log
# sea byte a byte identico al de la version secuencial y estable entre corridas.
buffers_alertas = {}

# Acumuladores globales protegidos por mutex_globales.
global_archivos_proc = 0
global_lineas_totales = 0
global_validas = 0
global_invalidas = 0
global_alertas = 0
global_temp_sum = 0.0
global_hum_sum = 0.0
global_pm25_max = 0.0
global_ruido_max = 0.0
global_alertas_estacion = defaultdict(int)

# Registro compartido de archivos fallidos, protegido por mutex_errores.
# main() lo revisa al final para decidir el codigo de salida.
archivos_con_error = []


def avisar(mensaje):
    """Registra un aviso en stderr sin mezclar la salida de varios hilos."""
    with mutex_stderr:
        print(mensaje, file=sys.stderr)


def registrar_error(nombre_archivo, mensaje):
    """Anota un archivo fallido en la estructura compartida y avisa por stderr."""
    with mutex_errores:
        archivos_con_error.append(nombre_archivo)
    avisar(mensaje)


def es_numero_real(valor):
    """
    Acepta solo int o float reales y finitos.
    bool es subclase de int, se rechaza. inf, -inf y NaN se rechazan porque
    contaminan promedios y maximos. Un entero gigante que no cabe en float
    tampoco es un numero utilizable y se rechaza.
    """
    if isinstance(valor, bool):
        return False
    if not isinstance(valor, (int, float)):
        return False
    try:
        return math.isfinite(valor)
    except OverflowError:
        return False


def validar_fecha(fecha_str):
    """Valida que el timestamp use el formato AAAA-MM-DDTHH:MM:SS."""
    if not isinstance(fecha_str, str):
        return False
    try:
        datetime.strptime(fecha_str, FORMATO_TIMESTAMP)
        return True
    except ValueError:
        return False


def formatear_fecha_archivo(fecha_archivo):
    """
    Devuelve la fecha del nombre como AAAA-MM-DD solo si son 8 digitos.
    Si el nombre no calza con ese formato se devuelve el texto crudo, para
    no imprimir recortes sin sentido del estilo "2024--".
    """
    if len(fecha_archivo) == 8 and fecha_archivo.isascii() and fecha_archivo.isdigit():
        return f"{fecha_archivo[:4]}-{fecha_archivo[4:6]}-{fecha_archivo[6:]}"
    return fecha_archivo


def derivar_identidad(ruta_archivo):
    """
    Devuelve (codigo, fecha) derivados del nombre, o None si no calza.
    Se trabaja sobre Path.stem y con un patron anclado, nunca con un
    replace global de la extension.
    """
    coincidencia = PATRON_NOMBRE.match(ruta_archivo.stem)
    if coincidencia is None:
        return None
    return coincidencia.group("codigo"), coincidencia.group("fecha")


def planificar_archivos(archivos):
    """
    Resuelve la identidad de cada archivo y detecta colisiones de informe.

    Devuelve (plan, errores):
      plan    -> [(ruta, codigo, fecha)] en el orden de entrada recibido.
      errores -> [(nombre_archivo, mensaje)] para nombres no reconocidos y para
                 archivos distintos que apuntarian al mismo informe de salida.

    Dos entradas distintas que resuelven al mismo informe_CODIGO_FECHA.txt son
    una colision: gana la primera en orden de nombre y la segunda se registra
    como error en vez de sobrescribir el informe en silencio.

    Se ejecuta en el hilo principal antes de repartir la cola, para que la
    deteccion sea deterministica y no dependa de que trabajador llegue primero.
    """
    plan = []
    errores = []
    duenos_informe = {}

    for ruta_archivo in archivos:
        nombre_archivo = ruta_archivo.name

        identidad = derivar_identidad(ruta_archivo)
        if identidad is None:
            errores.append(
                (nombre_archivo,
                 f"AVISO: nombre de archivo no reconocido, se omite: {nombre_archivo}")
            )
            continue

        estacion_codigo, fecha_archivo = identidad
        nombre_informe = f"informe_{estacion_codigo}_{fecha_archivo}.txt"

        dueno = duenos_informe.get(nombre_informe)
        if dueno is not None:
            errores.append(
                (nombre_archivo,
                 f"ERROR: colisión de informe, {nombre_archivo} y {dueno} apuntan a "
                 f"{nombre_informe}; se omite {nombre_archivo}")
            )
            continue

        duenos_informe[nombre_informe] = nombre_archivo
        plan.append((ruta_archivo, estacion_codigo, fecha_archivo))

    return plan, errores


def validar_medicion(linea, estacion_codigo):
    """
    Devuelve el diccionario de la medicion si la linea es valida, o None.
    Nunca propaga excepciones: cualquier problema se traduce en None.
    """
    try:
        datos = json.loads(linea)
    except (json.JSONDecodeError, ValueError, RecursionError):
        # RecursionError deriva de RuntimeError, no de ValueError: hay que
        # nombrarlo aparte o una linea con anidamiento profundo mata el hilo.
        return None

    # Solo se aceptan objetos JSON.
    if not isinstance(datos, dict):
        return None

    # Estructura completa.
    for campo in CAMPOS_REQUERIDOS:
        if campo not in datos:
            return None

    # Consistencia de estacion y formato de fecha.
    if datos["estacion"] != estacion_codigo:
        return None
    if not validar_fecha(datos["timestamp"]):
        return None

    # Validacion de tipos ANTES de comparar rangos: comparar float con str
    # lanzaria TypeError. No se convierte con float(): un string es invalido.
    for campo in CAMPOS_NUMERICOS:
        if not es_numero_real(datos[campo]):
            return None

    # Validacion de rangos.
    if not (-20.0 <= datos["temperatura"] <= 60.0):
        return None
    if not (0.0 <= datos["humedad"] <= 100.0):
        return None
    if not (datos["pm25"] >= 0.0):
        return None
    if not (0.0 <= datos["ruido_db"] <= 140.0):
        return None

    return datos


def detectar_alertas(datos):
    """Devuelve la lista de alertas (indicador, valor, umbral) de una medicion."""
    alertas = []
    if datos["temperatura"] >= UMBRALES["temperatura"]:
        alertas.append(("Temperatura", datos["temperatura"], UMBRALES["temperatura"]))
    if datos["humedad"] <= UMBRALES["humedad"]:
        alertas.append(("Humedad", datos["humedad"], UMBRALES["humedad"]))
    if datos["pm25"] >= UMBRALES["pm25"]:
        alertas.append(("PM2.5", datos["pm25"], UMBRALES["pm25"]))
    if datos["ruido_db"] >= UMBRALES["ruido_db"]:
        alertas.append(("Ruido", datos["ruido_db"], UMBRALES["ruido_db"]))
    return alertas


def escribir_informe(estacion_codigo, fecha_archivo, nombre_archivo, resultado):
    """Escribe el informe individual salida/informe_CODIGO_AAAAMMDD.txt."""
    validas = resultado["validas"]
    prom_temp = resultado["sum_temp"] / validas if validas > 0 else 0.0
    prom_hum = resultado["sum_hum"] / validas if validas > 0 else 0.0
    fecha_formateada = formatear_fecha_archivo(fecha_archivo)
    primera_alerta = resultado["primera_alerta"] if resultado["primera_alerta"] else "Ninguna"

    ruta = DIR_SALIDA / f"informe_{estacion_codigo}_{fecha_archivo}.txt"
    # planificar_archivos() ya garantiza que dos entradas distintas no compartan
    # informe, pero la escritura igual va bajo mutex dedicado: el archivo de
    # salida es un recurso compartido y ningun par de trabajadores puede
    # escribirlo a la vez (una escritura desgarrada corrompe el informe).
    with mutex_informes, open(ruta, "w", encoding="utf-8", newline="\n") as f_out:
        f_out.write("INFORME DE ESTACIÓN AMBIENTAL\n")
        f_out.write(f"Archivo procesado: {nombre_archivo}\n")
        f_out.write(f"Estación: {estacion_codigo}\n")
        f_out.write(f"Fecha del archivo: {fecha_formateada}\n")
        f_out.write(f"Líneas leídas: {resultado['lineas']}\n")
        f_out.write(f"Mediciones válidas: {validas}\n")
        f_out.write(f"Mediciones inválidas: {resultado['invalidas']}\n")
        f_out.write(f"Temperatura promedio: {prom_temp:.2f} °C\n")
        f_out.write(f"Humedad promedio: {prom_hum:.2f} %\n")
        f_out.write(f"PM2.5 máximo: {resultado['pm25_max']:.2f} ug/m3\n")
        f_out.write(f"Ruido máximo: {resultado['ruido_max']:.2f} dB\n")
        f_out.write(f"Alertas detectadas: {resultado['alertas']}\n")
        f_out.write(f"Primera alerta: {primera_alerta}\n")


def analizar_archivo(ruta_archivo, estacion_codigo):
    """
    Recorre un archivo completo y devuelve sus contadores y sus alertas.
    Las alertas se acumulan en un buffer y no se escriben aqui, para que un
    archivo que falla a medio camino no deje lineas sueltas en el log.
    """
    nombre_archivo = ruta_archivo.name
    resultado = {
        "lineas": 0,
        "validas": 0,
        "invalidas": 0,
        "alertas": 0,
        "sum_temp": 0.0,
        "sum_hum": 0.0,
        "pm25_max": 0.0,
        "ruido_max": 0.0,
        "primera_alerta": None,
        "buffer_alertas": []
    }

    with open(ruta_archivo, "r", encoding=CODIFICACION_ENTRADA, errors=ERRORES_ENTRADA) as f:
        for linea in f:
            linea = linea.strip()
            # Regla 1 del contrato: la linea vacia se omite por completo.
            if not linea:
                continue

            resultado["lineas"] += 1
            datos = validar_medicion(linea, estacion_codigo)
            if datos is None:
                resultado["invalidas"] += 1
                continue

            resultado["validas"] += 1
            resultado["sum_temp"] += datos["temperatura"]
            resultado["sum_hum"] += datos["humedad"]
            if datos["pm25"] > resultado["pm25_max"]:
                resultado["pm25_max"] = datos["pm25"]
            if datos["ruido_db"] > resultado["ruido_max"]:
                resultado["ruido_max"] = datos["ruido_db"]

            alertas = detectar_alertas(datos)
            if alertas:
                if resultado["primera_alerta"] is None:
                    resultado["primera_alerta"] = datos["timestamp"]
                for indicador, valor, umbral in alertas:
                    resultado["alertas"] += 1
                    resultado["buffer_alertas"].append(
                        f"{nombre_archivo};{estacion_codigo};{datos['timestamp']};"
                        f"{indicador};{valor};{umbral}\n"
                    )

    return resultado


def estacion_con_mas_alertas(alertas_por_estacion):
    """
    Desempate deterministico: ante igualdad de alertas gana el codigo de
    estacion menor alfabeticamente, nunca el orden de insercion del dict.
    """
    if not alertas_por_estacion:
        return "Ninguna"
    return max(sorted(alertas_por_estacion), key=alertas_por_estacion.get)


def procesar_un_archivo(ruta_archivo, estacion_codigo, fecha_archivo):
    """
    Procesa un archivo tomado de la cola. Nunca propaga excepciones.
    La identidad (codigo, fecha) ya viene resuelta y libre de colisiones desde
    planificar_archivos(), asi que aqui no se deriva nada del nombre.
    """
    global global_archivos_proc, global_lineas_totales, global_validas, global_invalidas
    global global_alertas, global_temp_sum, global_hum_sum, global_pm25_max, global_ruido_max

    nombre_archivo = ruta_archivo.name

    try:
        resultado = analizar_archivo(ruta_archivo, estacion_codigo)
        escribir_informe(estacion_codigo, fecha_archivo, nombre_archivo, resultado)
    except Exception as error:
        # Un archivo que falla entero no se silencia ni se cuenta como
        # procesado: se avisa por stderr y cambia el codigo de salida.
        registrar_error(nombre_archivo,
                        f"ERROR: no se pudo procesar {nombre_archivo}: "
                        f"{type(error).__name__}: {error}")
        return

    # Log de alertas: el buffer del archivo se deposita en la estructura
    # compartida (protegida por mutex_alertas) recien cuando el archivo termino
    # bien. No se escribe a disco aqui: si cada trabajador volcara al terminar,
    # el orden de las lineas dependeria de que hilo gana la carrera y el log
    # dejaria de ser byte a byte identico al de la version secuencial.
    if resultado["buffer_alertas"]:
        with mutex_alertas:
            buffers_alertas[nombre_archivo] = resultado["buffer_alertas"]

    # Consolidacion global (solo archivos efectivamente procesados).
    with mutex_globales:
        global_archivos_proc += 1
        global_lineas_totales += resultado["lineas"]
        global_validas += resultado["validas"]
        global_invalidas += resultado["invalidas"]
        global_alertas += resultado["alertas"]
        global_temp_sum += resultado["sum_temp"]
        global_hum_sum += resultado["sum_hum"]
        global_alertas_estacion[estacion_codigo] += resultado["alertas"]
        if resultado["pm25_max"] > global_pm25_max:
            global_pm25_max = resultado["pm25_max"]
        if resultado["ruido_max"] > global_ruido_max:
            global_ruido_max = resultado["ruido_max"]


def trabajador():
    """
    Toma archivos de la cola hasta vaciarla y consolida bajo exclusion mutua.
    Ningun fallo puede matar al hilo: si un trabajador muriera, los archivos
    que quedan en la cola no se procesarian nunca y cola_archivos.join()
    quedaria bloqueado para siempre.
    """
    while True:
        try:
            ruta_archivo, estacion_codigo, fecha_archivo = cola_archivos.get_nowait()
        except queue.Empty:
            break

        try:
            procesar_un_archivo(ruta_archivo, estacion_codigo, fecha_archivo)
        except Exception as error:
            registrar_error(ruta_archivo.name,
                            f"ERROR: fallo inesperado en {ruta_archivo.name}: "
                            f"{type(error).__name__}: {error}")
        finally:
            cola_archivos.task_done()


def main(estacion_objetivo=None):
    """Devuelve el codigo de salida: 0 si todos los archivos se procesaron."""
    DIR_ENTRADA.mkdir(parents=True, exist_ok=True)
    DIR_SALIDA.mkdir(parents=True, exist_ok=True)
    DIR_ALERTAS.mkdir(parents=True, exist_ok=True)

    archivos = sorted(DIR_ENTRADA.glob("estacion_*.jsonl"))

    # Filtro opcional para generar solo el reporte de una estacion.
    if estacion_objetivo:
        archivos = [a for a in archivos if f"_{estacion_objetivo}_" in a.name]
        if not archivos:
            print(f"No se encontró un archivo para la estación {estacion_objetivo} en {DIR_ENTRADA}")
            return 0

    if not archivos:
        print(f"No se encontraron archivos de entrada en {DIR_ENTRADA}")
        return 0

    # Identidad y colisiones se resuelven antes de repartir la cola: dos
    # archivos de entrada distintos no pueden compartir el mismo informe.
    plan, errores_plan = planificar_archivos(archivos)
    for nombre_archivo, mensaje in errores_plan:
        registrar_error(nombre_archivo, mensaje)

    inicio_tiempo = time.time()

    for tarea in plan:
        cola_archivos.put(tarea)

    hilos = []
    for _ in range(NUM_TRABAJADORES):
        hilo = threading.Thread(target=trabajador)
        hilo.start()
        hilos.append(hilo)

    # Espera ordenada: primero la cola vacia, luego el termino de cada hilo.
    cola_archivos.join()
    for hilo in hilos:
        hilo.join()

    # Volcado unico del log de alertas, en orden de nombre de archivo de
    # entrada (el mismo que recorre la version secuencial). El lock se mantiene
    # alrededor del volcado: el archivo de alertas sigue siendo un recurso
    # compartido protegido por exclusion mutua.
    with mutex_alertas:
        with open(DIR_ALERTAS / "alertas_detectadas.log", "w", encoding="utf-8", newline="\n") as f_log:
            for ruta_archivo, _codigo, _fecha in plan:
                buffer_archivo = buffers_alertas.get(ruta_archivo.name)
                if buffer_archivo:
                    f_log.writelines(buffer_archivo)

    tiempo_total = time.time() - inicio_tiempo

    estacion_max = estacion_con_mas_alertas(global_alertas_estacion)

    prom_temp_global = global_temp_sum / global_validas if global_validas > 0 else 0.0
    prom_hum_global = global_hum_sum / global_validas if global_validas > 0 else 0.0

    # Los hilos ya terminaron, pero se lee bajo mutex por consistencia.
    with mutex_errores:
        fallidos = sorted(archivos_con_error)

    with open(DIR_SALIDA / "resumen_ambiental.txt", "w", encoding="utf-8", newline="\n") as f_resumen:
        f_resumen.write("RESUMEN CONSOLIDADO DE MONITOREO AMBIENTAL\n")
        f_resumen.write("Versión ejecutada: concurrente\n")
        f_resumen.write(f"Archivos procesados: {global_archivos_proc}\n")
        f_resumen.write(f"Líneas leídas: {global_lineas_totales}\n")
        f_resumen.write(f"Mediciones válidas: {global_validas}\n")
        f_resumen.write(f"Mediciones inválidas: {global_invalidas}\n")
        f_resumen.write(f"Temperatura promedio global: {prom_temp_global:.2f} °C\n")
        f_resumen.write(f"Humedad promedio global: {prom_hum_global:.2f} %\n")
        f_resumen.write(f"PM2.5 máximo global: {global_pm25_max:.2f} ug/m3\n")
        f_resumen.write(f"Ruido máximo global: {global_ruido_max:.2f} dB\n")
        f_resumen.write(f"Alertas totales: {global_alertas}\n")
        f_resumen.write(f"Estación con mayor cantidad de alertas: {estacion_max}\n")
        f_resumen.write(f"Tiempo total de ejecución: {tiempo_total:.3f} segundos\n")
        f_resumen.write(f"Cantidad de trabajadores: {NUM_TRABAJADORES}\n")
        # El resumen jamas declara una corrida limpia si hubo archivos fallidos.
        if fallidos:
            f_resumen.write(f"Archivos con error: {len(fallidos)}\n")
            f_resumen.write("ADVERTENCIA: corrida incompleta, archivos no procesados: "
                            f"{', '.join(fallidos)}\n")

    if fallidos:
        print(f"ERROR: {len(fallidos)} archivo(s) no se pudieron procesar.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    estacion = sys.argv[1] if len(sys.argv) > 1 else None
    sys.exit(main(estacion))
