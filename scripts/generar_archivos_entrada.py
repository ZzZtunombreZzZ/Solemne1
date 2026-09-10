"""
Generador de archivos de entrada JSONL para el monitoreo ambiental.

Cada archivo representa las mediciones de una única estación durante un día y
cumple las exigencias de la pauta:
  - extensión .jsonl y nombre estacion_CODIGO_AAAAMMDD.jsonl
  - una línea = un objeto JSON independiente, sin cabecera
  - mínimo 15 mediciones por archivo
  - al menos tres horarios distintos dentro de la fecha del propio nombre
  - al menos una medición que dispare alerta
  - al menos una línea inválida
  - mínimo 20 archivos en el conjunto

DISTRIBUCIÓN CONTROLADA DE ALERTAS
El gestor de incidencias clasifica cada informe según la cantidad de alertas:
  0 alertas        -> gestion_ambiental/sin_alertas/
  1 a 3 alertas    -> gestion_ambiental/con_alertas/
  4 o más alertas  -> gestion_ambiental/criticas/
Para que la evidencia de clasificación no demuestre una sola rama, las
mediciones normales se construyen dentro de rangos que NO disparan alerta y
después se inyecta la cantidad exacta de alertas planificada por archivo:
8 archivos quedan con 1 a 3 alertas (con_alertas) y 12 archivos con 4 o más
(criticas). Nada de esto queda al azar.

CONSECUENCIA DOCUMENTADA SOBRE sin_alertas/
La pauta exige que TODO archivo de entrada tenga al menos una medición que
dispare alerta. Por lo tanto, con datos que cumplan la pauta, ningún informe
puede terminar con 0 alertas y la carpeta gestion_ambiental/sin_alertas/ queda
legítimamente vacía. No se generan archivos sin alertas para forzarla, porque
eso incumpliría la pauta. La carpeta se crea igual y su existencia vacía es
parte de la evidencia: demuestra que la rama existe y que ningún caso la activó.

IMPORTANTE SOBRE EL CONTEO DE ALERTAS
El procesador cuenta una alerta por cada INDICADOR que supera su umbral, no una
por medición. Una sola medición puede aportar hasta cuatro alertas. Por eso el
reparto se hace a nivel de indicador.

Umbrales que disparan alerta (ver src/secuencial.py):
  temperatura >= 35.0 | humedad <= 20.0 | pm25 >= 35.0 | ruido_db >= 75.0
Rangos seguros que NO disparan alerta:
  temperatura < 35.0  | humedad > 20.0  | pm25 < 35.0  | ruido_db < 75.0

Reproducibilidad: semilla fija y recorrido en orden explícito, de modo que la
salida es idéntica entre corridas y entre Windows y Debian.
"""

import json
import random
from pathlib import Path

# Rutas relativas a la raíz del proyecto, no al directorio de trabajo actual.
RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
DIR_ENTRADA = RAIZ_PROYECTO / "entrada"

# Semilla fija: la pauta exige resultados reproducibles.
SEMILLA = 20240908

MEDICIONES_POR_ARCHIVO = 15

# Rangos seguros: ningún valor sorteado aquí puede disparar una alerta.
RANGO_SEGURO = {
    "temperatura": (10.0, 34.0),
    "humedad": (25.0, 90.0),
    "pm25": (2.0, 30.0),
    "ruido_db": (35.0, 70.0),
}

# Rangos de alerta: todo valor sorteado aquí supera el umbral del indicador
# pero sigue siendo válido para el validador (no se sale del dominio físico).
RANGO_ALERTA = {
    "temperatura": (36.0, 48.0),
    "humedad": (3.0, 18.0),
    "pm25": (40.0, 180.0),
    "ruido_db": (78.0, 110.0),
}

INDICADORES = ("temperatura", "humedad", "pm25", "ruido_db")

# Horas del día usadas en los timestamps. Son ocho, muy por encima de los tres
# horarios distintos que exige la pauta, y todas caen dentro de la fecha del
# nombre del archivo.
HORAS_DEL_DIA = (0, 3, 6, 9, 12, 15, 18, 21)

# Los cinco tipos de línea inválida que el validador debe descartar. Se reparten
# entre los 20 archivos para que la validación quede demostrada de verdad y no
# con veinte copias del mismo caso.
TIPOS_INVALIDA = (
    "json_malformado",       # texto que no parsea como JSON
    "atributo_ausente",      # falta un campo obligatorio
    "estacion_inconsistente",  # el campo estacion no coincide con el nombre
    "valor_fuera_rango",     # pm25 negativo, fuera del dominio físico
    "tipo_incorrecto",       # temperatura entregada como string
)

# Plan de generación: 20 archivos con códigos de estación coherentes.
# Las 16 estaciones cubren la fecha 2024-01-01; las cuatro estaciones de
# Santiago repiten con una segunda fecha para completar los 20 archivos.
# El tercer elemento es la cantidad EXACTA de alertas planificada por archivo.
PLAN_ARCHIVOS = (
    # (estacion, fecha, alertas_planificadas)  -> categoría esperada
    ("STG01", "20240101", 1),   # con_alertas
    ("STG02", "20240101", 5),   # criticas
    ("STG03", "20240101", 2),   # con_alertas
    ("STG04", "20240101", 6),   # criticas
    ("VAL01", "20240101", 3),   # con_alertas
    ("VAL02", "20240101", 4),   # criticas
    ("VAL03", "20240101", 1),   # con_alertas
    ("VAL04", "20240101", 7),   # criticas
    ("CON01", "20240101", 2),   # con_alertas
    ("CON02", "20240101", 4),   # criticas
    ("CON03", "20240101", 3),   # con_alertas
    ("CON04", "20240101", 5),   # criticas
    ("PUN01", "20240101", 2),   # con_alertas
    ("PUN02", "20240101", 6),   # criticas
    ("PUN03", "20240101", 1),   # con_alertas
    ("PUN04", "20240101", 4),   # criticas
    ("STG01", "20240102", 8),   # criticas
    ("STG02", "20240102", 5),   # criticas
    ("STG03", "20240102", 4),   # criticas
    ("STG04", "20240102", 6),   # criticas
)


def repartir_alertas(total, desfase):
    """Reparte 'total' alertas de indicador entre mediciones concretas.

    Devuelve una lista de tuplas; cada tupla es una medición y contiene los
    indicadores que esa medición debe disparar. La suma de los largos es
    exactamente 'total'. El 'desfase' rota qué indicadores se usan, para que no
    todos los archivos alerten siempre por lo mismo. Es determinístico: no
    interviene el azar.
    """
    patrones = []
    restante = total
    paso = desfase
    while restante > 0:
        # Cada tantas mediciones se dispara un par de indicadores a la vez,
        # así se demuestra que una sola medición puede aportar más de una alerta.
        cantidad = 2 if (paso % 3 == 0 and restante >= 2) else 1
        patrones.append(tuple(INDICADORES[(paso + j) % 4] for j in range(cantidad)))
        restante -= cantidad
        paso += 1
    return patrones


def construir_timestamp(fecha_archivo, indice):
    """Arma un timestamp que cae SIEMPRE dentro de la fecha del nombre.

    'fecha_archivo' viene en formato AAAAMMDD y el timestamp resultante usa el
    formato AAAA-MM-DDTHH:MM:SS que exige el validador.
    """
    anio = fecha_archivo[:4]
    mes = fecha_archivo[4:6]
    dia = fecha_archivo[6:8]
    hora = HORAS_DEL_DIA[indice % len(HORAS_DEL_DIA)]
    minuto = (indice * 7) % 60
    segundo = (indice * 13) % 60
    return f"{anio}-{mes}-{dia}T{hora:02d}:{minuto:02d}:{segundo:02d}"


def valor_seguro(indicador):
    """Sortea un valor dentro del rango que NO dispara alerta."""
    minimo, maximo = RANGO_SEGURO[indicador]
    return round(random.uniform(minimo, maximo), 2)


def valor_de_alerta(indicador):
    """Sortea un valor que supera el umbral del indicador y sigue siendo válido."""
    minimo, maximo = RANGO_ALERTA[indicador]
    return round(random.uniform(minimo, maximo), 2)


def construir_mediciones(estacion, fecha_archivo, alertas_planificadas, desfase):
    """Construye las 15 mediciones válidas de un archivo.

    Primero se arman todas las mediciones dentro de rangos seguros y después se
    inyectan las alertas exactas en las últimas posiciones. Así la cantidad de
    alertas del archivo es la planificada, no una consecuencia del azar.
    """
    patrones = repartir_alertas(alertas_planificadas, desfase)
    indice_primera_alerta = MEDICIONES_POR_ARCHIVO - len(patrones)

    mediciones = []
    for i in range(MEDICIONES_POR_ARCHIVO):
        medicion = {
            "timestamp": construir_timestamp(fecha_archivo, i),
            "estacion": estacion,
            "temperatura": valor_seguro("temperatura"),
            "humedad": valor_seguro("humedad"),
            "pm25": valor_seguro("pm25"),
            "ruido_db": valor_seguro("ruido_db"),
        }
        # Inyección explícita de las alertas planificadas.
        if i >= indice_primera_alerta:
            for indicador in patrones[i - indice_primera_alerta]:
                medicion[indicador] = valor_de_alerta(indicador)
        mediciones.append(medicion)
    return mediciones


def construir_linea_invalida(estacion, fecha_archivo, tipo):
    """Devuelve el texto exacto de la línea inválida del archivo.

    Se devuelve texto y no un diccionario porque uno de los casos es un JSON
    mal formado, que por definición no puede representarse como objeto.
    """
    # El timestamp de la línea inválida también respeta la fecha del nombre.
    timestamp = construir_timestamp(fecha_archivo, MEDICIONES_POR_ARCHIVO)

    if tipo == "json_malformado":
        # Llave sin cerrar y valor sin comillas: json.loads falla.
        return '{"timestamp": "' + timestamp + '", "estacion": ' + estacion + ', "temperatura": '

    if tipo == "atributo_ausente":
        # Falta el campo obligatorio pm25.
        return json.dumps({
            "timestamp": timestamp,
            "estacion": estacion,
            "temperatura": 24.5,
            "humedad": 55.0,
            "ruido_db": 48.0,
        })

    if tipo == "estacion_inconsistente":
        # El código de estación no coincide con el del nombre del archivo.
        estacion_ajena = "XXX99"
        return json.dumps({
            "timestamp": timestamp,
            "estacion": estacion_ajena,
            "temperatura": 24.5,
            "humedad": 55.0,
            "pm25": 12.0,
            "ruido_db": 48.0,
        })

    if tipo == "valor_fuera_rango":
        # pm25 negativo: fuera del dominio físico admitido.
        return json.dumps({
            "timestamp": timestamp,
            "estacion": estacion,
            "temperatura": 24.5,
            "humedad": 55.0,
            "pm25": -1.0,
            "ruido_db": 48.0,
        })

    if tipo == "tipo_incorrecto":
        # temperatura viene como string en vez de número.
        return json.dumps({
            "timestamp": timestamp,
            "estacion": estacion,
            "temperatura": "25.5",
            "humedad": 55.0,
            "pm25": 12.0,
            "ruido_db": 48.0,
        })

    raise ValueError(f"Tipo de línea inválida desconocido: {tipo}")


def generar():
    """Genera el conjunto completo de archivos de entrada e imprime la evidencia."""
    random.seed(SEMILLA)
    DIR_ENTRADA.mkdir(parents=True, exist_ok=True)

    # Se limpia la corrida anterior para que el conjunto sea exactamente el
    # planificado y no una mezcla con archivos viejos.
    for archivo_previo in sorted(DIR_ENTRADA.glob("estacion_*.jsonl")):
        archivo_previo.unlink()

    detalle = []
    for indice, (estacion, fecha_archivo, alertas) in enumerate(PLAN_ARCHIVOS):
        tipo_invalida = TIPOS_INVALIDA[indice % len(TIPOS_INVALIDA)]

        mediciones = construir_mediciones(estacion, fecha_archivo, alertas, indice)
        linea_invalida = construir_linea_invalida(estacion, fecha_archivo, tipo_invalida)

        nombre_archivo = f"estacion_{estacion}_{fecha_archivo}.jsonl"
        with open(DIR_ENTRADA / nombre_archivo, "w", encoding="utf-8", newline="\n") as f:
            for medicion in mediciones:
                f.write(json.dumps(medicion) + "\n")
            f.write(linea_invalida + "\n")

        categoria = "con_alertas" if 1 <= alertas <= 3 else "criticas"
        horas_distintas = len({m["timestamp"][11:13] for m in mediciones})
        detalle.append({
            "archivo": nombre_archivo,
            "mediciones": len(mediciones),
            "horas_distintas": horas_distintas,
            "alertas": alertas,
            "categoria": categoria,
            "invalida": tipo_invalida,
        })

    imprimir_resumen(detalle)
    return detalle


def imprimir_resumen(detalle):
    """Imprime el resumen que sirve como evidencia de la corrida."""
    print("=" * 78)
    print("GENERACIÓN DE ARCHIVOS DE ENTRADA")
    print("=" * 78)
    print(f"Carpeta de salida: {DIR_ENTRADA}")
    print(f"Semilla fija: {SEMILLA} (resultados reproducibles entre corridas y sistemas)")
    print(f"Archivos generados: {len(detalle)}")
    print()
    print(f"{'ARCHIVO':<34} {'MED':>4} {'HORAS':>6} {'ALERTAS':>8} {'CATEGORÍA':<12} LÍNEA INVÁLIDA")
    print("-" * 78)
    for fila in detalle:
        print(f"{fila['archivo']:<34} {fila['mediciones']:>4} {fila['horas_distintas']:>6} "
              f"{fila['alertas']:>8} {fila['categoria']:<12} {fila['invalida']}")

    con_alertas = [f for f in detalle if f["categoria"] == "con_alertas"]
    criticas = [f for f in detalle if f["categoria"] == "criticas"]
    total_alertas = sum(f["alertas"] for f in detalle)

    print("-" * 78)
    print("CLASIFICACIÓN ESPERADA EN gestion_ambiental/")
    print(f"  sin_alertas/  : 0 informes  (la pauta exige al menos una alerta por archivo,")
    print(f"                              por eso esta rama queda legítimamente vacía)")
    print(f"  con_alertas/  : {len(con_alertas)} informes (1 a 3 alertas)")
    print(f"  criticas/     : {len(criticas)} informes (4 o más alertas)")
    print(f"  Alertas totales esperadas: {total_alertas}")
    print()
    print("LÍNEAS INVÁLIDAS REPARTIDAS (una por archivo)")
    for tipo in TIPOS_INVALIDA:
        cantidad = sum(1 for f in detalle if f["invalida"] == tipo)
        print(f"  {tipo:<24}: {cantidad} archivos")
    print("=" * 78)


if __name__ == "__main__":
    generar()
