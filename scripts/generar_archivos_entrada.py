"""
Este script genera 15 archivos de entrada en formato JSON en la carpeta "entrada". 
Cada archivo de entrada debe tener extensión. jsonl
y representar las mediciones de una única estación durante un período determinado.

Condiciones:
Tres horarios distintos.
Al menos una medición que active una alerta. 
Al menos una línea inválida, ya sea por JSONl mal formado, atributo ausente, estación inconsistente o valor fuera de rango.
"""

import json
from pathlib import Path
import datetime
import random

entrada = Path("entrada")
entrada.mkdir(parents=True, exist_ok=True)


# Formato de mediciones random
# Generar 15 registros por mediciones por archivo, con 3 horarios distintos y al menos una medición que active una alerta.
# timestamp: Debe usar formato AAAA-MM-DDTHH:MM:SS
# estacion: ["STG0[1,2,3,4]","VAl0[1,2,3,4]", "CON0[1,2,3,4]", "PUN0[1,2,3,4])"]  # tienen que estar todas y que se tomen de manera random 
# temperatura: Número entre -20.0 y 60.0
# humedad: Número entre 0.0 y 100.0
# pm25: Número mayor o igual a 0.0
# ruido_db: Número entre 0.0 y 140.0
# formato de salida estacion_CODIGO_AAAAMMDD.json

estaciones = ["STG01", "STG02", "STG03", "STG04", "VAL01", "VAL02", "VAL03", "VAL04",
              "CON01", "CON02", "CON03", "CON04", "PUN01", "PUN02", "PUN03", "PUN04"]

for estacion in estaciones:
    # Generar 15 registros por archivo
    registros = []
    for i in range(15):
        # Generar un timestamp aleatorio dentro de un rango de fechas
        fecha_inicio = datetime.datetime(2024, 1, 1)
        fecha_fin = datetime.datetime(2024, 12, 31)
        delta = fecha_fin - fecha_inicio
        random_seconds = random.randint(0, int(delta.total_seconds()))
        timestamp = (fecha_inicio + datetime.timedelta(seconds=random_seconds)).strftime("%Y-%m-%dT%H:%M:%S")

        # Generar valores aleatorios para las mediciones
        temperatura = round(random.uniform(-20.0, 60.0), 2)
        humedad = round(random.uniform(0.0, 100.0), 2)
        pm25 = round(random.uniform(0.0, 500.0), 2)  # pm25 puede ser mayor a 0
        ruido_db = round(random.uniform(0.0, 140.0), 2)

        # Crear el registro de medición
        registro = {
            "timestamp": timestamp,
            "estacion": estacion,
            "temperatura": temperatura,
            "humedad": humedad,
            "pm25": pm25,
            "ruido_db": ruido_db
        }

        # Agregar el registro a la lista de registros
        registros.append(registro)

    # Guardar los registros en un archivo JSONL
    nombre_archivo = f"estacion_{estacion}_{fecha_inicio.strftime('%Y%m%d')}.jsonl"
    with open(entrada / nombre_archivo, 'w') as f:
        for registro in registros:
            f.write(json.dumps(registro) + "\n")