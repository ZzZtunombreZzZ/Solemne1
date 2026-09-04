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

entrada = Path("entrada")
entrada.mkdir(parents=True, exist_ok=True)


# Formato de mediciones random
# tiene que haber al menos 3 registros por archivo, con 3 horarios distintos
# timestamp: Debe usar formato AAAA-MM-DDTHH:MM:SS
# estacion: ["STG0[1,2,3,4]","VAl0[1,2,3,4]", "CON0[1,2,3,4]", "PUN0[1,2,3,4])"]  # tienen que estar todas y que se tomen de manera random 
# temperatura: Número entre -20.0 y 60.0
# humedad: Número entre 0.0 y 100.0
# pm25: Número mayor o igual a 0.0
# ruido_db: Número entre 0.0 y 140.0
# formato de salida estacion_CODIGO_AAAAMMDD.json

for i in range(1, 16):
    # Generar un nombre de archivo único para cada estación y fecha
    estacion = f"STG0{i % 4 + 1}"  # Estaciones STG01 a STG04
    fecha = datetime.datetime.now().strftime("%Y%m%d")
    nombre_archivo = f"{estacion}_{fecha}.jsonl"
    ruta_archivo = entrada / nombre_archivo

    with open(ruta_archivo, "w") as archivo:
        for j in range(3):  # Generar 3 registros por archivo
            timestamp = (datetime.datetime.now() - datetime.timedelta(hours=j)).strftime("%Y-%m-%dT%H:%M:%S")
            medicion = {
                "timestamp": timestamp,
                "estacion": estacion,
                "temperatura": round(-20.0 + (80.0 * j / 2), 2),  # Valores entre -20.0 y 60.0
                "humedad": round(100.0 * j / 2, 2),  # Valores entre 0.0 y 100.0
                "pm25": round(10.0 * j, 2),  # Valores mayores o iguales a 0.0
                "ruido_db": round(140.0 * j / 2, 2)  # Valores entre 0.0 y 140.0
            }
            archivo.write(json.dumps(medicion) + "\n")

        # Agregar una línea inválida al final del archivo
        archivo.write("{invalid_json_line}\n")
