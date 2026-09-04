import os
import glob
import json
import time
from datetime import datetime

DIR_ENTRADA = "entrada"
DIR_SALIDA = "salida"
DIR_ALERTAS = "alertas"

# tolerancias
UMBRALES = {
    "temperatura": 35.0,
    "humedad": 20.0,
    "pm25": 35.0,
    "ruido_db": 75.0
}

def validar_fecha(fecha_str):
    """Valida que el timestamp use el formato AAAA-MM-DDTHH:MM:SS"""
    try:
        datetime.strptime(fecha_str, "%Y-%m-%dT%H:%M:%S")
        return True
    except ValueError:
        return False

def procesar_archivos_secuencial():
    # crea los directorios si no existe, evita errores.
    os.makedirs(DIR_SALIDA, exist_ok=True)
    os.makedirs(DIR_ALERTAS, exist_ok=True)

    archivos = glob.glob(os.path.join(DIR_ENTRADA, "*.jsonl"))
    
    # variables
    global_archivos_proc = 0
    global_lineas = 0
    global_validas = 0
    global_invalidas = 0
    global_sum_temp = 0.0
    global_pm25_max = 0.0
    global_ruido_max = 0.0
    global_alertas_totales = 0
    alertas_por_estacion = {}
    
    # iniciar la medicion de tiempo
    inicio_tiempo = time.time()
    
    log_alertas_path = os.path.join(DIR_ALERTAS, "alertas_detectadas.log")
    
    with open(log_alertas_path, "w", encoding="utf-8") as f_log:
        for ruta_archivo in archivos:
            nombre_archivo = os.path.basename(ruta_archivo)
            
            # copiar estación y fecha del nombre
            partes = nombre_archivo.replace(".jsonl", "").split("_")
            if len(partes) != 3:
                continue 
            
            estacion_codigo = partes[1]
            fecha_archivo = partes[2]
            
            #contadores para el reporte
            loc_lineas = 0
            loc_validas = 0
            loc_invalidas = 0
            loc_sum_temp = 0.0
            loc_sum_hum = 0.0
            loc_pm25_max = 0.0
            loc_ruido_max = 0.0
            loc_alertas = 0
            primera_alerta = None
            
            if estacion_codigo not in alertas_por_estacion:
                alertas_por_estacion[estacion_codigo] = 0

            with open(ruta_archivo, "r", encoding="utf-8") as f:
                for linea in f:
                    loc_lineas += 1
                    try:
                        data = json.loads(linea.strip())
                        
                        # validacion de estructura
                        campos_requeridos = ["timestamp", "estacion", "temperatura", "humedad", "pm25", "ruido_db"]
                        if not all(k in data for k in campos_requeridos):
                            raise ValueError("Atributo ausente")
                            
                        # consistencia y validacion
                        if data["estacion"] != estacion_codigo:
                            raise ValueError("Estación inconsistente")
                        if not validar_fecha(data["timestamp"]):
                            raise ValueError("Formato de fecha inválido")
                        if not (-20.0 <= data["temperatura"] <= 60.0):
                            raise ValueError("Temperatura fuera de rango")
                        if not (0.0 <= data["humedad"] <= 100.0):
                            raise ValueError("Humedad fuera de rango")
                        if data["pm25"] < 0.0:
                            raise ValueError("PM2.5 inválido")
                        if not (0.0 <= data["ruido_db"] <= 140.0):
                            raise ValueError("Ruido fuera de rango")
                    
                        loc_validas += 1
                        loc_sum_temp += data["temperatura"]
                        loc_sum_hum += data["humedad"]
                        
                        if data["pm25"] > loc_pm25_max: loc_pm25_max = data["pm25"]
                        if data["ruido_db"] > loc_ruido_max: loc_ruido_max = data["ruido_db"]
                        
                        # revision de alertas
                        alertas_detectadas = []
                        if data["temperatura"] >= UMBRALES["temperatura"]:
                            alertas_detectadas.append(("Temperatura", data["temperatura"], UMBRALES["temperatura"]))
                        if data["humedad"] <= UMBRALES["humedad"]:
                            alertas_detectadas.append(("Humedad", data["humedad"], UMBRALES["humedad"]))
                        if data["pm25"] >= UMBRALES["pm25"]:
                            alertas_detectadas.append(("PM2.5", data["pm25"], UMBRALES["pm25"]))
                        if data["ruido_db"] >= UMBRALES["ruido_db"]:
                            alertas_detectadas.append(("Ruido", data["ruido_db"], UMBRALES["ruido_db"]))
                            
                        if alertas_detectadas:
                            if primera_alerta is None:
                                primera_alerta = data["timestamp"]
                            
                            for ind, val, umbral in alertas_detectadas:
                                loc_alertas += 1
                                f_log.write(f"{nombre_archivo};{estacion_codigo};{data['timestamp']};{ind};{val};{umbral}\n")
                                
                    except (json.JSONDecodeError, ValueError):
                        # detecta archivos JSON mal formados
                        loc_invalidas += 1
                        
            # crear informe individual en salida
            prom_temp = loc_sum_temp / loc_validas if loc_validas > 0 else 0.0
            prom_hum = loc_sum_hum / loc_validas if loc_validas > 0 else 0.0
            fecha_formateada = f"{fecha_archivo[:4]}-{fecha_archivo[4:6]}-{fecha_archivo[6:]}"
            
            with open(os.path.join(DIR_SALIDA, f"informe_{estacion_codigo}_{fecha_archivo}.txt"), "w", encoding="utf-8") as f_out:
                f_out.write("INFORME DE ESTACIÓN AMBIENTAL\n")
                f_out.write(f"Archivo procesado: {nombre_archivo}\n")
                f_out.write(f"Estación: {estacion_codigo}\n")
                f_out.write(f"Fecha del archivo: {fecha_formateada}\n")
                f_out.write(f"Líneas leídas: {loc_lineas}\n")
                f_out.write(f"Mediciones válidas: {loc_validas}\n")
                f_out.write(f"Mediciones inválidas: {loc_invalidas}\n")
                f_out.write(f"Temperatura promedio: {prom_temp:.2f}°C\n")
                f_out.write(f"Humedad promedio: {prom_hum:.2f}%\n")
                f_out.write(f"PM2.5 máximo: {loc_pm25_max:.2f} ug/m3\n")
                f_out.write(f"Ruido máximo: {loc_ruido_max:.2f} dB\n")
                f_out.write(f"Alertas detectadas: {loc_alertas}\n")
                if primera_alerta:
                    f_out.write(f"Primera alerta: {primera_alerta}\n")

            # juntar datos para el resumen global
            global_archivos_proc += 1
            global_lineas += loc_lineas
            global_validas += loc_validas
            global_invalidas += loc_invalidas
            global_sum_temp += loc_sum_temp
            if loc_pm25_max > global_pm25_max: global_pm25_max = loc_pm25_max
            if loc_ruido_max > global_ruido_max: global_ruido_max = loc_ruido_max
            global_alertas_totales += loc_alertas
            alertas_por_estacion[estacion_codigo] += loc_alertas

    # termino medición de tiempo y escritura del resumen
    tiempo_total = time.time() - inicio_tiempo
    
    estacion_max_alertas = max(alertas_por_estacion, key=alertas_por_estacion.get) if alertas_por_estacion else "N/A"
    prom_temp_global = global_sum_temp / global_validas if global_validas > 0 else 0.0
    
    with open(os.path.join(DIR_SALIDA, "resumen_ambiental.txt"), "w", encoding="utf-8") as f_resumen:
        f_resumen.write("RESUMEN CONSOLIDADO DE MONITOREO AMBIENTAL\n")
        f_resumen.write("Versión ejecutada: secuencial\n")
        f_resumen.write(f"Archivos procesados: {global_archivos_proc}\n")
        f_resumen.write(f"Líneas leídas: {global_lineas}\n")
        f_resumen.write(f"Mediciones válidas: {global_validas}\n")
        f_resumen.write(f"Mediciones inválidas: {global_invalidas}\n")
        f_resumen.write(f"Temperatura promedio global: {prom_temp_global:.2f}°C\n")
        f_resumen.write(f"PM2.5 máximo global: {global_pm25_max:.2f} ug/m3\n")
        f_resumen.write(f"Ruido máximo global: {global_ruido_max:.2f} dB\n")
        f_resumen.write(f"Alertas totales: {global_alertas_totales}\n")
        f_resumen.write(f"Estación con mayor cantidad de alertas: {estacion_max_alertas}\n")
        f_resumen.write(f"Tiempo total de ejecución: {tiempo_total:.3f} segundos\n")
        f_resumen.write("Cantidad de trabajadores: 1\n")

if __name__ == "__main__":
    procesar_archivos_secuencial()            