import threading
import queue
import json
import time
from datetime import datetime
from pathlib import Path
from collections import defaultdict

cola_archivos = queue.Queue()
mutex_globales = threading.Lock()
mutex_alertas = threading.Lock()

global_lineas_totales = 0
global_validas = 0
global_invalidas = 0
global_alertas = 0
global_temp_sum = 0.0
global_hum_sum = 0.0
global_pm25_max = 0.0
global_ruido_max = 0.0
global_alertas_estacion = defaultdict(int)

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
DIR_ENTRADA = RAIZ_PROYECTO / "entrada"
DIR_SALIDA = RAIZ_PROYECTO / "salida"
DIR_ALERTAS = RAIZ_PROYECTO / "alertas"

def preparar_directorios():
    DIR_ENTRADA.mkdir(parents=True, exist_ok=True)
    DIR_SALIDA.mkdir(parents=True, exist_ok=True)
    DIR_ALERTAS.mkdir(parents=True, exist_ok=True)

def procesar_archivo():
    global global_lineas_totales, global_validas, global_invalidas, global_alertas
    global global_temp_sum, global_hum_sum, global_pm25_max, global_ruido_max

    while True:
        try:
            archivo_path = cola_archivos.get_nowait()
        except queue.Empty:
            break

        l_lineas = 0
        l_validas = 0
        l_invalidas = 0
        l_alertas = 0
        l_temp_sum = 0.0
        l_hum_sum = 0.0
        l_pm25_max = 0.0
        l_ruido_max = 0.0
        primera_alerta = None

        nombre_archivo = archivo_path.name
        partes_nombre = nombre_archivo.replace('.jsonl', '').split('_')
        estacion_arch = partes_nombre[1] if len(partes_nombre) > 1 else "Desc"
        fecha_arch = partes_nombre[2] if len(partes_nombre) > 2 else "Desc"

        alertas_buffer = []

        try:
            with open(archivo_path, 'r', encoding='utf-8') as f:
                for linea in f:
                    l_lineas += 1
                    linea = linea.strip()
                    if not linea: continue
                    
                    try:
                        datos = json.loads(linea)
                        ts = datos["timestamp"]
                        est = datos["estacion"]
                        t = float(datos["temperatura"])
                        h = float(datos["humedad"])
                        pm = float(datos["pm25"])
                        r = float(datos["ruido_db"])

                        if est != estacion_arch: raise ValueError
                        if not (-20.0 <= t <= 60.0): raise ValueError
                        if not (0.0 <= h <= 100.0): raise ValueError
                        if pm < 0.0: raise ValueError
                        if not (0.0 <= r <= 140.0): raise ValueError
                        datetime.strptime(ts, "%Y-%m-%dT%H:%M:%S")

                        l_validas += 1
                        l_temp_sum += t
                        l_hum_sum += h
                        if pm > l_pm25_max: l_pm25_max = pm
                        if r > l_ruido_max: l_ruido_max = r

                        es_alerta = False
                        if t >= 35.0:
                            alertas_buffer.append(f"{nombre_archivo};{est};{ts};Temperatura;{t};35.0\n")
                            es_alerta = True
                        if h <= 20.0:
                            alertas_buffer.append(f"{nombre_archivo};{est};{ts};Humedad;{h};20.0\n")
                            es_alerta = True
                        if pm >= 35.0:
                            alertas_buffer.append(f"{nombre_archivo};{est};{ts};PM2.5;{pm};35.0\n")
                            es_alerta = True
                        if r >= 75.0:
                            alertas_buffer.append(f"{nombre_archivo};{est};{ts};Ruido;{r};75.0\n")
                            es_alerta = True

                        if es_alerta:
                            l_alertas += 1
                            if primera_alerta is None:
                                primera_alerta = ts

                    except (json.JSONDecodeError, KeyError, ValueError, TypeError):
                        l_invalidas += 1

            if alertas_buffer:
                with mutex_alertas:
                    with open(DIR_ALERTAS / "alertas_detectadas.log", "a", encoding="utf-8") as f_alerta:
                        f_alerta.writelines(alertas_buffer)

            prom_t = l_temp_sum / l_validas if l_validas > 0 else 0
            prom_h = l_hum_sum / l_validas if l_validas > 0 else 0
            str_primera = primera_alerta if primera_alerta else "Ninguna"
            nombre_salida = f"informe_{estacion_arch}_{fecha_arch}.txt"
            
            reporte = (
                f"INFORME ESTACION AMBIENTAL\n"
                f"Archivo procesado: {nombre_archivo}\n"
                f"Estacion: {estacion_arch}\n"
                f"Fecha del archivo: {fecha_arch}\n"
                f"Lineas leidas: {l_lineas}\n"
                f"Mediciones validas: {l_validas}\n"
                f"Mediciones invalidas: {l_invalidas}\n"
                f"Temperatura promedio: {prom_t:.2f} °C\n"
                f"Humedad promedio: {prom_h:.2f} %\n"
                f"PM2.5 maximo: {l_pm25_max:.2f} ug/m3\n"
                f"Ruido maximo: {l_ruido_max:.2f} dB\n"
                f"Alertas detectadas: {l_alertas}\n"
                f"Primera alerta: {str_primera}\n"
            )
            with open(DIR_SALIDA / nombre_salida, "w", encoding="utf-8") as f_rep:
                f_rep.write(reporte)

            with mutex_globales:
                global_lineas_totales += l_lineas
                global_validas += l_validas
                global_invalidas += l_invalidas
                global_alertas += l_alertas
                global_temp_sum += l_temp_sum
                global_hum_sum += l_hum_sum
                global_alertas_estacion[estacion_arch] += l_alertas
                if l_pm25_max > global_pm25_max: global_pm25_max = l_pm25_max
                if l_ruido_max > global_ruido_max: global_ruido_max = l_ruido_max

        except Exception:
            pass
        finally:
            cola_archivos.task_done()

def main():
    t_inicio = time.time()
    preparar_directorios()
    
    log_al = DIR_ALERTAS / "alertas_detectadas.log"
    if log_al.exists(): log_al.unlink()

    archivos = list(DIR_ENTRADA.glob("estacion_*.jsonl"))
    if not archivos: return
        
    for arch in archivos: cola_archivos.put(arch)

    hilos = []
    num_trabajadores = 3
    for _ in range(num_trabajadores):
        t = threading.Thread(target=procesar_archivo)
        t.start()
        hilos.append(t)

    cola_archivos.join()
    for t in hilos: t.join()

    t_fin = time.time() - t_inicio
    p_temp = global_temp_sum / global_validas if global_validas > 0 else 0
    
    if global_alertas_estacion:
        estacion_max = max(global_alertas_estacion, key=global_alertas_estacion.get)
    else:
        estacion_max = "Ninguna"

    resumen = (
        "RESUMEN DE MONITOREO AMBIENTAL\n"
        "Version ejecutada: concurrente\n"
        f"Archivos procesados: {len(archivos)}\n"
        f"Lineas leidas: {global_lineas_totales}\n"
        f"Mediciones validas: {global_validas}\n"
        f"Mediciones invalidas: {global_invalidas}\n"
        f"Temperatura promedio global: {p_temp:.2f} °C\n"
        f"PM2.5 maximo global: {global_pm25_max:.2f} ug/m3\n"
        f"Ruido maximo global: {global_ruido_max:.2f} dB\n"
        f"Alertas totales: {global_alertas}\n"
        f"Estacion con mayor cantidad de alertas: {estacion_max}\n"
        f"Tiempo total de ejecucion: {t_fin:.3f} segundos\n"
        f"Cantidad de trabajadores: {num_trabajadores}\n"
    )

    with open(DIR_SALIDA / "resumen_ambiental.txt", "w", encoding="utf-8") as f_res:
        f_res.write(resumen)

if __name__ == "__main__":
    main()