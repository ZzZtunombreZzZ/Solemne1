import time
import json
import glob
import os
from datetime import datetime

dir_entrada="entrada"
dir_salida="salida"
dir_alertas="alertas"

#Tolerancias para las alertas
umbrales={
    "temperatura": 35.0,
    "humedad": 20.0,
    "pm25": 35.0,
    "ruido": 75.0
}

def v_fecha(fecha_str):
    try:
        datetime.strptime(fecha_str, "%Y-%m-%dT%H:%M:%S")
        return True
    except ValueError:
        return False

def pro_ar_sec():
    #Para crear archivos y no tener errores
    os.makedirs(dir_salida, exist_ok=True)
    os.makedirs(dir_alertas, exist_ok=True)

    ar= glob.glob(os.path.jopin(dir_entrada, "*.jsonl"))

    global_ar_proc=0
    global_lineas=0
    global_validas=0
    global_invalidas=0
    global_sum_temp=0.0
    global_pm25_max=0.0
    global_ale_total=0
    ale_estacion={}

    ini_tiempo=time.time()

    log_ale_path=os.path.join(dir_alertas, "alertas_detectadas.log")

    with open(log_ale_path, "w", encoding="utf-8") as f_log:
        for r_ar in ar:
            n_ar=os.path.basename(r_ar)
            part= n_ar.replace(".jsonl", "").split("_")
            if len(part) !=3:
                continue
            estacion_cod= part[1]
            fecha_ar= part[2]
            
            loc_lineas=0
            loc_validas=0
            loc_invalidas=0
            loc_sum_temp=0.0
            loc_sum_hum=0.0
            loc_pm25_max=0.0
            loc_ruido_max=0
            primera_alerta=None

            if estacion_cod not in ale_estacion:
                ale_estacion[estacion_cod] = 0
            with open(r_ar, "r", encoding="utf-8") as f:
                for linea in f:
                    loc_lineas += 1
                    try:
                        data=json.loads(linea.strip())
                        #para valdiar estructura y atributos
                        camops_req=["timestamp", "estacion", "termperatura", "humedad", "pm25", "ruido"]