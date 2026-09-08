import os
import shutil
import json
import re
from datetime import datetime

DIR_SALIDA = "salida"
DIR_ALERTAS = "alertas"
DIR_GESTION = "gestion_ambiental"
DIR_LOGS = "logs"

def inicializar_directorios():
    directorios = [
        DIR_LOGS,
        f"{DIR_GESTION}/sin_alertas",
        f"{DIR_GESTION}/con_alertas",
        f"{DIR_GESTION}/criticas",
        f"{DIR_GESTION}/alertas_por_indicador/temperatura",
        f"{DIR_GESTION}/alertas_por_indicador/humedad",
        f"{DIR_GESTION}/alertas_por_indicador/pm25",
        f"{DIR_GESTION}/alertas_por_indicador/ruido"
    ]
    for d in directorios:
        os.makedirs(d, exist_ok=True)

def registrar_bitacora(mensaje):
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ruta_log = os.path.join(DIR_LOGS, "gestion_ambiental.log")
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write(f"[{fecha}] {mensaje}\n")

def obtener_nombre_seguro(ruta_destino):
    if not os.path.exists(ruta_destino):
        return ruta_destino
    
    base, ext = os.path.splitext(ruta_destino)
    contador = 1
    while os.path.exists(f"{base}_v{contador}{ext}"):
        contador += 1
    return f"{base}_v{contador}{ext}"

def procesar_informes():
    contadores = {"sin_alertas": 0, "con_alertas": 0, "criticas": 0}
    
    if not os.path.exists(DIR_SALIDA):
        registrar_bitacora(f"ERROR: La carpeta {DIR_SALIDA} no existe.")
        return contadores

    for archivo in os.listdir(DIR_SALIDA):
        if archivo.startswith("informe_") and archivo.endswith(".txt"):
            ruta_origen = os.path.join(DIR_SALIDA, archivo)
            
            try:
                alertas = 0
                with open(ruta_origen, 'r', encoding="utf-8") as f:
                    contenido = f.read()
                    match = re.search(r'(?i)(?:alertas\s*totales|cantidad\s*de\s*alertas|alertas)\s*:\s*(\d+)', contenido)
                    if match:
                        alertas = int(match.group(1))
                    else:
                        registrar_bitacora(f"ADVERTENCIA: No se encontró etiqueta de alertas en {archivo}. Se asume 0.")
                if alertas == 0:
                    carpeta_dest = "sin_alertas"
                    contadores["sin_alertas"] += 1
                elif 1 <= alertas <= 3:
                    carpeta_dest = "con_alertas"
                    contadores["con_alertas"] += 1
                else:
                    carpeta_dest = "criticas"
                    contadores["criticas"] += 1
                    
                ruta_dest_base = os.path.join(DIR_GESTION, carpeta_dest, archivo)
                ruta_dest_segura = obtener_nombre_seguro(ruta_dest_base)
                
                shutil.move(ruta_origen, ruta_dest_segura)
                registrar_bitacora(f"Informe movido: {archivo} -> {carpeta_dest} ({alertas} alertas)")
                
            except Exception as e:
                registrar_bitacora(f"ERROR procesando informe {archivo}: {e}")
                
    return contadores

def procesar_alertas():
    ruta_alertas_log = os.path.join(DIR_ALERTAS, "alertas_detectadas.log")
    
    archivos_indicadores = {
        "temperatura": f"{DIR_GESTION}/alertas_por_indicador/temperatura/alertas_temperatura.log",
        "humedad": f"{DIR_GESTION}/alertas_por_indicador/humedad/alertas_humedad.log",
        "pm2.5": f"{DIR_GESTION}/alertas_por_indicador/pm25/alertas_pm25.log",
        "ruido": f"{DIR_GESTION}/alertas_por_indicador/ruido/alertas_ruido.log"
    }
    
    try:
        with open(ruta_alertas_log, 'r', encoding="utf-8") as log_origen:
            for linea in log_origen:
                linea = linea.strip()
                if not linea:
                    continue
                
                partes = linea.split(';')
                if len(partes) >= 4:
                    indicador = partes[3].strip().lower()
                    
                    if indicador in archivos_indicadores:
                        ruta_destino = archivos_indicadores[indicador]
                        with open(ruta_destino, 'a', encoding="utf-8") as archivo_destino:
                            archivo_destino.write(linea + '\n')
                    else:
                        registrar_bitacora(f"Alerta con indicador desconocido '{indicador}': {linea}")
                else:
                    registrar_bitacora(f"Alerta con formato inválido: {linea}")
                    
    except FileNotFoundError:
        registrar_bitacora(f"ADVERTENCIA: No se encontró el log de alertas en {ruta_alertas_log}")
    except Exception as e:
        registrar_bitacora(f"ERROR procesando archivo de alertas: {e}")

def main():
    inicializar_directorios()
    registrar_bitacora("--- INICIO DE GESTIÓN DE INCIDENCIAS ---")
    
    ruta_resumen_origen = os.path.join(DIR_SALIDA, "resumen_ambiental.txt")
    ruta_resumen_destino = os.path.join(DIR_GESTION, "resumen_resguardado.txt")
    
    try:
        ruta_resumen_segura = obtener_nombre_seguro(ruta_resumen_destino)
        shutil.copy(ruta_resumen_origen, ruta_resumen_segura)
        registrar_bitacora(f"Resumen copiado correctamente a {ruta_resumen_segura}")
    except FileNotFoundError:
        registrar_bitacora(f"ADVERTENCIA: No se encontró {ruta_resumen_origen}")
    except Exception as e:
        registrar_bitacora(f"ERROR copiando resumen: {e}")
    
    inventario = procesar_informes()
    
    procesar_alertas()
    
    ruta_inventario = os.path.join(DIR_GESTION, "inventario_ambiental.json")
    try:
        with open(ruta_inventario, "w", encoding="utf-8") as f:
            json.dump(inventario, f, indent=4)
        registrar_bitacora("Inventario generado exitosamente")
    except Exception as e:
        registrar_bitacora(f"ERROR generando inventario JSON: {e}")
        
    registrar_bitacora("--- FIN DE GESTIÓN DE INCIDENCIAS ---")

if __name__ == "__main__":
    main()
