"""
Deja el arbol del proyecto en el estado PREVIO a la corrida documentada.

POR QUE EXISTE ESTE SCRIPT
--------------------------
El entregable tiene dos obligaciones que se estorban entre si:

  1. La pauta exige entregar los informes YA CLASIFICADOS dentro de
     gestion_ambiental/con_alertas/ y gestion_ambiental/criticas/, mas el
     inventario, la bitacora y los logs por indicador. Es la evidencia del
     trabajo, asi que el arbol viaja con la salida de la corrida oficial.

  2. La pauta exige un README reproducible: cualquier persona debe poder
     seguir los pasos y llegar a los mismos numeros.

Si se ejecutan los pasos del README sobre el arbol tal como se entrega, los
programas vuelven a producir los 20 informes y el gestor los clasifica ENCIMA
de los 20 que ya venian. Como el gestor esta obligado por la pauta (punto 6) a
no sobrescribir, crea 20 duplicados _v1 y el resultado pasa de 20 informes a
40. No es un error del gestor: es que la reproduccion parte de un arbol que ya
tiene la corrida hecha.

Este script resuelve el choque: borra SOLO lo que los programas vuelven a
generar y deja intacto todo lo que es fuente o evidencia. Despues de correrlo,
los pasos del README parten del mismo estado en que partio la corrida oficial y
entregan exactamente 0 / 8 / 12 informes, 79 alertas y 0 anomalias.

QUE BORRA
---------
  salida/informe_*.txt                      informes individuales regenerables
  salida/resumen_ambiental.txt              resumen consolidado regenerable
  alertas/alertas_detectadas.log            log plano de alertas
  gestion_ambiental/sin_alertas/*           informes ya clasificados
  gestion_ambiental/con_alertas/*
  gestion_ambiental/criticas/*
  gestion_ambiental/alertas_por_indicador/*/*.log   logs por indicador
  gestion_ambiental/inventario_ambiental.json
  gestion_ambiental/resumen_resguardado.txt
  gestion_ambiental/gestion_ambiental_fallback.log  bitacora de emergencia
  logs/gestion_ambiental.log                bitacora de gestion
  *.tmp sueltos de los logs por indicador

QUE CONSERVA (nunca se toca)
----------------------------
  entrada/                 los 20 .jsonl del conjunto oficial
  salida/LEEME.txt         nota del equipo para el corrector
  .gitkeep                 marcadores de carpeta exigida por la pauta
  evidencias/              incluido evidencias/salida_parte1/
  docs/, src/, scripts/, README.md

El script es idempotente y no falla si algo ya no existe. Termina siempre en 0.
"""

import sys
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent

DIR_SALIDA = RAIZ_PROYECTO / "salida"
DIR_ALERTAS = RAIZ_PROYECTO / "alertas"
DIR_GESTION = RAIZ_PROYECTO / "gestion_ambiental"
DIR_LOGS = RAIZ_PROYECTO / "logs"
DIR_INDICADORES = DIR_GESTION / "alertas_por_indicador"

CATEGORIAS = ("sin_alertas", "con_alertas", "criticas")
CARPETAS_INDICADORES = ("temperatura", "humedad", "pm25", "ruido")

# Nombres que JAMAS se borran, aunque esten dentro de una carpeta que se limpia.
# LEEME.txt es parte del entregable y .gitkeep es lo unico que hace existir
# gestion_ambiental/sin_alertas/ en Git, que la pauta exige en su arbol.
NOMBRES_PROTEGIDOS = frozenset({"LEEME.txt", ".gitkeep"})


def ruta_relativa(ruta):
    """Ruta portable respecto de la raiz del proyecto, para los mensajes."""
    try:
        return ruta.resolve().relative_to(RAIZ_PROYECTO).as_posix()
    except ValueError:
        return ruta.as_posix()


def borrar(ruta, borrados, errores):
    """Borra un archivo si existe. Un fallo se reporta pero no detiene la limpieza."""
    if ruta.name in NOMBRES_PROTEGIDOS:
        return
    try:
        if ruta.is_file():
            ruta.unlink()
            borrados.append(ruta_relativa(ruta))
    except Exception as error:
        errores.append(f"{ruta_relativa(ruta)}: {error}")


def vaciar_carpeta(carpeta, borrados, errores):
    """Borra los archivos de una carpeta respetando los nombres protegidos."""
    if not carpeta.is_dir():
        return
    try:
        contenido = sorted(carpeta.iterdir())
    except Exception as error:
        errores.append(f"listado de {ruta_relativa(carpeta)}: {error}")
        return
    for ruta in contenido:
        borrar(ruta, borrados, errores)


def main():
    borrados = []
    errores = []

    # salida/: se van los informes y el resumen; LEEME.txt queda protegido.
    vaciar_carpeta(DIR_SALIDA, borrados, errores)

    # alertas/: el log plano lo regenera secuencial.py / concurrente.py.
    borrar(DIR_ALERTAS / "alertas_detectadas.log", borrados, errores)

    # gestion_ambiental/: informes clasificados, logs por indicador, inventario,
    # resumen resguardado y bitacora de emergencia.
    for categoria in CATEGORIAS:
        vaciar_carpeta(DIR_GESTION / categoria, borrados, errores)
    for carpeta in CARPETAS_INDICADORES:
        vaciar_carpeta(DIR_INDICADORES / carpeta, borrados, errores)
    borrar(DIR_GESTION / "inventario_ambiental.json", borrados, errores)
    borrar(DIR_GESTION / "resumen_resguardado.txt", borrados, errores)
    borrar(DIR_GESTION / "gestion_ambiental_fallback.log", borrados, errores)

    # logs/: la bitacora de gestion.
    borrar(DIR_LOGS / "gestion_ambiental.log", borrados, errores)

    print("LIMPIEZA PREVIA A LA REPRODUCCION")
    print("=================================")
    print(f"Raiz del proyecto: {RAIZ_PROYECTO}")
    print(f"Archivos eliminados: {len(borrados)}")
    for ruta in borrados:
        print(f"  - {ruta}")
    if not borrados:
        print("  (el arbol ya estaba limpio)")

    if errores:
        print(f"\nNo se pudieron eliminar {len(errores)} rutas:")
        for detalle in errores:
            print(f"  ! {detalle}")

    print("\nSe conservaron intactos: entrada/, salida/LEEME.txt, los marcadores")
    print(".gitkeep, evidencias/ (incluido evidencias/salida_parte1/), docs/,")
    print("src/, scripts/ y README.md.")
    print("\nEl arbol quedo listo para ejecutar los pasos 1 a 5 del README.")

    # Siempre 0: la limpieza es una ayuda, no un paso que pueda tumbar la entrega.
    return 0


if __name__ == "__main__":
    sys.exit(main())
