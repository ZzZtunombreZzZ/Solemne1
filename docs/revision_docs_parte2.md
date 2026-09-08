# Guía de Acciones Pendientes - Documentación Parte 2 (Forma B)

Este documento contiene **exclusivamente las acciones concretas y pendientes** para completar con éxito la documentación y entregables de la **Parte 2**, según la pauta y rúbrica oficial (`Solemne01PracticaParte2FormaB.pdf`).

---

## Resumen Ejecutivo de Acciones

```
DOCUMENTACIÓN
├── 1. Hito Presencial (hito_parte2hp_equipo07.pdf)
│   ├── [x] Corregir nombre de integrante: corregido a "José Palma" en hito_parte2hp_equipo07.pdf
│   ├── [ ] Reducir a 1 o 2 páginas (actualmente tiene 5)
│   ├── [ ] Pegar redacción en "Flujo planificado" (estaba vacío)
│   ├── [ ] Agregar sección "Cinco evidencias para el informe final" (faltaba)
│   └── [ ] Agregar riesgo de sobrescritura del gestor (_v1, _v2)
│
└── 2. Informe Final Asincrónico (8 a 12 páginas en PDF)
    ├── [ ] Tomar 4 capturas pendientes en la VM Debian (ps, free/df, stat, error)
    ├── [ ] Integrar capturas existentes de "evidencias/Parte 2 solemne..."
    ├── [ ] Redactar explicación técnica: Mover vs Copiar y campos del sistema
    └── [ ] Exportar PDF final de 8 a 12 páginas
```

---

## BLOQUE 1: Acciones para el Hito Presencial (`docs/hito_parte2hp_equipo07.pdf`)

> **Nota de Rúbrica:** El documento debe tener estrictamente **entre 1 y 2 páginas**.

### ✔️ Acción 1.1: Corregir nombre del integrante en la portada (COMPLETADO)
- **Corrección aplicada en `hito_parte2hp_equipo07.pdf`:** Se corrigió el nombre bajo *Integrantes* a:
  ```
  Integrantes:
  Benjamín Zamora
  José Palma
  Franco Maripil
  Nicolas Portilla
  Thomas Márquez
  ```

---

###  Acción 1.2: Compactar el formato a máximo 2 páginas
- **Problema actual:** El PDF tiene 5 páginas debido a espaciados excesivos y saltos de página en Google Docs.
- **Solución:** Reducir márgenes, tamaño de fuente (10-11 pt) y espaciado de tablas para que todo quepa en **2 páginas**.

---

###  Acción 1.3: Pegar el contenido en la sección "Flujo Planificado"
*Copiar y pegar este texto directamente debajo del título que quedó en blanco:*

```markdown
1. Validación e Ingesta: Lectura de mediciones JSONL en entrada/ validando tipos, rangos físicos y formato de timestamp.
2. Procesamiento Concurrente: Distribución de archivos mediante cola compartida (queue.Queue) hacia 3 trabajadores (threading.Thread), protegiendo métricas y log de alertas con mutex (threading.Lock).
3. Salida de Reportes: Generación de informes por estación en salida/informe_*.txt y resumen en salida/resumen_ambiental.txt.
4. Gestión de Incidencias: 
   - Clasificación y traslado de informes según alertas: sin_alertas/ (0), con_alertas/ (1-3) y criticas/ (≥4), aplicando sufijos correlativos (_v1, _v2) para evitar sobrescrituras.
   - Respaldo del resumen general en gestion_ambiental/resumen_resguardado.txt.
   - Separación de alertas por indicador en gestion_ambiental/alertas_por_indicador/{temperatura,humedad,pm25,ruido}/.
   - Generación de gestion_ambiental/inventario_ambiental.json y bitácora trazable en logs/gestion_ambiental.log.
```

---

###  Acción 1.4: Agregar la sección "Cinco evidencias para el informe final"
*Copiar y pegar esta sección requerida por la pauta (actualmente omitida en el hito):*

```markdown
Cinco evidencias planificadas para el informe final:
1. Configuración de VM Debian 13: Capturas de VirtualBox (4 GB RAM, 4 vCPUs, 25 GB disco, red NAT) y primer login.
2. Preparación del Sistema: Salida de sudo apt update, upgrade y verificación de python3 --version (Python 3.13+).
3. Estructura del Gestor de Incidencias: Salida de ls -R en gestion_ambiental/ demostrando informes clasificados y alertas separadas.
4. Monitoreo de Procesos y Memoria: Captura de ps aux (PID, PPID, STAT, %CPU, RSS, VSZ) durante la concurrencia y salidas de free -h y df -h.
5. Metadatos y Tolerancia a Fallos: Salida de stat logs/gestion_ambiental.log y evidencia en la bitácora del control de una excepción sin detener el sistema.
```

---

###  Acción 1.5: Agregar riesgo del gestor de incidencias a la tabla de riesgos
*Añadir esta fila a la tabla de riesgos:*

| Riesgo Identificado | Medida de Mitigación |
| :--- | :--- |
| Sobrescritura accidental de reportes previos al reclasificar o procesar nuevas fechas en `gestion_ambiental/`. | Implementación de la función `obtener_nombre_seguro()`, que valida la existencia del archivo de destino y añade un sufijo correlativo automático (`_v1`, `_v2`, etc.) preservando intacto el historial. |

---

## BLOQUE 2: Acciones para el Informe Final Asincrónico (8 a 12 Páginas)

Franco ya recopiló en `evidencias/Parte 2 solemne Sistemas Operativos.pdf` las capturas de:
- Descarga de ISO y VirtualBox.
- Asignación de recursos (4 GB RAM, 4 CPUs, 25 GB).
- Instalación de Debian y pantalla de login.
- `apt update`, `apt upgrade` y versión de Python.
- Estructura de carpetas con `mkdir -p` y `ls -R`.

###  Acción 2.1: Tomar las 4 capturas faltantes en la VM Debian (Pauta punto e - 8 pts)

Ejecutar estos comandos en la terminal de Debian y guardar las capturas:

1. **Monitoreo de procesos concurrentes:**
   ```bash
   # En una terminal ejecutar el procesador y capturarlo inmediatamente con ps:
   python3 src/concurrente.py & ps aux | grep concurrente.py
   ```
   *Debe permitir identificar:* **PID**, **PPID**, columna **STAT** (ej: `S` o `R`), **%CPU**, **%MEM**, **RSS** (memoria física residente) y **VSZ** (memoria virtual).

2. **Recursos del sistema:**
   ```bash
   free -h
   df -h
   du -sh ~/Documents/Solemne1-main
   ```

3. **Metadatos de la bitácora:**
   ```bash
   stat logs/gestion_ambiental.log
   ```
   *Muestra inodos, tamaño, permisos, propietario (`equipo_02`), grupo y marcas de tiempo.*

4. **Evidencia de tolerancia a fallos / error controlado:**
   - Provocar un error leve (por ejemplo, agregar una línea con texto cualquiera o un indicador no reconocido en `alertas/alertas_detectadas.log`, o ejecutar el gestor cuando falta un archivo).
   - Mostrar cómo en `logs/gestion_ambiental.log` queda registrado:
     `[FECHA] Alerta con indicador desconocido...` o `[FECHA] ADVERTENCIA...`
   - Demostrar que el gestor **no se cayó** y completó el inventario exitosamente.

---

###  Acción 2.2: Redacción teórica obligatoria requerida por la pauta

Incluir estas dos explicaciones técnicas en el cuerpo del informe:

#### 1. Diferencia entre Mover y Copiar archivos en GNU/Linux
> - **Mover (`shutil.move` / `mv`):** Se utiliza para los informes individuales (`informe_*.txt`) porque representa una **transferencia de estado y reclasificación definitiva**. A nivel de sistema de archivos en la misma partición, mover solo actualiza las entradas del directorio (punteros a inodos) sin duplicar datos en disco, evitando que un informe pendiente permanezca en la bandeja de salida.
> - **Copiar (`shutil.copy` / `cp`):** Se utiliza para `resumen_ambiental.txt` hacia `resumen_resguardado.txt` porque se requiere un **mecanismo de respaldo y auditoría (backup)**. El archivo original debe permanecer en `salida/` para consultas directas del módulo de monitoreo, mientras que la copia resguardada queda protegida bajo la administración del área ambiental.

#### 2. Interpretación de los parámetros de `ps aux`
> - **PID (Process ID):** Identificador unívoco asignado por el kernel al proceso concurrente.
> - **PPID (Parent Process ID):** PID del proceso padre (habitualmente la shell Bash o terminal).
> - **STAT:** Estado del proceso (`R` = Running/ejecutándose en CPU; `S` = Interruptible Sleep esperando I/O o locks; `l` = multihilo).
> - **%CPU y %MEM:** Porcentaje de tiempo de procesador y de memoria RAM ocupados.
> - **VSZ (Virtual Memory Size):** Total de memoria virtual asignada en KiB (código, librerías compartidas, buffers).
> - **RSS (Resident Set Size):** Memoria física real cargada en RAM en KiB (excluye espacio swap y páginas no cargadas).

---

###  Acción 2.3: Ensamblar el Informe Final (Estructura de 8 a 12 páginas)

| Sección | Contenido | Insumo a utilizar |
| :--- | :--- | :--- |
| **Pág. 1** | Portada formal institucional | Datos del equipo (NRC 18897). |
| **Pág. 2-3** | Instalación de Debian 13 en VirtualBox | Capturas de `evidencias/Parte 2 solemne...` (págs 1 y 2). |
| **Pág. 4** | Preparación del entorno | Captura de `apt update` y Python (pág 4 de evidencias). |
| **Pág. 5-6** | Gestor de Incidencias y clasificación | Capturas de `mkdir -p`, `ls -R` e `inventario_ambiental.json`. |
| **Pág. 7-8** | Observación de Procesos y Sistema | Nuevas capturas de `ps aux`, `free -h`, `df -h` y `stat`. |
| **Pág. 9** | Mover vs Copiar y Manejo de Errores | Redacción teórica y captura del log con error controlado. |
| **Pág. 10** | Conclusiones y Trazabilidad | Síntesis técnica de la solución ambiental. |
