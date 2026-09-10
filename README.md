# Solemne 1 - Monitor de Estaciones Ambientales

Procesa archivos JSONL con mediciones de estaciones ambientales (temperatura, humedad, PM2.5 y
ruido), valida cada medición, detecta alertas por umbral y genera informes por estación más un
resumen consolidado. Después un gestor de incidencias clasifica esos informes y separa las alertas
por indicador.

Existen dos versiones equivalentes del mismo procesamiento:

- **Secuencial** (`src/secuencial.py`): procesa un archivo completo antes de pasar al siguiente,
  sin hilos ni procesos.
- **Concurrente** (`src/concurrente.py`): reparte los archivos entre 3 trabajadores
  (`threading.Thread`) mediante una `queue.Queue`, con `threading.Lock` sobre los recursos
  compartidos.

Ambas producen métricas idénticas sobre el mismo conjunto de entrada.

**Equipo 07** — Sección NRC 18897 — Benjamín Zamora, José Palma, Franco Maripil, Nicolás Portilla,
Thomas Márquez.

El informe final de la Parte 2 es `docs/Informe_Final_Parte2_Equipo07.pdf` (fuente editable en
`.docx`), generado por `scripts/generar_informe.py`.

# Requisitos

Python 3, **solo biblioteca estándar**: sin `pip`, sin entorno virtual, sin compilación.
Verificado en Debian 13.6 con Python 3.13.5 y en Windows con 3.14.

> En Linux/macOS usar `python3`; en Windows puede usarse `py`.

La máquina virtual de la Parte 2 corre sobre Hyper-V (generación 1), Debian 13.6 de 64 bits,
4 GB de RAM, 4 procesadores virtuales, disco dinámico de 25 GB y red NAT. El hito presencial
declaró VirtualBox; el cambio de hipervisor se documenta en el informe y los recursos
comprometidos se respetaron.

# Ejecución

Ejecutar siempre desde la raíz del proyecto, en este orden.

## 0. Limpieza previa (obligatoria para reproducir)

```bash
python3 scripts/limpiar.py
```

El proyecto se entrega **con la salida de la corrida ya hecha**, porque la pauta exige los informes
clasificados como evidencia. Si se reejecuta sin limpiar, el gestor no sobrescribe (punto 6 de la
pauta) y crea 20 duplicados `_v1`, con lo que el resultado pasa de 20 informes a 40.
`limpiar.py` borra solo lo regenerable y conserva `entrada/`, `salida/LEEME.txt` y `evidencias/`.

## 1. Generar datos de entrada

```bash
python3 scripts/generar_archivos_entrada.py
```

## 2. Secuencial

```bash
python3 src/secuencial.py
```

## 3. Concurrente

```bash
python3 src/concurrente.py
```

## 4. Gestor de incidencias

Clasifica los informes generados y organiza las alertas por indicador:

```bash
python3 src/gestor_incidencias.py
```

> El gestor **mueve** los informes de `salida/` a sus carpetas de clasificación, tal como exige la
> pauta. Por eso `salida/` queda solo con el resumen. Los 20 informes tal como los dejó la Parte 1
> están respaldados en `evidencias/salida_parte1/` (ver `salida/LEEME.txt`).

Los procesadores devuelven `0` si todo se procesó y `1` si algún archivo falló. El gestor siempre
termina en `0` y reporta las anomalías en la bitácora y en el inventario.

# Estructura de carpetas

```
├───alertas/                    → log de alertas detectadas
│   └─── alertas_detectadas.log
├───docs/                       → informe final, pautas y hito presencial
├───entrada/                    → 20 archivos JSONL con las mediciones
├───evidencias/                 → comparación, salidas de comandos en Debian y capturas
│   ├───debian/                 → salida real de los comandos ejecutados en la VM
│   ├───fotos/                  → capturas de la instalación y de la consola
│   ├───hyperv/                 → configuración de la VM y SHA256 del ISO
│   └───salida_parte1/          → los 20 informes antes de que el gestor los moviera
├───gestion_ambiental/          → salida del gestor de incidencias
│   ├───alertas_por_indicador/  → alertas separadas por tipo
│   ├───con_alertas/            → informes con 1 a 3 alertas
│   ├───criticas/               → informes con 4 o más alertas
│   ├───sin_alertas/            → informes sin alertas
│   ├───inventario_ambiental.json
│   └───resumen_resguardado.txt
├───logs/                       → bitácora de gestión de incidencias
├───salida/                     → resumen consolidado
├───scripts/                    → generador de entrada, limpieza y generador del informe
└───src/                        → código fuente de los procesadores y del gestor
```

# Comparación

| Característica | Secuencial | Concurrente |
|----------------|------------|-------------|
| Archivos procesados | 20 | 20 |
| Mediciones válidas | 300 | 300 |
| Mediciones inválidas | 20 | 20 |
| Alertas totales | 79 | 79 |
| PM2.5 máximo global | 177.87 ug/m3 | 177.87 ug/m3 |
| Ruido máximo global | 109.77 dB | 109.77 dB |
| Tiempo de ejecución | 0.005 s | 0.007 s |
| Cantidad de trabajadores | 1 | 3 |

Las seis métricas que la pauta exige comparar coinciden, y los 20 informes individuales y
`alertas_detectadas.log` son byte a byte idénticos entre ambas versiones. El tiempo es la única
métrica que fluctúa entre corridas.

La versión concurrente resulta **ligeramente más lenta**: con 20 archivos de 16 líneas el costo de
crear hilos y coordinar la cola y los mutex domina sobre el trabajo útil, y el GIL limita el
paralelismo de código Python puro. Con una carga amplificada de 4000 archivos sí se midió más de
100 % de CPU con 4 hilos vivos. El análisis completo está en `evidencias/comparacion_parte1.txt`.

# Resultado del gestor

```
# 0 / 8 / 12 informes   (sin_alertas / con_alertas / criticas)
```

`sin_alertas/` queda vacía de forma legítima: la pauta exige al menos una alerta por archivo de
entrada, así que ningún informe llega con cero. La rama está implementada y se demuestra en
`evidencias/debian/06-control-de-errores.txt`.

Las 79 alertas se reparten así, y la suma cuadra con `alertas_detectadas.log`, con el resumen
consolidado y con el inventario:

| Indicador | Alertas |
|-----------|---------|
| Temperatura | 21 |
| Humedad | 17 |
| PM2.5 | 20 |
| Ruido | 21 |

# Formato JSONL

Cada línea del archivo de entrada es un objeto JSON con exactamente estos campos:

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `timestamp` | string | Fecha y hora en formato `AAAA-MM-DDTHH:MM:SS` |
| `estacion` | string | Código de estación. Ej: `STG01`, `VAL03` |
| `temperatura` | float | Temperatura en °C |
| `humedad` | float | Humedad en % |
| `pm25` | float | Concentración de material particulado en ug/m3 |
| `ruido_db` | float | Nivel de ruido en dB |

Ejemplo de línea:

```json
{"timestamp": "2024-01-01T00:00:00", "estacion": "STG01", "temperatura": 24.19, "humedad": 51.62, "pm25": 12.9, "ruido_db": 47.05}
```

Los archivos se nombran `estacion_CODIGO_AAAAMMDD.jsonl` (ej: `estacion_STG01_20240101.jsonl`).
Se usan las estaciones `STG01`–`STG04`, `VAL01`–`VAL04`, `CON01`–`CON04` y `PUN01`–`PUN04`,
20 archivos en total con una segunda fecha para cuatro estaciones.

# Reglas de validación y alerta

## Validación de mediciones

Una medición es válida solo si cumple todas estas reglas:

- La línea es JSON bien formado y contiene los seis campos obligatorios.
- `timestamp` usa el formato `AAAA-MM-DDTHH:MM:SS`.
- `estacion` coincide con el código del nombre del archivo.
- Los cuatro campos numéricos son **números reales** de Python y **finitos**: se rechaza `bool`
  (aunque sea subclase de `int`), se rechazan cadenas como `"25.5"` y se rechazan `Infinity` y `NaN`.
- Rangos: `temperatura` entre `-20.0` y `60.0` °C; `humedad` entre `0.0` y `100.0` %;
  `pm25` mayor o igual a `0.0`; `ruido_db` entre `0.0` y `140.0` dB.

Los tipos se comprueban **antes** que los rangos. Las líneas en blanco se omiten por completo y no
cuentan como leídas. Los archivos se leen con `utf-8-sig` y `errors="replace"`, de modo que un BOM
o un byte dañado no detienen el proceso: la línea cae como inválida.

## Reglas de alerta

Se genera una alerta cuando la medición alcanza o supera el umbral:

| Indicador | Condición |
|-----------|-----------|
| Temperatura | ≥ 35.0 °C |
| Humedad | ≤ 20.0 % |
| PM2.5 | ≥ 35.0 ug/m3 |
| Ruido | ≥ 75.0 dB |

Una misma medición puede activar varias alertas. Cada una se registra en
`alertas/alertas_detectadas.log` en formato separado por `;`:
`archivo;estacion;timestamp;indicador;valor;umbral`

# Evidencia

| Qué | Dónde |
|-----|-------|
| Comparación exigida por la Parte 1 | `evidencias/comparacion_parte1.txt` |
| Salidas reales de comandos en Debian | `evidencias/debian/01` a `07` |
| Configuración de la VM y SHA256 del ISO | `evidencias/hyperv/configuracion-vm-hyperv.txt` |
| Capturas de la instalación y de la consola | `evidencias/fotos/` |
| Los 20 informes de la Parte 1 | `evidencias/salida_parte1/` |
