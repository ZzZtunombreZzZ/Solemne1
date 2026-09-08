# Solemne 1 - Procesador de Mediciones meteorologicas

Los programas realizan un procesamiento de archivos de mediciones meteorológicas, generando un archivo de salida con los resultados del procesamiento.
donde existen 2 tipos de procesamientos

- Secuencial: Procesa los archivos de mediciones meteorológicas de manera secuencial, es decir, uno después del otro.
- Concurrente: Procesa los archivos de mediciones meteorológicas de manera concurrente, es decir, varios archivos al mismo tiempo.

# Requisitos
Python: 3.14

Los programas usan exclusivamente la librería estándar de Python, no requieren dependencias externas ni compilación.

# Ejecución

Ejecutar siempre desde la raíz del proyecto.

## Generar datos de entrada
```bash
py scripts/generar_archivos_entrada.py
```

## Secuencial
para ejecutar el programa secuencial, se debe ejecutar el siguiente comando en la terminal:

```bash
py src/secuencial.py
```

## Concurrente
para ejecutar el programa concurrente, se debe ejecutar el siguiente comando en la terminal:

```bash
py src/concurrente.py
```

## Gestor de incidencias
Clasifica los informes generados y organiza las alertas por indicador:

```bash
py src/gestor_incidencias.py
```

# Estructura de carpetas
```
├───alertas/                    → log de alertas detectadas
│   └─── alertas_detectadas.log
├───docs/                       → documentación de la práctica
├───entrada/                    → archivos JSONL con las mediciones
├───evidencias/                 → capturas del procesamiento
├───gestion_ambiental/          → salida del gestor de incidencias
│   ├───alertas_por_indicador/  → alertas separadas por tipo
│   ├───con_alertas/            → informes con 1 a 3 alertas
│   ├───criticas/               → informes con más de 3 alertas
│   ├───sin_alertas/            → informes sin alertas
│   ├───inventario_ambiental.json
│   └───resumen_resguardado.txt
├───logs/                       → bitácora de gestión de incidencias
├───salida/                     → informes y resumen de cada procesamiento
├───scripts/                    → generador de archivos de entrada
└───src/                        → código fuente de los procesadores
```
# Comparacion 
 
| Caracteristica | Secuencial | Concurrente |
|----------------|------------|-------------|
| Archivos procesados | 20 | 20 |
| Mediciones validas | 300 | 300 |
| Mediciones invalidas | 20 | 20 |
| Alertas detectadas | 562 | 562 |
| PM2.5 maximo global | 499.95 ug/m3 | 499.95 ug/m3 |
| Ruido maximo global | 139.55 dB | 139.55 dB |
| Tiempo de ejecucion | 0.020 s | 0.025 s |
| Cantidad de trabajadores | 1 | 3 |

# formato JSONL 

Cada línea del archivo de entrada es un objeto JSON con exactamente los siguientes campos:

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
{"timestamp": "2024-01-01T08:15:00", "estacion": "STG01", "temperatura": 25.50, "humedad": 45.00, "pm25": 12.80, "ruido_db": 40.30}
```

Los archivos se nombran como `estacion_CODIGO_AAAAMMDD.jsonl` (ej: `estacion_STG01_20240101.jsonl`). Se admiten las estaciones `STG01`–`STG04`, `VAL01`–`VAL04`, `CON01`–`CON04` y `PUN01`–`PUN04`. El conjunto generado contiene 20 archivos, incluyendo una segunda fecha para cuatro estaciones.

# reglas de alerta/validación.

## Validación de mediciones

Una medición es válida solo si cumple todas estas reglas:

- Contiene los campos obligatorios: `timestamp`, `estacion`, `temperatura`, `humedad`, `pm25`, `ruido_db`.
- La línea es JSON bien formado.
- `timestamp` usa el formato `AAAA-MM-DDTHH:MM:SS`.
- `estacion` coincide con el código incluido en el nombre del archivo.
- `temperatura` está entre `-20.0` y `60.0` °C.
- `humedad` está entre `0.0` y `100.0` %.
- `pm25` es mayor o igual a `0.0` ug/m3.
- `ruido_db` está entre `0.0` y `140.0` dB.

Si alguna regla no se cumple, la línea se cuenta como medición inválida.

## Reglas de alerta

Se genera una alerta cuando la medición supera (o iguala) el umbral del indicador:

| Indicador | Condición |
|-----------|-----------|
| Temperatura | ≥ 35.0 °C |
| Humedad    | ≤ 20.0 %  |
| PM2.5      | ≥ 35.0 ug/m3 |
| Ruido      | ≥ 75.0 dB |

Cada alerta se registra en `alertas/alertas_detectadas.log` en formato separado por `;`:
`archivo;estacion;timestamp;indicador;valor;umbral`
