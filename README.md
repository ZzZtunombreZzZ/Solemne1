# Solemne 1 — Procesador de Mediciones Ambientales

Sistema de procesamiento de mediciones ambientales (temperatura, humedad, PM2.5 y ruido) que lee
archivos **JSONL** por estación, valida cada medición, detecta alertas por umbral, genera informes
individuales y un resumen consolidado, y finalmente clasifica esos informes mediante un gestor de
incidencias.

El proyecto implementa **dos versiones equivalentes** del mismo contrato de procesamiento:

- **Secuencial** (`src/secuencial.py`): procesa un archivo completo antes de pasar al siguiente.
  Sin hilos, procesos ni pools.
- **Concurrente** (`src/concurrente.py`): reparte los archivos entre 3 trabajadores
  (`threading.Thread`) mediante una `queue.Queue`, protegiendo los acumuladores globales y la
  escritura de los logs con `threading.Lock`.

Ambas versiones producen **métricas idénticas** sobre el mismo conjunto de entrada.

## Qué se entrega

El **entregable central de la Parte 2 es el informe final**:
`docs/Informe_Final_Parte2_Equipo07.pdf`, con su fuente editable
`docs/Informe_Final_Parte2_Equipo07.docx`. Ambos se producen desde una única fuente de contenido
con `scripts/generar_informe.py`, que **no escribe a mano ninguna cifra**: extrae las métricas, los
tiempos, los conteos y los tamaños directamente de los archivos de `evidencias/debian/` en el
momento de generar el documento.

Este README es la **guía de reproducción** del código y el mapa de la evidencia: explica cómo
volver a obtener exactamente los mismos números y dónde está cada archivo que los respalda.

## Identificación

| Dato | Valor |
|------|-------|
| Evaluación | Solemne 1 — Práctica, Parte 1 y Parte 2 (Forma B) |
| Asignatura | Sistemas Operativos |
| Sección | NRC 18897 |
| Equipo | Equipo 07 |
| Lenguaje | Python |
| Integrantes | Benjamín Zamora, José Palma, Franco Maripil, Nicolás Portilla, Thomas Márquez |

---

## 1. Requisitos

- **Python 3** y **exclusivamente su biblioteca estándar**.
- **Sin `pip`**, **sin entorno virtual**, **sin dependencias externas**, **sin compilación**.
  El proyecto se clona y se ejecuta directamente.

Módulos estándar utilizados: `json`, `math`, `re`, `sys`, `time`, `threading`, `queue`, `shutil`,
`os`, `random`, `collections`, `datetime`, `pathlib`.

### Entornos verificados

| Entorno | Versión de Python | Intérprete |
|---------|-------------------|------------|
| Debian GNU/Linux 13 (trixie) 13.6, 64 bits, kernel 6.12.107+deb13-amd64 | 3.13.5 (`/usr/bin/python3`) | `python3` |
| Windows | 3.14 | `py` o `python` |

En Linux y macOS se invoca `python3`; en Windows puede usarse `py`.

La máquina virtual de la Parte 2 corre sobre **Microsoft Hyper-V**, generación 1 (arranque BIOS),
con 4 GB de RAM, 4 procesadores virtuales, disco dinámico de 25 GB y red NAT. La ISO utilizada fue
`debian-13.6.0-amd64-netinst.iso` (755 MB), con SHA256 verificado
`65273beed27b2df543b68b65630ba525cfbad8df2b12035732b2dff87d6664e7`, idéntico al publicado en
`SHA256SUMS` de `cdimage.debian.org`. La salida literal de los cmdlets de Hyper-V está en
`evidencias/hyperv/configuracion-vm-hyperv.txt`, y las capturas de la instalación en
`evidencias/fotos/instalacion-hyperv/`.

> **Nota de honestidad técnica:** el hito presencial declaró **VirtualBox** como hipervisor. La
> implementación final se realizó sobre **Hyper-V**. El cambio de hipervisor se documenta de forma
> explícita: los recursos comprometidos en el hito (4 GB de RAM, 4 vCPU, 25 GB de disco, red NAT)
> se respetaron exactamente.

---

## 2. Ejecución paso a paso

Todos los comandos se ejecutan **desde la raíz del proyecto** y **en este orden**. Los scripts
resuelven sus rutas contra la raíz del repositorio, no contra el directorio de trabajo actual, por
lo que también funcionan invocados desde otra ubicación.

### Paso 0 — Limpieza previa (obligatorio para reproducir)

```bash
python3 scripts/limpiar.py
```

> ### POR QUÉ HACE FALTA ESTE PASO
>
> **El árbol entregado viaja con la salida de la corrida ya documentada.** La pauta exige entregar
> los informes **ya clasificados** dentro de `gestion_ambiental/con_alertas/` y
> `gestion_ambiental/criticas/`, junto con el inventario, la bitácora y los logs por indicador: es
> la evidencia del trabajo. Por eso esos 20 informes vienen versionados en el repositorio.
>
> Eso choca con la reproducción. Si se ejecutan los pasos 1 a 5 sobre el árbol tal como se entrega,
> los programas vuelven a producir los 20 informes y el gestor los clasifica **encima** de los 20
> que ya venían. Como el punto 6 de la pauta le prohíbe sobrescribir, el gestor hace lo correcto y
> crea 20 duplicados `_v1`: el resultado pasa de 20 informes a 40 y deja de coincidir con lo que
> declaran este README y el inventario.
>
> No es un defecto del gestor: es que la reproducción parte de un árbol que ya tiene la corrida
> hecha. `scripts/limpiar.py` borra **solo lo que los programas vuelven a generar** y deja intacto
> todo lo que es fuente o evidencia.

| Se borra (regenerable) | Se conserva (fuente o evidencia) |
|------------------------|----------------------------------|
| `salida/informe_*.txt`, `salida/resumen_ambiental.txt` | `entrada/` (los 20 `.jsonl`) |
| `alertas/alertas_detectadas.log` | `salida/LEEME.txt` |
| `gestion_ambiental/{sin_alertas,con_alertas,criticas}/*` | `gestion_ambiental/sin_alertas/.gitkeep` |
| `gestion_ambiental/alertas_por_indicador/*/*.log` | `evidencias/`, incluido `evidencias/salida_parte1/` |
| `gestion_ambiental/inventario_ambiental.json`, `resumen_resguardado.txt` | `docs/`, `src/`, `scripts/`, `README.md` |
| `logs/gestion_ambiental.log` | |

El script es **idempotente**, no falla si algo ya no existe y termina siempre en `0`.

Si se prefiere no usar el script, estos son los comandos equivalentes exactos.

Debian, Linux y macOS:

```bash
rm -f salida/informe_*.txt salida/resumen_ambiental.txt
rm -f alertas/alertas_detectadas.log
rm -f gestion_ambiental/sin_alertas/informe_*.txt
rm -f gestion_ambiental/con_alertas/informe_*.txt
rm -f gestion_ambiental/criticas/informe_*.txt
rm -f gestion_ambiental/alertas_por_indicador/*/*.log
rm -f gestion_ambiental/inventario_ambiental.json gestion_ambiental/resumen_resguardado.txt
rm -f logs/gestion_ambiental.log
```

Windows (PowerShell):

```powershell
Remove-Item salida\informe_*.txt, salida\resumen_ambiental.txt -Force -ErrorAction SilentlyContinue
Remove-Item alertas\alertas_detectadas.log -Force -ErrorAction SilentlyContinue
Remove-Item gestion_ambiental\sin_alertas\informe_*.txt -Force -ErrorAction SilentlyContinue
Remove-Item gestion_ambiental\con_alertas\informe_*.txt -Force -ErrorAction SilentlyContinue
Remove-Item gestion_ambiental\criticas\informe_*.txt -Force -ErrorAction SilentlyContinue
Remove-Item gestion_ambiental\alertas_por_indicador\*\*.log -Force -ErrorAction SilentlyContinue
Remove-Item gestion_ambiental\inventario_ambiental.json, gestion_ambiental\resumen_resguardado.txt -Force -ErrorAction SilentlyContinue
Remove-Item logs\gestion_ambiental.log -Force -ErrorAction SilentlyContinue
```

> **Ojo con `salida/`:** la limpieza manual borra por patrón (`informe_*.txt` y
> `resumen_ambiental.txt`) precisamente para **no** borrar `salida/LEEME.txt`, que sí es parte del
> entregable. Un `rm -f salida/*` se llevaría también esa nota.

Al terminar los pasos 1 a 5, el árbol vuelve al mismo estado en que se entregó. Se comprobó
archivo por archivo contra el árbol entregado: **las únicas diferencias son las tres que dependen
del reloj**, y ninguna afecta a una métrica.

| Diferencia esperada | Dónde | Por qué |
|---------------------|-------|---------|
| Marcas de tiempo `[AAAA-MM-DD HH:MM:SS]` | `logs/gestion_ambiental.log` | Hora de la corrida. Tienen ancho fijo, así que el archivo conserva sus 28 líneas y sus 2442 bytes. |
| Campo `marca_de_ejecucion` | `gestion_ambiental/inventario_ambiental.json` | Declarado como informativo y fuera del inventario reproducible. |
| Línea `Tiempo total de ejecución` | `salida/resumen_ambiental.txt` y sus dos copias | Es la **única** métrica no determinista del proyecto, ya declarada como tal en la sección 7.2. |

Todo lo demás —los 20 archivos de entrada, los 20 informes, las 79 alertas, los cuatro logs por
indicador y el resto del inventario— sale byte a byte igual.

Esto vale **en Debian y en Windows**. Todos los programas escriben sus artefactos con `newline="\n"`
explícito, de modo que Python no traduzca el salto de línea según la plataforma; `.gitattributes`
aplica la misma regla (`eol=lf`) al clonar. Comprobado ejecutando el ciclo completo en Windows y
comparando contra el árbol entregado: 45 de 45 archivos idénticos byte a byte.

### Paso 1 — Generar el conjunto de entrada

```bash
python3 scripts/generar_archivos_entrada.py
```

Genera los 20 archivos `entrada/estacion_CODIGO_AAAAMMDD.jsonl` con **semilla fija**
(`SEMILLA = 20240908`). La salida es idéntica entre corridas y entre sistemas operativos: este es
el punto de partida de la reproducibilidad.

### Paso 2 — Ejecutar la versión secuencial

```bash
python3 src/secuencial.py
```

Escribe `salida/informe_CODIGO_AAAAMMDD.txt` (20 informes), `salida/resumen_ambiental.txt` y
`alertas/alertas_detectadas.log`.

### Paso 3 — Ejecutar la versión concurrente

```bash
python3 src/concurrente.py
```

Sobrescribe las mismas salidas con los resultados de la versión concurrente. Las métricas deben
coincidir exactamente con las del paso 2.

> Ambas versiones aceptan un argumento opcional con un código de estación (por ejemplo
> `python3 src/secuencial.py STG01`) para procesar solo esa estación. El conjunto oficial de
> métricas se obtiene **sin** ese argumento.

### Paso 4 — Respaldar la evidencia de la Parte 1 (obligatorio antes del paso 5)

> ### ADVERTENCIA IMPORTANTE
>
> **El gestor de incidencias MUEVE (no copia) los informes individuales fuera de `salida/`.**
> `shutil.move()` los traslada a `gestion_ambiental/con_alertas/`, `gestion_ambiental/criticas/` o
> `gestion_ambiental/sin_alertas/` según su cantidad de alertas. Después de correr el gestor,
> `salida/` queda únicamente con `resumen_ambiental.txt`.
>
> La Parte 1 exige entregar los informes individuales. **Si se quiere conservar esa evidencia hay
> que respaldar `salida/` ANTES de ejecutar el gestor.**

Debian, Linux y macOS:

```bash
mkdir -p evidencias/salida_parte1
cp -a salida/informe_*.txt salida/resumen_ambiental.txt evidencias/salida_parte1/
```

Windows (PowerShell):

```powershell
New-Item -ItemType Directory -Force evidencias\salida_parte1 | Out-Null
Copy-Item salida\informe_*.txt, salida\resumen_ambiental.txt evidencias\salida_parte1\ -Force
```

> Se copia **por patrón** y no `salida/*`, para que `salida/LEEME.txt` no termine dentro del
> respaldo: ese respaldo debe contener exactamente los 21 archivos de la Parte 1.

El respaldo ya versionado en este repositorio está en `evidencias/salida_parte1/` (20 informes más
`resumen_ambiental.txt`: 21 archivos). Por eso, en el árbol entregado, `salida/` contiene solo
`resumen_ambiental.txt` y `salida/LEEME.txt`, que explica exactamente dónde quedaron los 20 informes
individuales y cómo regenerarlos.

### Paso 5 — Ejecutar el gestor de incidencias

```bash
python3 src/gestor_incidencias.py
```

Clasifica los informes por cantidad de alertas, reparte las alertas por indicador, resguarda el
resumen, escribe la bitácora `logs/gestion_ambiental.log` y genera
`gestion_ambiental/inventario_ambiental.json`.

El gestor es **idempotente** en lo que importa: puede ejecutarse las veces que sea necesario y la
clasificación (0/8/12), las alertas por indicador (21/17/20/21) y el inventario (20 informes, 79
alertas, 0 anomalías) quedan siempre iguales, sin duplicados `_v1`. Los logs por indicador se
reconstruyen en cada corrida y el inventario se construye leyendo el estado real de las carpetas de
destino, no solo lo movido en la corrida actual.

> **Matiz honesto sobre la bitácora.** `logs/gestion_ambiental.log` se trunca al inicio de cada
> corrida, así que refleja **esa** corrida y no el acumulado. Las 28 líneas y 2442 bytes que se
> documentan corresponden a la corrida canónica, la que sigue a un ciclo completo (limpiar →
> generar → procesar → gestionar). Si se ejecuta el gestor una segunda vez sobre un árbol ya
> gestionado, la bitácora queda más corta (8 líneas) porque no hay informes que mover: la línea
> `Informes movidos en esta corrida: 0` lo deja explícito. El resultado del procesamiento no cambia.

---

## 3. Estructura de carpetas

```
Solemne1/
├── alertas/
│   └── alertas_detectadas.log              → bitácora plana de todas las alertas
├── docs/
│   ├── Informe_Final_Parte2_Equipo07.pdf   → INFORME FINAL de la Parte 2 (entregable central)
│   ├── Informe_Final_Parte2_Equipo07.docx  → misma fuente del informe, en formato editable
│   ├── Solemne01PracticaParte01FormaB.pdf  → pauta oficial de la Parte 1
│   ├── Solemne01PracticaParte2FormaB.pdf   → pauta oficial de la Parte 2
│   └── hito_parte2hp_equipo07.pdf          → hito presencial ya entregado
├── entrada/
│   └── estacion_CODIGO_AAAAMMDD.jsonl      → 20 archivos de mediciones (JSONL)
├── evidencias/
│   ├── comparacion_parte1.txt              → comparación secuencial vs concurrente
│   ├── debian/
│   │   ├── 01-preparar-ambiente.txt        → apt update / upgrade y entorno Python
│   │   ├── 02-entorno.txt                  → distribución, kernel, CPU, memoria, disco, red
│   │   ├── 03-ejecucion-parte1.txt         → generador, ambas versiones y tabla de métricas
│   │   ├── 04-observacion-procesos.txt     → time -v, ps, ps -L, /proc, free, df, du
│   │   ├── 05-gestor-incidencias.txt       → find, ls -lah, inventario, stat, mover vs copiar
│   │   ├── 06-control-de-errores.txt       → 4 anomalías, rama sin_alertas, idempotencia
│   │   ├── 07-anti-sobrescritura.txt       → sufijos _v1 / _v2 sin pérdidas
│   │   ├── resumen_secuencial.txt          → resumen consolidado de la corrida secuencial
│   │   ├── resumen_concurrente.txt         → resumen consolidado de la corrida concurrente
│   │   ├── alertas_secuencial.log          → log de alertas de la corrida secuencial
│   │   └── alertas_concurrente.log         → log de alertas de la corrida concurrente
│   ├── fotos/
│   │   ├── debian-01..04-*.png             → 4 capturas de la consola de la VM definitiva
│   │   ├── instalacion-hyperv/01..05-*.png → 5 capturas de la instalación real en Hyper-V
│   │   └── recortes/                       → recortes de esas 9 capturas, usados como figuras
│   │                                         del informe final (más el marcador .version)
│   ├── hyperv/
│   │   └── configuracion-vm-hyperv.txt     → salida de los cmdlets Hyper-V y SHA256 de la ISO
│   └── salida_parte1/                      → respaldo de los 20 informes + resumen (Parte 1)
├── gestion_ambiental/
│   ├── alertas_por_indicador/
│   │   ├── humedad/alertas_humedad.log
│   │   ├── pm25/alertas_pm25.log
│   │   ├── ruido/alertas_ruido.log
│   │   └── temperatura/alertas_temperatura.log
│   ├── con_alertas/                        → informes con 1 a 3 alertas
│   ├── criticas/                           → informes con 4 o más alertas
│   ├── sin_alertas/                        → informes con 0 alertas
│   │   └── .gitkeep                        → marcador: Git no versiona carpetas vacías
│   ├── inventario_ambiental.json           → inventario consolidado del gestor
│   └── resumen_resguardado.txt             → copia de salida/resumen_ambiental.txt
├── logs/
│   └── gestion_ambiental.log               → bitácora del gestor (se reescribe en cada corrida)
├── salida/
│   ├── informe_CODIGO_AAAAMMDD.txt         → 20 informes (el gestor los MUEVE de aquí)
│   ├── resumen_ambiental.txt               → resumen consolidado
│   └── LEEME.txt                           → explica por qué aquí solo queda el resumen
├── scripts/
│   ├── generar_archivos_entrada.py         → generador reproducible del conjunto de entrada
│   ├── limpiar.py                          → deja el árbol en estado previo a la corrida
│   └── generar_informe.py                  → genera el informe final (.docx y .pdf) leyendo
│                                             las cifras desde evidencias/debian/
├── src/
│   ├── secuencial.py                       → versión secuencial
│   ├── concurrente.py                      → versión concurrente (3 hilos)
│   └── gestor_incidencias.py               → gestor de incidencias
└── README.md
```

---

## 4. Formato JSONL de entrada

Cada archivo de `entrada/` contiene una medición por línea. Cada línea es un objeto JSON
independiente: sin cabecera, sin comas de separación y sin arreglo envolvente.

### Nombre de archivo

`estacion_CODIGO_AAAAMMDD.jsonl`

El nombre se valida con un patrón anclado contra el `stem` del archivo
(`^estacion_(?P<codigo>[^_]+)_(?P<fecha>[^_]+)$`), que **prohíbe guiones bajos** dentro del código y
de la fecha. Los códigos usados son `STG01`–`STG04`, `VAL01`–`VAL04`, `CON01`–`CON04` y
`PUN01`–`PUN04`. El conjunto oficial son 20 archivos: las 16 estaciones con fecha `20240101` más
las cuatro estaciones de Santiago repetidas con fecha `20240102`.

### Campos

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `timestamp` | `string` | Fecha y hora en formato `AAAA-MM-DDTHH:MM:SS` |
| `estacion` | `string` | Código de la estación; debe coincidir con el del nombre del archivo |
| `temperatura` | `number` (int o float) | Temperatura en °C |
| `humedad` | `number` (int o float) | Humedad relativa en % |
| `pm25` | `number` (int o float) | Material particulado fino en ug/m3 |
| `ruido_db` | `number` (int o float) | Nivel de ruido en dB |

### Línea de ejemplo real

Primera línea de `entrada/estacion_STG01_20240101.jsonl`:

```json
{"timestamp": "2024-01-01T00:00:00", "estacion": "STG01", "temperatura": 15.76, "humedad": 51.24, "pm25": 22.33, "ruido_db": 35.9}
```

Cada archivo contiene **15 mediciones válidas** distribuidas en **8 horarios distintos** del día,
más **1 línea inválida** deliberada: 16 líneas por archivo, 320 líneas en total.

---

## 5. Reglas de validación

Este es el criterio **real implementado en el código** (`validar_medicion()`, idéntico en
`src/secuencial.py` y `src/concurrente.py`). Es más estricto que una simple validación por rangos.

### Lectura del archivo

- Los archivos se abren con **`encoding="utf-8-sig"`**, de modo que un BOM opcional se consume y no
  invalida la primera línea.
- Se abren con **`errors="replace"`**, de modo que un byte que no es UTF-8 válido no mata el proceso
  ni el hilo: la línea dañada queda con caracteres de reemplazo (U+FFFD) y cae de forma natural
  como medición inválida.
- **Las líneas en blanco se omiten por completo**: no se cuentan como leídas, ni como válidas, ni
  como inválidas. Por eso `Líneas leídas` puede ser menor que el número de líneas físicas del
  archivo.

### Condiciones que debe cumplir una medición para ser válida

1. La línea parsea como **JSON bien formado**. Los errores `json.JSONDecodeError`, `ValueError` y
   `RecursionError` se traducen en línea inválida, nunca en una excepción propagada.
2. El valor resultante es un **objeto JSON** (`dict`). Un arreglo, un número o una cadena suelta son
   inválidos.
3. Están presentes **los seis campos obligatorios**: `timestamp`, `estacion`, `temperatura`,
   `humedad`, `pm25`, `ruido_db`.
4. `estacion` **coincide exactamente** con el código derivado del nombre del archivo.
5. `timestamp` es una cadena que parsea con el formato **`%Y-%m-%dT%H:%M:%S`**.
6. Los **cuatro campos numéricos** (`temperatura`, `humedad`, `pm25`, `ruido_db`) son **números
   reales de Python**, verificado por `es_numero_real()`:
   - Deben ser `int` o `float`. **Una cadena se rechaza**: `"25.5"` es inválido. **No se aplica una
     conversión con `float()`.**
   - **`bool` se rechaza explícitamente**, aunque en Python sea subclase de `int`: `true` y `false`
     en el JSON son inválidos.
   - Deben ser **finitos** (`math.isfinite`). Se rechazan `Infinity`, `-Infinity` y `NaN`, porque
     contaminan promedios y máximos.
   - Un entero gigante que provoca `OverflowError` al evaluarse también se rechaza.
7. Los rangos físicos se comprueban **después** de la validación de tipos, porque comparar un
   `float` con una `str` lanzaría `TypeError`:

   | Campo | Rango válido |
   |-------|--------------|
   | `temperatura` | `-20.0 ≤ v ≤ 60.0` °C |
   | `humedad` | `0.0 ≤ v ≤ 100.0` % |
   | `pm25` | `v ≥ 0.0` ug/m3 |
   | `ruido_db` | `0.0 ≤ v ≤ 140.0` dB |

Si alguna condición falla, la línea se contabiliza como **medición inválida** y el procesamiento
continúa con la línea siguiente.

### Validación a nivel de archivo

- Un archivo cuyo nombre **no calza** con el patrón se omite con un aviso por `stderr` y se registra
  como archivo con error.
- Dos archivos de entrada distintos que resolverían al **mismo informe de salida** constituyen una
  **colisión**: gana el primero en orden alfabético y el segundo se registra como error, en lugar
  de sobrescribir el informe en silencio.
- Un archivo que falla por completo durante su lectura se avisa por `stderr`, **no** se cuenta como
  procesado y hace que el programa termine con código de salida `1`.

### Tipos de línea inválida presentes en el conjunto oficial

El generador inyecta **una línea inválida por archivo**, rotando entre cinco tipos (4 archivos por
tipo), de modo que la validación quede demostrada de verdad y no con veinte copias del mismo caso:

| Tipo | Qué prueba |
|------|-----------|
| `json_malformado` | Texto que no parsea como JSON |
| `atributo_ausente` | Falta un campo obligatorio |
| `estacion_inconsistente` | El campo `estacion` no coincide con el nombre del archivo |
| `valor_fuera_rango` | `pm25` negativo, fuera del dominio físico |
| `tipo_incorrecto` | `temperatura` entregada como cadena |

---

## 6. Reglas de alerta

Se genera **una alerta por cada indicador** que cruza su umbral, no una por medición: una sola
medición puede aportar **hasta cuatro alertas**.

| Indicador | Condición de alerta | Umbral |
|-----------|---------------------|--------|
| Temperatura | `temperatura >= 35.0` | 35.0 °C |
| Humedad | `humedad <= 20.0` | 20.0 % |
| PM2.5 | `pm25 >= 35.0` | 35.0 ug/m3 |
| Ruido | `ruido_db >= 75.0` | 75.0 dB |

Solo se evalúan alertas sobre mediciones **válidas**.

### Formato de línea del log de alertas

Cada alerta se anexa a `alertas/alertas_detectadas.log` en formato separado por `;`:

```
archivo;estacion;timestamp;indicador;valor;umbral
```

Ejemplo real, primera línea de `alertas/alertas_detectadas.log`:

```
estacion_CON01_20240101.jsonl;CON01;2024-01-01T15:31:49;Temperatura;45.98;35.0
```

Los nombres de indicador escritos en el log son `Temperatura`, `Humedad`, `PM2.5` y `Ruido`.

En la versión concurrente las alertas se acumulan en buffers en memoria, uno por archivo de
entrada, y se vuelcan al log **una sola vez al final, en orden de nombre de archivo**. Así el log
resulta byte a byte idéntico al de la versión secuencial y estable entre corridas, pese a los tres
hilos.

---

## 7. Resultados esperados

Los siguientes valores son los que debe reproducir cualquier persona que clone el repositorio y
ejecute los pasos 1 a 5 en un Debian limpio.

### 7.1 Métricas del conjunto oficial de 20 archivos

Las métricas son **idénticas** en la versión secuencial y en la concurrente.

| Métrica | Secuencial | Concurrente |
|---------|-----------:|------------:|
| Archivos procesados | 20 | 20 |
| Líneas leídas | 320 | 320 |
| Mediciones válidas | 300 | 300 |
| Mediciones inválidas | 20 | 20 |
| Temperatura promedio global | 23.46 °C | 23.46 °C |
| Humedad promedio global | 54.85 % | 54.85 % |
| PM2.5 máximo global | 177.87 ug/m3 | 177.87 ug/m3 |
| Ruido máximo global | 109.77 dB | 109.77 dB |
| Alertas totales | 79 | 79 |
| Estación con mayor cantidad de alertas | STG04 | STG04 |
| Cantidad de trabajadores | 1 | 3 |

Los 20 informes individuales también coinciden uno a uno, y el log de alertas contiene las mismas
79 alertas en ambas versiones, sin pérdidas ni duplicados.

### 7.2 Tiempos medidos en Debian

| Medición | Secuencial | Concurrente |
|----------|-----------:|------------:|
| Tiempo real (`time`) | `0m0.019s` | `0m0.023s` |
| Tiempo interno reportado por el programa | 0.005 s | 0.007 s |
| Trabajadores | 1 | 3 |

> **La versión concurrente es ligeramente MÁS LENTA que la secuencial en este conjunto.** Es un
> resultado real y se documenta sin maquillarlo. Con 20 archivos pequeños (16 líneas cada uno, 320
> líneas en total) el trabajo útil se completa en milisegundos, mientras que el costo fijo de crear
> los hilos, alimentar la `queue.Queue` y adquirir y liberar los mutex que protegen los
> acumuladores globales domina el tiempo total. La concurrencia solo compensa ese costo fijo cuando
> el trabajo por archivo es sustancialmente mayor; su ventaja se observa recién con carga
> amplificada, no con el conjunto que exige la pauta.
>
> El tiempo es la **única** métrica que fluctúa entre corridas; las once métricas de la tabla 7.1
> son deterministas.
>
> Estos cuatro valores salen de una sola corrida, la que quedó congelada como evidencia: los tiempos
> reales están en `evidencias/debian/03-ejecucion-parte1.txt` y los tiempos internos en
> `evidencias/debian/resumen_secuencial.txt` y `salida/resumen_ambiental.txt` (respaldado también en
> `evidencias/debian/resumen_concurrente.txt`).

### 7.3 Recursos de la corrida oficial

Medidos con `/usr/bin/time -v python3 src/concurrente.py` sobre los 20 archivos:

| Recurso | Valor |
|---------|-------|
| User time | 0.01 s |
| System time | 0.00 s |
| Percent of CPU | 100 % |
| Elapsed (wall clock) | 0:00.02 |
| Maximum resident set size | 12 964 KB |
| Voluntary context switches | 149 |
| File system inputs | 0 |
| File system outputs | 184 |

Tamaño del proyecto en disco al momento de la medición: `du -sh` entregó **624K**. Espacio del
sistema: `/dev/sda1` con 24 G totales, 1.2 G usados y 21 G disponibles (6 % de uso).

> **Sobre la observación en vivo del proceso concurrente:** se realizó sobre una **carga amplificada
> de 4000 archivos (64 000 líneas)** (los 20 oficiales replicados con otras fechas válidas), en una
> **copia** del proyecto, porque el conjunto oficial se procesa en milisegundos y `ps` no alcanza a
> tomar una muestra útil. Se tomaron 270 muestras de `ps` con el proceso vivo. **Todas las métricas
> entregables del proyecto provienen del conjunto oficial de 20 archivos.** El detalle completo
> (PID, PPID, `STAT`, `%CPU`, `NLWP` = 4 hilos, `ps -L` y `/proc/<PID>/status`) está en
> `evidencias/debian/04-observacion-procesos.txt`.
>
> **Precisión sobre las muestras de `%CPU` transcritas.** No todas superan el 100 %, y no tendrían
> por qué: `%CPU` en `ps` es un **promedio acumulado** desde que arrancó el proceso, así que sube a
> medida que los hilos avanzan. Las muestras transcritas en
> `evidencias/debian/04-observacion-procesos.txt` van de **0.0 % a 103 %**: la primera marca
> `0.0 %` con `NLWP` = 1, tomada **antes** de que arrancaran los tres trabajadores, y hay otra de
> `75.0 %` tomada ya con los cuatro hilos vivos. **Lo que prueba el trabajo simultáneo en varios
> núcleos son las muestras que pasan del 100 %: 100 %, 102 % y 103 %**, imposibles para un proceso
> que solo pudiera ocupar un núcleo.

### 7.4 Clasificación del gestor de incidencias

Criterio de clasificación (`clasificar()`):

| Alertas del informe | Carpeta destino |
|---------------------|-----------------|
| 0 | `gestion_ambiental/sin_alertas/` |
| 1 a 3 | `gestion_ambiental/con_alertas/` |
| 4 o más | `gestion_ambiental/criticas/` |

Resultado sobre el conjunto oficial:

| Categoría | Informes |
|-----------|---------:|
| `sin_alertas` | 0 |
| `con_alertas` | 8 |
| `criticas` | 12 |
| **Total** | **20** |

Alertas repartidas por indicador:

| Indicador | Alertas |
|-----------|--------:|
| Temperatura | 21 |
| Humedad | 17 |
| PM2.5 | 20 |
| Ruido | 21 |
| **Total** | **79** |

**Cuadratura exacta:** `21 + 17 + 20 + 21 = 79`, y ese 79 es el mismo número que aparece en
`alertas/alertas_detectadas.log`, en `salida/resumen_ambiental.txt` y en
`gestion_ambiental/inventario_ambiental.json`. Las cuatro fuentes coinciden.

El inventario de la corrida entregada declara además **0 anomalías registradas** y **0 errores de
proceso**, y la bitácora `logs/gestion_ambiental.log` (28 líneas, 2442 bytes) cierra con la línea
`CONTROL DE ERRORES: no se detectaron anomalías en esta corrida.`

> **La bitácora es por corrida, no acumulativa.** El gestor la **reescribe completa** cada vez que
> se ejecuta: la primera línea de la corrida trunca el archivo y las siguientes se agregan. Por eso
> esas 28 líneas y esos 2442 bytes son un valor **reproducible** y no el resultado de haber corrido
> el gestor exactamente una vez. Solo cambian las marcas de tiempo, que tienen ancho fijo y no
> alteran el tamaño. La alternativa —abrir siempre en modo *append*— hacía crecer el archivo y
> volvía irreproducible el dato que la pauta manda inspeccionar con
> `stat logs/gestion_ambiental.log`; la otra alternativa —rotar a `.log.1`, `.log.2`, …— ensuciaba
> `logs/` con históricos cuya cantidad depende de cuántas veces se ejecutó el gestor. Ambas se
> descartaron por escrito en `src/gestor_incidencias.py`.
>
> Truncar **no** debilita el requisito 13 (controlar al menos un error sin detener el
> procesamiento): esa evidencia se produce **dentro de una sola corrida**, porque cada anomalía se
> registra línea por línea, el proceso continúa después de cada una y la corrida cierra siempre con
> la línea `CONTROL DE ERRORES:`. La evidencia de las cuatro anomalías inyectadas
> (`evidencias/debian/06-control-de-errores.txt`) es exactamente la bitácora de una única corrida.

### 7.5 Por qué `sin_alertas/` queda vacía

`gestion_ambiental/sin_alertas/` queda **legítimamente vacía**, y esto no es un defecto:

- La pauta exige que **todo archivo de entrada** contenga **al menos una medición que dispare
  alerta**.
- Por lo tanto, con datos que cumplen la pauta, **ningún informe puede terminar con 0 alertas**.
- No se generan archivos sin alertas para forzar esa rama, porque hacerlo **incumpliría la pauta**.
- La carpeta se crea de todas formas, y su existencia vacía es parte de la evidencia: demuestra que
  la rama existe y que ningún caso la activó. Como **Git no versiona carpetas vacías**, la carpeta
  viaja con un marcador `gestion_ambiental/sin_alertas/.gitkeep`: sin él, la carpeta que exige el
  árbol de la pauta simplemente no existiría en el repositorio entregado. El marcador **no** se
  cuenta como informe en el inventario, que sigue declarando `sin_alertas: 0`.
- La rama **está implementada y se demostró funcionando** con un informe de control de 0 alertas,
  documentado en `evidencias/debian/06-control-de-errores.txt`.

### 7.6 Robustez verificada del gestor

| Comprobación | Resultado |
|--------------|-----------|
| **Control de errores** | Se inyectaron 4 fallas: línea sin separadores, línea con campos insuficientes, indicador desconocido `Radiacion` e informe sin la línea de alertas. El gestor registró las 4 anomalías, **no se detuvo**, conservó las 79 alertas y generó el inventario igual. |
| **Anti-sobrescritura** | Partiendo de 21 informes acumulados, dos ciclos adicionales de procesamiento y gestión llevaron el total a 41 y luego a 61 (20 y 40 de ellos con sufijo de versión `_v1` / `_v2`). Comprobación: 21 + 20 + 20 = 61 esperados, 61 encontrados. **Cero archivos perdidos.** |
| **Idempotencia** | Tres corridas seguidas del gestor entregan siempre 79 alertas e inventario con 79. |

> **Por qué el árbol entregado muestra 0 anomalías.** Las tres comprobaciones de la tabla son
> **destructivas**: inyectan líneas corruptas, agregan un informe de control y vuelven a procesar
> varias veces. Se ejecutan sobre una **copia del proyecto** (`/home/equipo07/demo_destructiva`), no
> sobre el árbol entregable, precisamente para que la evidencia congelada corresponda a **una sola
> corrida limpia**. Por eso `gestion_ambiental/inventario_ambiental.json` del entregable reporta
> `anomalias_registradas: 0` y `errores_de_proceso: 0`, y `gestion_ambiental/sin_alertas/` está
> vacía, mientras que `evidencias/debian/06-control-de-errores.txt` y
> `evidencias/debian/07-anti-sobrescritura.txt` muestran las 4 anomalías, el informe de control en
> `sin_alertas/` y los sufijos `_v1` / `_v2`. Las marcas de tiempo e inventarios de esos dos
> archivos son los de la copia, no los del árbol entregado.

---

## 8. Códigos de salida

| Programa | Código | Significado |
|----------|-------:|-------------|
| `src/secuencial.py` | `0` | Todos los archivos de entrada se procesaron correctamente |
| `src/secuencial.py` | `1` | Al menos un archivo falló: nombre no reconocido, colisión de informe o excepción durante su procesamiento |
| `src/concurrente.py` | `0` | Todos los archivos de entrada se procesaron correctamente |
| `src/concurrente.py` | `1` | Al menos un archivo falló, con el mismo criterio que la versión secuencial |
| `src/gestor_incidencias.py` | `0` | **Siempre**, por diseño |

En ambos procesadores un archivo fallido **nunca se pierde en silencio**: se avisa por `stderr`, no
se cuenta como procesado y el resumen consolidado agrega las líneas `Archivos con error: N` y
`ADVERTENCIA: corrida incompleta, archivos no procesados: ...`. El resumen jamás declara una corrida
limpia si hubo fallas.

El **gestor de incidencias termina siempre en `0` por diseño**: su contrato es que ningún error
detiene el procesamiento. Toda anomalía se captura, se registra en la bitácora
`logs/gestion_ambiental.log` y se publica en `gestion_ambiental/inventario_ambiental.json` (bloque
`anomalias`, con su conteo por tipo y ejemplos). La bitácora cierra siempre con una línea
`CONTROL DE ERRORES:` que declara cuántas anomalías hubo, o bien que la corrida terminó limpia. Para
saber si el gestor encontró problemas hay que leer la bitácora y el inventario, no el código de
salida.

---

## 9. Dónde está cada evidencia

### Salidas del programa

| Evidencia | Ubicación |
|-----------|-----------|
| Archivos de entrada JSONL | `entrada/` |
| Informes individuales, respaldo de la Parte 1 (21 archivos) | `evidencias/salida_parte1/` |
| Resumen consolidado | `salida/resumen_ambiental.txt` |
| Explicación de por qué `salida/` solo conserva el resumen | `salida/LEEME.txt` |
| Log plano de las 79 alertas | `alertas/alertas_detectadas.log` |
| Informes clasificados | `gestion_ambiental/con_alertas/`, `gestion_ambiental/criticas/`, `gestion_ambiental/sin_alertas/` |
| Alertas separadas por indicador | `gestion_ambiental/alertas_por_indicador/{temperatura,humedad,pm25,ruido}/` |
| Inventario consolidado del gestor | `gestion_ambiental/inventario_ambiental.json` |
| Copia resguardada del resumen | `gestion_ambiental/resumen_resguardado.txt` |
| Bitácora del gestor | `logs/gestion_ambiental.log` |

### Evidencia de ejecución en Debian

| Punto de la pauta | Archivo |
|-------------------|---------|
| Preparación del ambiente: `apt update` / `upgrade`, entorno Python | `evidencias/debian/01-preparar-ambiente.txt` |
| Entorno: distribución, kernel, CPU, memoria, disco, red | `evidencias/debian/02-entorno.txt` |
| Ejecución de la Parte 1: generador, ambas versiones y tabla de las 6 métricas exigidas | `evidencias/debian/03-ejecucion-parte1.txt` |
| Observación de procesos y sistema de archivos: `time -v`, `ps`, `ps -L`, `/proc`, `free`, `df`, `du` | `evidencias/debian/04-observacion-procesos.txt` |
| Gestor de incidencias: `find`, `ls -lah`, inventario, `stat`, mover vs copiar | `evidencias/debian/05-gestor-incidencias.txt` |
| Control de errores: 4 anomalías, rama `sin_alertas`, idempotencia | `evidencias/debian/06-control-de-errores.txt` |
| Anti-sobrescritura: sufijos `_v1` y `_v2` sin pérdidas | `evidencias/debian/07-anti-sobrescritura.txt` |
| Resúmenes y logs de alertas de cada versión, tal como quedaron en la corrida | `evidencias/debian/resumen_secuencial.txt`, `evidencias/debian/resumen_concurrente.txt`, `evidencias/debian/alertas_secuencial.log`, `evidencias/debian/alertas_concurrente.log` |
| Comparación secuencial vs concurrente | `evidencias/comparacion_parte1.txt` |
| Configuración real de la máquina virtual en Hyper-V y SHA256 de la ISO | `evidencias/hyperv/configuracion-vm-hyperv.txt` |

### Capturas de pantalla

| Captura | Archivo |
|---------|---------|
| Consistencia de la Parte 1 | `evidencias/fotos/debian-01-consistencia-parte1.png` |
| Observación de procesos | `evidencias/fotos/debian-02-observacion-procesos.png` |
| Estructura generada por el gestor | `evidencias/fotos/debian-03-estructura-gestor.png` |
| `stat` y control de errores | `evidencias/fotos/debian-04-stat-y-control-de-errores.png` |
| Instalación real de Debian 13 en Hyper-V (menú del instalador, línea de arranque, instalación del sistema base, instalación final, primer inicio) | `evidencias/fotos/instalacion-hyperv/01-menu-instalador-debian13.png` … `05-primer-inicio.png` |
| Recortes de esas 9 capturas, tal como los inserta el informe final | `evidencias/fotos/recortes/` |

Son **9 capturas en total**: 4 de la consola de la máquina definitiva y 5 de su instalación real en
Hyper-V. `evidencias/fotos/recortes/` no aporta capturas nuevas: contiene el recorte de cada una de
esas 9, generado por `scripts/generar_informe.py` para insertarlas como figuras del informe final
sin los márgenes negros de la consola. El archivo `evidencias/fotos/recortes/.version` es el
marcador de versión del algoritmo de recorte: si cambia, la caché se regenera.

> **`evidencias/fotos/debian-02-observacion-procesos.png` es de una corrida distinta.** Esa captura
> y la transcripción de `evidencias/debian/04-observacion-procesos.txt` documentan **la misma
> observación con la misma carga amplificada de 4000 archivos y 64 000 líneas, pero en dos
> ejecuciones separadas**. Por eso sus números **no coinciden uno a uno y no deberían hacerlo**: la
> foto muestra el PID 12199 y un `%CPU` máximo de 107 %, mientras que el archivo de texto muestra el
> PID 10043 y un máximo de 103 %. Lo que ambas prueban es lo mismo —4 hilos vivos (`NLWP` = 4) y
> `%CPU` por sobre 100 %—, y coinciden en lo que no depende de la corrida: 624K de proyecto en disco
> y `/dev/sda1` con 24 G, 1.2 G usados y 21 G disponibles. **Los valores que este README declara son
> siempre los del archivo de texto**, no los de la foto.

### Documentación

| Documento | Archivo |
|-----------|---------|
| **Informe final de la Parte 2 (entregable central)** | `docs/Informe_Final_Parte2_Equipo07.pdf` |
| Fuente editable del informe final | `docs/Informe_Final_Parte2_Equipo07.docx` |
| Generador reproducible del informe final | `scripts/generar_informe.py` |
| Pauta oficial de la Parte 1 | `docs/Solemne01PracticaParte01FormaB.pdf` |
| Pauta oficial de la Parte 2 | `docs/Solemne01PracticaParte2FormaB.pdf` |
| Hito presencial ya entregado | `docs/hito_parte2hp_equipo07.pdf` |

El informe final no se redacta a mano: `scripts/generar_informe.py` lo arma desde una sola fuente de
contenido y **lee cada cifra de los archivos de `evidencias/debian/`** en el momento de generarlo,
de modo que el informe, este README y la evidencia no puedan divergir.

---

## 10. Reproducibilidad, resumen

Para llegar exactamente a los mismos números desde un Debian limpio:

```bash
git clone <url-del-repositorio>
cd Solemne1

python3 --version                              # verificado con 3.13.5

python3 scripts/limpiar.py                     # OBLIGATORIO: ver el paso 0

python3 scripts/generar_archivos_entrada.py    # semilla fija 20240908
python3 src/secuencial.py                      # 20 archivos, 300 válidas, 79 alertas
python3 src/concurrente.py                     # mismas métricas, 3 trabajadores

mkdir -p evidencias/salida_parte1              # ANTES del gestor
cp -a salida/informe_*.txt salida/resumen_ambiental.txt evidencias/salida_parte1/

python3 src/gestor_incidencias.py              # 0 / 8 / 12 informes, 79 alertas, 0 anomalías
```

**El paso de limpieza no es opcional.** El árbol se entrega con la salida de la corrida ya
documentada, porque la pauta exige los informes clasificados como evidencia. Sin limpiar,
el gestor reclasifica esos 20 informes sobre los 20 que ya venían y su protección
anti-sobrescritura crea 20 duplicados `_v1`: se obtienen 40 informes en vez de 20. El paso 0
explica el detalle y ofrece los comandos manuales para Debian y para Windows.

Este ciclo se verificó ejecutándolo **dos veces seguidas** sobre un árbol exportado limpio: la
segunda corrida entrega exactamente los mismos 0 / 8 / 12 informes, las mismas 79 alertas y las
mismas 0 anomalías que la primera.

No hace falta instalar nada. La reproducibilidad se apoya en cinco decisiones de diseño:

1. **Semilla fija** en el generador, con recorrido en orden explícito.
2. **Contrato de validación idéntico** en ambas versiones del procesador.
3. **Volcado ordenado del log de alertas** en la versión concurrente, para que la salida no dependa
   del entrelazado de los hilos.
4. **Bitácora por corrida**: `logs/gestion_ambiental.log` se reescribe completo en cada ejecución
   del gestor en vez de crecer, de modo que su tamaño es siempre el mismo y coincide con el que
   declara la sección 7.4. La decisión y su alternativa descartada (rotar) están documentadas en
   `src/gestor_incidencias.py`.
5. **Punto de partida explícito**: `scripts/limpiar.py` deja el árbol en el mismo estado desde el
   que partió la corrida oficial.
