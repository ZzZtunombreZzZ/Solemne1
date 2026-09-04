# Solemne 1 - Procesador de Mediciones meteorologicas

Los programas realizan un procesamiento de archivos de mediciones meteorológicas, generando un archivo de salida con los resultados del procesamiento.
donde existen 2 tipos de procesamientos

- Secuencial: Procesa los archivos de mediciones meteorológicas de manera secuencial, es decir, uno después del otro.
- Concurrente: Procesa los archivos de mediciones meteorológicas de manera concurrente, es decir, varios archivos al mismo tiempo.

# Requisitos
Python: 3.14

# Ejecución

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

# Estructura de archivos
```
├───alertas
├───entrada
├───salida
├───scripts
└───src
```
# Comparacion 
 
| Caracteristica | Secuencial | Concurrente |
|----------------|------------|-------------|
| Archivos procesados | 16 | 16 |
| Mediciones validas | 240 | 240 |
| Mediciones invalidas | 0 | 0 |
| Alertas detectadas | 458 | 236 |
| PM2.5 maximo global | 497.26 ug/m3 | 497.26 ug/m3 |
| Ruido maximo global | 138.23 dB | 138.23 dB |
| Tiempo de ejecucion | 0.018 s | 0.013 s |
| Cantidad de trabajadores | 1 | 3 |

# formato JSONL 

# reglas de alerta/validación.
