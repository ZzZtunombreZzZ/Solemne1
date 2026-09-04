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
├───00Info
├───alertas
├───entrada
├───salida
├───scripts
└───src
```
# Comparacion 

| Caracteristica | Secuencial | Concurrente |
|----------------|------------|-------------|
| Archivos procesados | Na | Na |
| MedicioNaes válidas | Na | Na |
| Mediciones inválidas | Na | Na |
| Alertas detectadas | Na | Na |
| PM2.5 máximo global | Na | Na |
| Ruido máximo global | Na | Na |
| Tiempo de ejecución | Na | Na |
| Cantidad de trabajadores | 1 | 4 |

# formato JSONL 

# reglas de alerta/validación.