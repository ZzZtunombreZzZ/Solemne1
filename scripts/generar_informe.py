# -*- coding: utf-8 -*-
"""
Generador reproducible del INFORME FINAL de la Parte 2 (Forma B), Equipo 07.

Produce dos archivos a partir de la MISMA fuente de contenido:

    docs/Informe_Final_Parte2_Equipo07.docx   (editable en Word / Google Docs)
    docs/Informe_Final_Parte2_Equipo07.pdf    (entregable de BlackBoard)

Principio de diseno: ningun dato del informe se escribe a mano en este script
si puede leerse de los archivos de evidencia del repositorio. Todas las cifras
(metricas, tiempos, PID, RSS, conteos de informes, alertas por indicador,
anomalias, tamanos, etc.) se extraen en tiempo de generacion de:

    evidencias/debian/01-preparar-ambiente.txt
    evidencias/debian/02-entorno.txt
    evidencias/debian/03-ejecucion-parte1.txt
    evidencias/debian/04-observacion-procesos.txt
    evidencias/debian/05-gestor-incidencias.txt
    evidencias/debian/06-control-de-errores.txt
    evidencias/debian/07-anti-sobrescritura.txt
    gestion_ambiental/  alertas/  logs/  salida/   (estado real del proyecto)

Si un dato no aparece en la evidencia, el script se detiene con un error en vez
de inventarlo.

Dependencias: python-docx, reportlab, Pillow (y pymupdf solo la primera vez,
para extraer de un PDF las capturas de instalacion; despues quedan en cache
dentro de evidencias/fotos/instalacion/).

Uso:
    python scripts/generar_informe.py
"""

import json
import re
import sys
import unicodedata
from pathlib import Path

# ---------------------------------------------------------------------------
# Rutas del proyecto
# ---------------------------------------------------------------------------

RAIZ = Path(__file__).resolve().parent.parent
EV = RAIZ / "evidencias" / "debian"
EV_HYPERV = RAIZ / "evidencias" / "hyperv"
FOTOS = RAIZ / "evidencias" / "fotos"
CACHE_INSTALACION = FOTOS / "instalacion"
FOTOS_HYPERV = FOTOS / "instalacion-hyperv"
CACHE_RECORTES = FOTOS / "recortes"
DOCS = RAIZ / "docs"
PDF_CAPTURAS = RAIZ / "evidencias" / "Parte 2 solemne Sistemas Operativos.pdf"

SALIDA_DOCX = DOCS / "Informe_Final_Parte2_Equipo07.docx"
SALIDA_PDF = DOCS / "Informe_Final_Parte2_Equipo07.pdf"

# ---------------------------------------------------------------------------
# Datos administrativos y de la maquina virtual que no viven en un archivo de
# evidencia de texto. Se declaran aqui de forma explicita para que queden a la
# vista y puedan corregirse en un solo lugar.
# ---------------------------------------------------------------------------

# ATENCION - VERIFICAR ANTES DE ENTREGAR: el nombre de la universidad NO
# aparece ni en la pauta (docs/Solemne01PracticaParte2FormaB.pdf) ni en el
# documento del hito (docs/hito_parte2hp_equipo07.pdf). Se dedujo del nombre de
# la carpeta del repositorio ("...uss..."). Si el ramo no es de la Universidad
# San Sebastian, corregir aqui: es el UNICO lugar donde aparece.
UNIVERSIDAD = "Universidad San Sebastian"
ASIGNATURA = "Sistemas Operativos"
EVALUACION = "Solemne 01 Practico - Parte 2 - Forma B"
SUBTITULO = "Informe final asincronico"
EQUIPO = "Equipo 07"
SECCION = "Seccion NRC 18897"
LENGUAJE = "Lenguaje utilizado en la Parte 1: Python"
INTEGRANTES = [
    "Benjamin Zamora",
    "Jose Palma",
    "Franco Maripil",
    "Nicolas Portilla",
    "Thomas Marquez",
]
FECHA = "09 de septiembre de 2026"

# El nombre del ISO, su suma SHA256 y la configuracion real de la maquina
# virtual YA NO se escriben aqui: se leen de
# evidencias/hyperv/configuracion-vm-hyperv.txt (salida de los cmdlets de
# Hyper-V y de Get-FileHash en el equipo anfitrion). Ver leer_hyperv().

# Unico dato de esta seccion que no vive en una salida de comando: lo que el
# equipo se comprometio a usar en el hito presencial.
HIPERVISOR_PLANIFICADO = "Oracle VirtualBox"

USUARIO_VM = "equipo07"
HOSTNAME_VM = "debian-so-equipo07"
RUTA_PROYECTO_VM = "/home/equipo07/solemne_so_equipo07"

# Umbrales de alerta, tomados del contrato comun de src/secuencial.py y
# src/concurrente.py (constante UMBRALES).
UMBRALES = [
    ("Temperatura", "mayor o igual a 35.0 C"),
    ("Humedad", "menor o igual a 20.0 %"),
    ("PM2.5", "mayor o igual a 35.0 ug/m3"),
    ("Ruido", "mayor o igual a 75.0 dB"),
]

RANGOS_VALIDOS = [
    ("Temperatura", "-20.0 C a 60.0 C"),
    ("Humedad", "0.0 % a 100.0 %"),
    ("PM2.5", "mayor o igual a 0.0 ug/m3"),
    ("Ruido", "0.0 dB a 140.0 dB"),
    ("Timestamp", "formato AAAA-MM-DDTHH:MM:SS"),
]

# Capturas de la instalacion REAL, hecha en Hyper-V sobre la maquina definitiva.
# Son la evidencia principal de la seccion de instalacion.
FOTOS_INSTALACION_HYPERV = [
    "01-menu-instalador-debian13.png",
    "02-linea-de-arranque-preseed.png",
    "03-instalacion-sistema-base.png",
    "04-instalacion-final.png",
    "05-primer-inicio.png",
]

# ---------------------------------------------------------------------------
# Numeracion de secciones.
#
# Los numeros NO se escriben a mano en ninguna parte del texto: h1()/h2() los
# asignan por orden de aparicion y los registran; el texto referencia una
# seccion por CLAVE mediante ref(). Al final, resolver_referencias() sustituye
# cada marca por su numero y aborta si la clave no corresponde a ninguna
# seccion existente. Asi una referencia cruzada no puede volver a apuntar a una
# seccion inexistente ni desincronizarse al reordenar el documento.
# ---------------------------------------------------------------------------

SECCIONES = {}
_CONTADOR = {"h1": 0, "h2": 0}
MARCA_REF = "sec:%s"


def _registrar(clave, numero):
    if clave in SECCIONES:
        morir("la clave de seccion %r esta duplicada" % clave)
    SECCIONES[clave] = numero
    return numero


def ref(clave):
    """Referencia a una seccion por clave; se resuelve al final."""
    return MARCA_REF % clave


def resolver_referencias(bloques):
    patron = re.compile("sec:([a-z0-9_]+)")

    def sub(texto):
        def uno(m):
            clave = m.group(1)
            if clave not in SECCIONES:
                morir("referencia cruzada a la seccion %r, que no existe en el "
                      "documento. Secciones validas: %s"
                      % (clave, ", ".join(sorted(SECCIONES))))
            return SECCIONES[clave]
        return patron.sub(uno, texto)

    def recorrer(v):
        if isinstance(v, str):
            return sub(v)
        if isinstance(v, list):
            return [recorrer(x) for x in v]
        if isinstance(v, dict):
            return {k: recorrer(x) for k, x in v.items()}
        return v

    return recorrer(bloques)


# ---------------------------------------------------------------------------
# Utilidades de lectura y extraccion. Fallan de forma ruidosa.
# ---------------------------------------------------------------------------

def morir(mensaje):
    print("[generar_informe] ERROR: " + mensaje, file=sys.stderr)
    raise SystemExit(1)


def leer(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        morir("no existe el archivo de evidencia %s" % ruta)
    return ruta.read_text(encoding="utf-8", errors="replace")


def buscar(texto, patron, etiqueta, flags=0, grupo=1):
    m = re.search(patron, texto, flags)
    if not m:
        morir("no se pudo extraer %s (patron %r)" % (etiqueta, patron))
    return m.group(grupo).strip()


def buscar_todos(texto, patron, etiqueta, flags=0, minimo=1):
    res = re.findall(patron, texto, flags)
    if len(res) < minimo:
        morir("se esperaban al menos %d coincidencias de %s" % (minimo, etiqueta))
    return res


def bloque_entre(texto, inicio, fin, etiqueta, incluir_inicio=False):
    """Devuelve las lineas entre dos marcas, sin incluirlas (salvo el inicio)."""
    lineas = texto.splitlines()
    i = None
    for n, ln in enumerate(lineas):
        if inicio in ln:
            i = n
            break
    if i is None:
        morir("no se encontro la marca de inicio %r para %s" % (inicio, etiqueta))
    j = None
    for n in range(i + 1, len(lineas)):
        if fin in lineas[n]:
            j = n
            break
    if j is None:
        morir("no se encontro la marca de fin %r para %s" % (fin, etiqueta))
    inicio_real = i if incluir_inicio else i + 1
    return "\n".join(lineas[inicio_real:j]).strip("\n")


def bloque_tras(texto, marca, n_lineas, etiqueta, saltar_vacias=True):
    """Devuelve n_lineas utiles a continuacion de una marca."""
    lineas = texto.splitlines()
    idx = None
    for n, ln in enumerate(lineas):
        if marca in ln:
            idx = n
            break
    if idx is None:
        morir("no se encontro la marca %r para %s" % (marca, etiqueta))
    salida = []
    n = idx + 1
    while n < len(lineas) and len(salida) < n_lineas:
        ln = lineas[n]
        if saltar_vacias and not ln.strip() and not salida:
            n += 1
            continue
        if not ln.strip():
            break
        salida.append(ln.rstrip())
        n += 1
    if not salida:
        morir("no habia contenido tras la marca %r (%s)" % (marca, etiqueta))
    return "\n".join(salida)


def contar_lineas(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        morir("no existe %s" % ruta)
    texto = ruta.read_text(encoding="utf-8", errors="replace")
    return len([ln for ln in texto.splitlines() if ln.strip()])


def contar_archivos(directorio, patron="*.txt"):
    d = Path(directorio)
    if not d.exists():
        morir("no existe el directorio %s" % d)
    return len(sorted(d.glob(patron)))


def tam_bytes(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        morir("no existe %s" % ruta)
    return ruta.stat().st_size


REEMPLAZOS = {
    "≥": ">=", "≤": "<=", "→": "->", "←": "<-",
    "–": "-", "—": "-", "‘": "'", "’": "'",
    "“": '"', "”": '"', "…": "...", "µ": "u",
    "μ": "u", "·": "-", "•": "-", "●": "-",
    "─": "-", "│": "|", "├": "|", "└": "|",
    "┴": "|", "┬": "|", "┼": "|", "�": "",
    " ": " ", "\t": "    ",
}


def sin_acentos(texto):
    """Quita tildes y dieresis, conservando el resto del texto."""
    normal = unicodedata.normalize("NFD", texto)
    return "".join(ch for ch in normal if unicodedata.category(ch) != "Mn")


def limpiar(texto):
    """Normaliza el texto a caracteres representables por las fuentes base."""
    if texto is None:
        return ""
    for k, v in REEMPLAZOS.items():
        texto = texto.replace(k, v)
    salida = []
    for ch in texto:
        if ch in ("\n", "\r"):
            salida.append(ch)
        elif ord(ch) < 256:
            salida.append(ch)
        else:
            salida.append("?")
    return "".join(salida)


# ---------------------------------------------------------------------------
# Preparacion de imagenes
# ---------------------------------------------------------------------------

def extraer_capturas_instalacion():
    """Extrae del PDF de capturas del equipo las imagenes de instalacion.

    Se cachean en evidencias/fotos/instalacion/ para que el informe pueda
    regenerarse aunque pymupdf no este disponible.
    """
    deseadas = {
        "01-descarga-iso.png": (1, 17),
        "02-recursos-vm.png": (2, 21),
        "03-instalacion-en-curso.png": (2, 22),
        "04-primer-inicio.png": (2, 23),
        "05-preparacion-entorno.png": (4, 28),
    }
    CACHE_INSTALACION.mkdir(parents=True, exist_ok=True)
    faltantes = {k: v for k, v in deseadas.items()
                 if not (CACHE_INSTALACION / k).exists()}
    if faltantes:
        try:
            import pymupdf
        except ImportError:
            try:
                import fitz as pymupdf  # nombre antiguo
            except ImportError:
                morir("faltan capturas en %s y no hay pymupdf para extraerlas "
                      "de %s" % (CACHE_INSTALACION, PDF_CAPTURAS))
        doc = pymupdf.open(str(PDF_CAPTURAS))
        for nombre, (pagina, xref) in faltantes.items():
            px = pymupdf.Pixmap(doc, xref)
            if px.n > 4:
                px = pymupdf.Pixmap(pymupdf.csRGB, px)
            px.save(str(CACHE_INSTALACION / nombre))
        doc.close()
    return {k: CACHE_INSTALACION / k for k in deseadas}


def capturas_instalacion_real():
    """Capturas de la instalacion definitiva, hecha sobre Hyper-V.

    No se extraen de ningun PDF: son archivos del repositorio. Si falta alguna,
    el generador se detiene, porque son la evidencia principal de la seccion de
    instalacion.
    """
    salida = {}
    for nombre in FOTOS_INSTALACION_HYPERV:
        ruta = FOTOS_HYPERV / nombre
        if not ruta.exists():
            morir("falta la captura de la instalacion real %s" % ruta)
        # Se recortan al area util igual que las de consola: sin recortar, el
        # framebuffer completo encoge tanto en el informe que el texto no se lee.
        salida[nombre] = recortar_consola(ruta)
    return salida


# Version del algoritmo de recorte. Si cambia, la cache de recortes se
# invalida completa: de otro modo un recorte hecho con la version anterior
# sobreviviria a la mejora y la figura seguiria saliendo con el fondo vacio.
VERSION_RECORTE = "2"


def _color_de_fondo(im):
    """Color de fondo de una captura, deducido del anillo de borde completo.

    No basta con mirar las cuatro esquinas: las capturas del framebuffer de
    Hyper-V traen un artefacto de unos pocos pixeles en la esquina superior
    izquierda (el cursor), y si esa esquina cae dentro del artefacto el color
    de fondo queda mal deducido y el recorte no recorta nada. Tomando el color
    mas frecuente de todo el borde, un artefacto de pocos pixeles no puede
    ganar la votacion.
    """
    from collections import Counter
    w, h = im.size
    muestras = []
    muestras += [im.getpixel((x, 0)) for x in range(0, w, 3)]
    muestras += [im.getpixel((x, h - 1)) for x in range(0, w, 3)]
    muestras += [im.getpixel((0, y)) for y in range(0, h, 3)]
    muestras += [im.getpixel((w - 1, y)) for y in range(0, h, 3)]
    return Counter(muestras).most_common(1)[0][0]


def _mascara_de_contenido(im, fondo=None, umbral=40):
    """Mascara binaria: 255 donde la imagen se aparta del color de fondo."""
    from PIL import Image, ImageChops
    if fondo is None:
        fondo = _color_de_fondo(im)
    dif = ImageChops.difference(im, Image.new("RGB", im.size, fondo))
    return dif.convert("L").point(lambda v: 255 if v > umbral else 0)


def recortar_consola(origen):
    """Recorta una captura de pantalla de la VM a su area util.

    Las capturas son el framebuffer completo (1280x800 o 1600x1200) y suelen
    tener la mayor parte vacia. Si se encoge la imagen entera para que quepa en
    el informe, el texto queda ilegible: ese fue el defecto que hacia que las
    figuras del instalador aparecieran al 15 % de su tamano nativo con dos
    tercios de fondo azul.

    El fondo NO siempre es negro: la consola de texto es negra, pero las
    pantallas del instalador de Debian son azules con un cuadro de dialogo
    gris. Por eso el color de fondo se deduce del borde de la imagen.

    El recorte NO usa getbbox(): un solo pixel encendido en una esquina (el
    cursor del framebuffer de Hyper-V) basta para que getbbox() devuelva la
    imagen completa y el recorte no sirva de nada. En vez de eso se construye
    el perfil de contenido por filas y por columnas y se conservan solo las
    filas y columnas cuyo contenido supera un umbral proporcional al tamano de
    la imagen, de modo que las motas aisladas no arrastran el recorte.
    """
    from PIL import Image

    origen = Path(origen)
    CACHE_RECORTES.mkdir(parents=True, exist_ok=True)
    marca = CACHE_RECORTES / ".version"
    vigente = (marca.read_text(encoding="utf-8").strip()
               if marca.exists() else None)
    if vigente != VERSION_RECORTE:
        # Sin marca, la cache viene de una version desconocida del algoritmo y
        # tambien hay que rehacerla: si no, un recorte antiguo sobrevive a la
        # mejora y la figura sigue saliendo con el fondo vacio.
        for viejo in CACHE_RECORTES.glob("*.png"):
            viejo.unlink()
        marca.write_text(VERSION_RECORTE, encoding="utf-8")

    destino = CACHE_RECORTES / origen.name
    if destino.exists() and destino.stat().st_mtime >= origen.stat().st_mtime:
        return destino

    im = Image.open(origen).convert("RGB")
    masc = _mascara_de_contenido(im)
    w, h = masc.size
    px = masc.load()

    filas = [sum(1 for x in range(w) if px[x, y]) for y in range(h)]
    cols = [sum(1 for y in range(h) if px[x, y]) for x in range(w)]
    # Umbral proporcional: una linea de texto enciende decenas de pixeles en su
    # fila; una mota de cursor enciende dos o tres.
    umbral_fila = max(3, int(0.004 * w))
    umbral_col = max(3, int(0.004 * h))
    ys = [y for y, c in enumerate(filas) if c >= umbral_fila]
    xs = [x for x, c in enumerate(cols) if c >= umbral_col]

    if ys and xs:
        m = 10
        caja = (max(0, xs[0] - m), max(0, ys[0] - m),
                min(w, xs[-1] + 1 + m), min(h, ys[-1] + 1 + m))
        if (caja[2] - caja[0]) * (caja[3] - caja[1]) < w * h * 0.98:
            im = im.crop(caja)
    im.save(destino)
    return destino


def bandas_de_contenido(ruta, hueco=3, umbral=0.01):
    """Bandas horizontales con contenido de una captura ya recortada.

    Devuelve una lista de tuplas (y_inicial, y_final, llenado_maximo), donde
    llenado_maximo es la fraccion de la anchura encendida en la fila mas llena
    de la banda. Sirve para razonar sobre lo que la captura MUESTRA, en vez de
    suponerlo desde el texto del informe.
    """
    from PIL import Image

    im = Image.open(ruta).convert("RGB")
    masc = _mascara_de_contenido(im, umbral=30)
    w, h = masc.size
    px = masc.load()
    perfil = [sum(1 for x in range(w) if px[x, y]) / float(w) for y in range(h)]

    bandas = []
    ini = None
    blanco = 0
    for y, llenado in enumerate(perfil):
        if llenado > umbral:
            if ini is None:
                ini = y
            blanco = 0
        elif ini is not None:
            blanco += 1
            if blanco > hueco:
                fin = y - blanco
                bandas.append((ini, fin, max(perfil[ini:fin + 1])))
                ini = None
    if ini is not None:
        bandas.append((ini, h - 1, max(perfil[ini:])))
    return bandas


def altura_de_linea_px(ruta):
    """Altura, en pixeles, de una linea de texto dentro de una captura.

    Se toma la mediana de las bandas de contenido con tamano de linea de texto
    (hasta 40 px). Si la captura no tiene al menos tres bandas de ese tipo no
    es una transcripcion de terminal -por ejemplo el cuadro de progreso del
    instalador, que es un unico bloque grafico- y se devuelve None: ahi la
    altura de linea no es la medida que decide si se lee.
    """
    import statistics
    alturas = sorted(b[1] - b[0] + 1 for b in bandas_de_contenido(ruta)
                     if (b[1] - b[0] + 1) <= 40)
    if len(alturas) < 3:
        return None
    return statistics.median(alturas)


def filas_visibles_de_la_tabla(ruta):
    """Cuantas filas de datos MUESTRA la captura de comparacion de la Parte 1.

    La tabla impresa en la consola esta delimitada por dos lineas de guiones
    que ocupan casi todo el ancho. Si entre esas dos lineas no hay ninguna
    banda de contenido, la tabla salio vacia y el pie de la figura no puede
    afirmar que "las seis metricas coinciden": la imagen no lo muestra.

    Devuelve el numero de bandas de texto entre los dos separadores, o None si
    no se reconocen los dos separadores (en cuyo caso el pie tampoco debe
    prometer nada).
    """
    bandas = bandas_de_contenido(ruta)
    separadores = [b for b in bandas if b[2] >= 0.45 and (b[1] - b[0]) <= 6]
    if len(separadores) < 2:
        return None
    arriba, abajo = separadores[0], separadores[1]
    return sum(1 for b in bandas if b[0] > arriba[1] and b[1] < abajo[0])


# ---------------------------------------------------------------------------
# Lectura de TODA la evidencia
# ---------------------------------------------------------------------------

def leer_hyperv():
    """Configuracion real de la VM y verificacion del ISO.

    Sustituye a las constantes que antes estaban escritas a mano en este
    script: todo sale de evidencias/hyperv/configuracion-vm-hyperv.txt, que es
    la salida literal de los cmdlets de Hyper-V y de Get-FileHash.
    """
    t = leer(EV_HYPERV / "configuracion-vm-hyperv.txt")
    h = {}
    h["vm_nombre"] = buscar(t, r"^Name\s+: (\S+)$", "nombre de la VM", re.M)
    h["vm_estado"] = buscar(t, r"^State\s+: (\S+)$", "estado de la VM", re.M)
    h["generacion"] = buscar(t, r"^Generation\s+: (\d+)$", "generacion de la VM", re.M)
    h["vcpu"] = buscar(t, r"^ProcessorCount\s+: (\d+)$", "vCPU de la VM", re.M)
    h["ram_gb"] = buscar(t, r"^MemoryStartupGB\s+: (\d+)$", "RAM de la VM", re.M)
    h["ram_dinamica"] = buscar(t, r"^DynamicMemoryEnabled\s+: (\S+)$",
                               "memoria dinamica", re.M)
    h["vhd_tipo"] = buscar(t, r"^VhdType\s+: (\S+)$", "tipo de disco virtual", re.M)
    h["vhd_gb"] = buscar(t, r"^SizeGB\s+: (\d+)$", "tamano del disco virtual", re.M)
    h["vhd_real_gb"] = buscar(t, r"^FileSizeGB\s+: (\d+)$", "ocupacion del VHDX", re.M)
    h["switch"] = buscar(t, r"^SwitchName\s+: (.+)$", "conmutador virtual", re.M)
    h["switch_tipo"] = buscar(t, r"^SwitchType\s+: (\S+)$", "tipo de conmutador", re.M)
    h["mac"] = buscar(t, r"^MacAddress\s+: (\S+)$", "MAC de la VM", re.M)
    h["ip_hyperv"] = buscar(t, r"^IPAddresses\s+: ([0-9.]+)", "IP vista por Hyper-V", re.M)
    h["nat"] = buscar(t, r"^(El Default Switch de Hyper-V provee .+)$",
                      "descripcion del NAT", re.M)
    h["orden_arranque"] = buscar(t, r"^StartupOrder\s+: (.+)$", "orden de arranque", re.M)
    h["iso_nombre"] = buscar(t, r"Get-FileHash (\S+\.iso)", "nombre del ISO")
    h["iso_sha_calculado"] = buscar(t, r"Hash calculado\s+: ([0-9a-f]{64})",
                                    "SHA256 calculado del ISO")
    h["iso_sha_publicado"] = buscar(t, r"Hash publicado\s+: ([0-9a-f]{64})",
                                    "SHA256 publicado del ISO")
    h["iso_fuente"] = buscar(t, r"Fuente: (\S+)", "fuente del SHA256SUMS")
    h["iso_coincide"] = buscar(t, r"^Coincide\s+: (\S+)$", "veredicto del SHA256", re.M)
    h["iso_tamano"] = buscar(t, r"^Tamano\s+: (.+)$", "tamano del ISO", re.M)

    if h["iso_sha_calculado"] != h["iso_sha_publicado"] or h["iso_coincide"] != "SI":
        morir("la evidencia de Hyper-V declara que el SHA256 del ISO no coincide")

    h["hipervisor_implementado"] = (
        "Microsoft Hyper-V (maquina de Generacion " + h["generacion"]
        + ", arranque BIOS)")
    return h


def recolectar_datos():
    d = {}

    t01 = leer(EV / "01-preparar-ambiente.txt")
    t02 = leer(EV / "02-entorno.txt")
    t03 = leer(EV / "03-ejecucion-parte1.txt")
    t04 = leer(EV / "04-observacion-procesos.txt")
    t05 = leer(EV / "05-gestor-incidencias.txt")
    t06 = leer(EV / "06-control-de-errores.txt")
    t07 = leer(EV / "07-anti-sobrescritura.txt")

    d["hv"] = leer_hyperv()

    # --- 01 preparacion del ambiente ---------------------------------------
    d["apt_update"] = bloque_entre(t01, "$ sudo apt update", "$ sudo apt upgrade -y",
                                   "salida de apt update").strip()
    d["apt_upgrade"] = bloque_entre(t01, "$ sudo apt upgrade -y",
                                    "--- Entorno del lenguaje",
                                    "salida de apt upgrade").strip()
    d["python_version"] = buscar(t01, r"^(Python 3\.[0-9.]+)$", "version de Python",
                                 re.M)
    d["apt_list"] = bloque_tras(t01, "$ apt list --installed", 4,
                                "paquetes instalados")
    d["sin_dependencias"] = buscar(
        t01, r"^(El proyecto usa exclusivamente la biblioteca estandar de Python)",
        "declaracion de dependencias", re.M)
    d["compilacion"] = buscar(t01, r"^(Los \w+ modulos compilan sin errores\.)$",
                              "comprobacion de compilacion", re.M)

    # --- 02 entorno ---------------------------------------------------------
    d["pretty_name"] = buscar(t02, r'PRETTY_NAME="([^"]+)"', "distribucion")
    d["version_debian"] = bloque_tras(t02, "$ cat /etc/debian_version", 1,
                                      "version de Debian")
    d["kernel"] = bloque_tras(t02, "$ uname -a", 1, "kernel")
    d["hipervisor_detectado"] = bloque_tras(t02, "$ systemd-detect-virt", 1,
                                            "hipervisor detectado")
    d["cpus"] = buscar(t02, r"^CPU\(s\):\s+(\d+)", "numero de vCPU", re.M)
    d["modelo_cpu"] = buscar(t02, r"^Model name:\s+(.+)$", "modelo de CPU", re.M)
    d["hipervisor_vendor"] = buscar(t02, r"^Hypervisor vendor:\s+(.+)$",
                                    "fabricante del hipervisor", re.M)
    d["memoria"] = bloque_tras(t02, "$ free -h", 3, "salida de free")
    d["lsblk"] = bloque_tras(t02, "$ lsblk", 8, "salida de lsblk")
    d["df_raiz"] = bloque_tras(t02, "Filesystem      Size  Used Avail Use% Mounted on",
                               1, "df de la raiz")
    d["red"] = bloque_tras(t02, "$ ip -4 -br addr show", 4, "configuracion de red")
    d["ip_eth0"] = buscar(t02, r"eth0\s+UP\s+([0-9./]+)", "direccion IP")
    d["gateway"] = buscar(t02, r"default via ([0-9.]+) dev eth0", "puerta de enlace")
    d["hostname_vm"] = buscar(t02, r"^(debian-so-\S+)$", "nombre de host", re.M)
    d["usuario_vm"] = buscar(t02, r"^uid=\d+\((\w+)\)", "usuario de la VM", re.M)
    d["dns"] = bloque_tras(t02, "$ getent hosts deb.debian.org", 1,
                           "resolucion DNS")

    # --- 03 ejecucion de la Parte 1 ----------------------------------------
    d["dir_trabajo"] = buscar(t03, r"Directorio: (\S+)", "directorio de trabajo")
    d["archivos_generados"] = buscar(t03, r"Archivos \.jsonl generados: (\d+)",
                                     "archivos de entrada generados")
    d["semilla"] = buscar(t03, r"Semilla fija: (\d+)", "semilla del generador")
    d["checksum_entrada"] = buscar(t03, r"^([0-9a-f]{64})\s+-$",
                                   "suma de verificacion de la entrada", re.M)

    seg_sec = bloque_entre(t03, "=== 2. VERSION SECUENCIAL ===",
                           "=== 3. VERSION CONCURRENTE", "bloque secuencial")
    seg_con = bloque_entre(t03, "=== 3. VERSION CONCURRENTE",
                           "# 4. CONSISTENCIA EXIGIDA", "bloque concurrente")

    def tiempos(seg, etiqueta):
        return {
            "real": buscar(seg, r"real\s+(\S+)", "tiempo real " + etiqueta),
            "user": buscar(seg, r"user\s+(\S+)", "tiempo user " + etiqueta),
            "sys": buscar(seg, r"sys\s+(\S+)", "tiempo sys " + etiqueta),
            "interno": buscar(seg, r"Tiempo total de ejecuci[^:]*: ([0-9.]+) segundos",
                              "tiempo interno " + etiqueta),
            "trabajadores": buscar(seg, r"Cantidad de trabajadores: (\d+)",
                                   "trabajadores " + etiqueta),
        }

    d["t_sec"] = tiempos(seg_sec, "secuencial")
    d["t_con"] = tiempos(seg_con, "concurrente")
    d["resumen_secuencial"] = bloque_entre(
        seg_sec, "RESUMEN CONSOLIDADO DE MONITOREO AMBIENTAL",
        "Informes individuales generados", "resumen secuencial",
        incluir_inicio=True)
    d["informes_generados"] = buscar(seg_sec, r"Informes individuales generados: (\d+)",
                                     "informes generados")

    filas = []
    en_tabla = False
    for ln in t03.splitlines():
        if ln.startswith("Metrica") and "Secuencial" in ln:
            en_tabla = True
            continue
        if en_tabla:
            if set(ln.strip()) == {"-"}:
                if filas:
                    break
                continue
            if not ln.strip():
                continue
            partes = re.split(r"\s{2,}", ln.strip())
            if len(partes) == 4:
                filas.append(partes)
    if len(filas) != 6:
        morir("se esperaban 6 metricas comparadas y se leyeron %d" % len(filas))
    d["tabla_metricas"] = filas
    # No se escribe "6 de 6" a mano: se cuenta sobre la tabla recien leida.
    coinciden = [f for f in filas if f[3].upper().startswith("SI")]
    if len(coinciden) != len(filas):
        morir("la evidencia declara %d de %d metricas coincidentes"
              % (len(coinciden), len(filas)))
    d["metricas_comparadas"] = "%d de %d" % (len(coinciden), len(filas))
    d["n_metricas"] = str(len(filas))
    d["resultado_consistencia"] = buscar(t03, r"^RESULTADO: (.+)$",
                                         "resultado de consistencia", re.M)
    d["informes_comparados"] = buscar(t03, r"Informes comparados: (\d+)",
                                      "informes comparados")
    d["informes_con_diferencias"] = buscar(t03, r"con diferencias: (\d+)",
                                           "informes con diferencias")
    seg_log = bloque_entre(t03, "=== 6. Bitacora de alertas ===",
                           "IDENTICO BYTE A BYTE", "bloque de la bitacora de alertas")
    d["alertas_sec"] = buscar(seg_log, r"Secuencial\s*: (\d+) alertas",
                              "alertas del secuencial")
    d["alertas_con"] = buscar(seg_log, r"Concurrente\s*: (\d+) alertas",
                              "alertas del concurrente")
    d["veredicto_log"] = buscar(t03, r"^(IDENTICO BYTE A BYTE .+)$",
                                "veredicto del log de alertas", re.M)

    resumen = {}
    for ln in d["resumen_secuencial"].splitlines():
        if ":" in ln:
            k, v = ln.split(":", 1)
            resumen[sin_acentos(k).strip()] = limpiar(v).strip()
    for clave in ("Archivos procesados", "Lineas leidas", "Mediciones validas",
                  "Mediciones invalidas", "Alertas totales"):
        if clave not in resumen:
            morir("el resumen consolidado no trae la clave %r" % clave)
    d["resumen"] = resumen

    # --- 04 observacion de procesos ----------------------------------------
    d["time_v"] = "\n".join(
        ln.rstrip() for ln in bloque_entre(t04, "$ /usr/bin/time -v",
                                           "(B) Observacion en vivo",
                                           "salida de /usr/bin/time -v").splitlines()
        if ln.strip() and set(ln.strip()) != {"="}).strip()
    d["carga_archivos"] = buscar(t04, r"Archivos: (\d+) \| Lineas: \d+",
                                 "archivos de la carga amplificada")
    d["carga_lineas"] = buscar(t04, r"Archivos: \d+ \| Lineas: (\d+)",
                               "lineas de la carga amplificada")
    d["muestras_ps"] = buscar(t04, r"Muestras de ps tomadas[^:]*: (\d+)",
                              "muestras de ps")
    d["ps_cmd"] = buscar(t04, r"^(\$ ps -o [^\n]+)$", "comando ps", re.M)
    d["ps_cabecera"] = buscar(t04, r"^(\s*PID\s+PPID STAT.*)$", "cabecera de ps", re.M)
    linea_ps = buscar(
        t04,
        r"^(\s*\d+\s+\d+ Sl\s+[\d.]+\s+[\d.]+\s+\d+\s+\d+\s+\d+\s+\S+\s+python3.*)$",
        "muestra de ps", re.M)
    d["ps_linea"] = linea_ps
    campos = linea_ps.split()
    d["pid"] = campos[0]
    d["ppid"] = campos[1]
    d["stat"] = campos[2]
    d["pcpu"] = campos[3]
    d["pmem"] = campos[4]
    d["rss"] = campos[5]
    d["vsz"] = campos[6]
    d["nlwp"] = campos[7]
    d["ps_evolucion"] = bloque_entre(
        t04, "--- Evolucion durante la ejecucion",
        "--- Hilos del proceso", "evolucion de CPU").strip("\n")
    bloque_hilos = bloque_entre(t04, "--- Hilos del proceso",
                                "--- Proceso padre", "hilos del proceso").strip("\n")
    d["ps_hilos"] = "\n".join(ln for ln in bloque_hilos.splitlines()
                              if not re.match(r"^\w+:\s", ln)).strip("\n")
    d["proc_status"] = "\n".join(ln for ln in bloque_hilos.splitlines()
                                 if re.match(r"^\w+:\s", ln)).strip("\n")
    if not d["proc_status"]:
        morir("no se encontro la vista de /proc en la evidencia 04")
    d["proc_threads"] = buscar(t04, r"^Threads:\s+(\d+)", "hilos segun /proc", re.M)
    d["proc_vmrss"] = buscar(t04, r"^VmRSS:\s+(\d+) kB", "VmRSS", re.M)
    d["proc_estado"] = buscar(t04, r"^State:\s+(.+)$", "estado en /proc", re.M)
    d["padre"] = bloque_entre(t04, "--- Proceso padre ---",
                              "Memoria del sistema, espacio disponible",
                              "proceso padre").strip().strip("=").strip()
    d["free_h"] = bloque_tras(t04, "$ free -h", 3, "free -h")
    d["df_h"] = bloque_tras(t04, "$ df -h", 12, "df -h")
    d["du_proyecto"] = buscar(t04, r"^(\S+)\s+/home/\w+/solemne_so_equipo07$",
                              "tamano del proyecto", re.M)
    d["du_detalle"] = bloque_tras(t04, "$ du -sh ~/solemne_so_equipo07/*", 8,
                                  "detalle de du")
    d["max_rss"] = buscar(t04, r"Maximum resident set size \(kbytes\): (\d+)",
                          "RSS maximo")
    d["pct_cpu_oficial"] = buscar(t04, r"Percent of CPU this job got: (\S+)",
                                  "porcentaje de CPU oficial")
    d["ctx_vol"] = buscar(t04, r"Voluntary context switches: (\d+)",
                          "cambios de contexto voluntarios")
    d["ctx_invol"] = buscar(t04, r"Involuntary context switches: (\d+)",
                            "cambios de contexto involuntarios")
    d["fs_outputs"] = buscar(t04, r"File system outputs: (\d+)",
                             "escrituras al sistema de archivos")
    d["pcpu_max"] = max(
        buscar_todos(t04, r"^\s*\d+\s+\d+ \S+\s+([\d.]+)\s", "%CPU", re.M),
        key=lambda x: float(x))
    d["copias_demo"] = buscar(
        t04, r"^(Ambas(?: partes)? se ejecutan sobre COPIAS del proyecto.+)$",
        "declaracion de uso de copias", re.M)

    # --- 05 gestor de incidencias ------------------------------------------
    d["find_dirs"] = bloque_entre(t05, "$ find gestion_ambiental logs -type d | sort",
                                  "$ find gestion_ambiental -type f",
                                  "directorios generados").strip()
    lista_archivos = bloque_entre(t05, "$ find gestion_ambiental -type f | sort",
                                  "--- Informes clasificados ---",
                                  "archivos bajo gestion_ambiental")
    d["find_files"] = [ln.strip() for ln in lista_archivos.splitlines() if ln.strip()]
    d["find_files_total"] = str(len(d["find_files"]))
    d["ls_inventario"] = bloque_tras(
        t05, "$ ls -lah gestion_ambiental/inventario_ambiental.json", 2,
        "listado del inventario")
    d["stat_bitacora"] = bloque_entre(t05, "$ stat logs/gestion_ambiental.log",
                                      "$ ls -lah logs/gestion_ambiental.log",
                                      "stat de la bitacora").strip("\n")
    d["ls_bitacora"] = bloque_tras(t05, "$ ls -lah logs/gestion_ambiental.log", 1,
                                   "ls de la bitacora")
    d["mover_copiar"] = bloque_entre(
        t05, "--- MOVER contra COPIAR ---", "# Metadatos de la bitacora",
        "comprobacion mover vs copiar").strip().strip("#").strip()
    d["bitacora_final"] = bloque_tras(t05, "--- Cierre de la bitacora ---", 20,
                                      "ultimas lineas de la bitacora")
    d["json_valido"] = buscar(t05, r"^(JSON valido)\.", "validacion del JSON", re.M)
    d["json_claves"] = buscar(t05, r"JSON valido\. Claves: \[(.+)\]",
                              "claves del inventario")
    d["n_claves_inventario"] = str(len(d["json_claves"].split(",")))
    d["gestor_exit"] = buscar(t05, r"Codigo de salida: (\d+)",
                              "codigo de salida del gestor")
    d["respaldo_parte1"] = buscar(t05, r"respaldo: (\d+) archivos",
                                  "respaldo de la Parte 1")

    # Conteos CONGELADOS de la corrida: se leen de la evidencia, no del disco.
    # El estado vivo se compara despues en verificar_coherencia().
    d["ev_n_bitacora"] = buscar(t05, r"^Lineas de la bitacora: (\d+)$",
                                "lineas de la bitacora segun la evidencia", re.M)
    for clave, carpeta in (("ev_n_sin", "sin_alertas"), ("ev_n_con", "con_alertas"),
                           ("ev_n_crit", "criticas")):
        d[clave] = buscar(
            t05,
            r"\$ ls -lah gestion_ambiental/" + carpeta
            + r"[^\n]*\n(?:.*\n)*?\s*-> (\d+) informes",
            "informes en " + carpeta + " segun la evidencia")
    d["ev_indicadores"] = {}
    for ind in ("temperatura", "humedad", "pm25", "ruido"):
        d["ev_indicadores"][ind] = buscar(
            t05, r"^\s+" + ind + r"\s+(\d+) alertas", "alertas de " + ind, re.M)
    d["ev_suma_logs"] = buscar(t05, r"SUMA de los 4 logs\s*: (\d+)",
                               "suma de los logs por indicador")
    d["ev_alertas_log"] = buscar(t05, r"alertas/alertas_detectadas\.log : (\d+)",
                                 "alertas en la bitacora original")
    d["ev_alertas_resumen"] = buscar(t05, r"resumen consolidado\s*: (\d+)",
                                     "alertas en el resumen consolidado")
    d["ev_alertas_inventario"] = buscar(t05, r"inventario_ambiental\.json\s*: (\d+)",
                                        "alertas en el inventario")
    d["ev_entradas"] = d["archivos_generados"]

    stat = d["stat_bitacora"]
    d["stat_size"] = buscar(stat, r"Size: (\d+)", "tamano de la bitacora")
    d["stat_inodo"] = buscar(stat, r"Inode: (\d+)", "inodo de la bitacora")
    d["stat_enlaces"] = buscar(stat, r"Links: (\d+)", "enlaces de la bitacora")
    d["stat_permisos"] = buscar(stat, r"Access: \((\S+)\)", "permisos de la bitacora")
    d["stat_uid"] = buscar(stat, r"Uid: \(\s*(\d+/\s*\S+)\)", "propietario")
    d["stat_gid"] = buscar(stat, r"Gid: \(\s*(\d+/\s*\S+)\)", "grupo")
    d["stat_bloques"] = buscar(stat, r"Blocks: (\d+)", "bloques")
    d["stat_ioblock"] = buscar(stat, r"IO Block: (\d+)", "tamano de bloque")
    d["stat_dispositivo"] = buscar(stat, r"Device: (\S+)", "dispositivo")

    # --- 06 control de errores ----------------------------------------------
    d["copia_control"] = buscar(
        t06, r"COPIA del proyecto\s*\n?\((\S+?)\)",
        "ruta de la copia usada en el control de errores")
    d["anomalias_lineas"] = [ln.strip() for ln in t06.splitlines()
                             if "ANOMALIA:" in ln]
    if len(d["anomalias_lineas"]) < 4:
        morir("se esperaban 4 anomalias registradas en la bitacora")
    d["control_errores"] = buscar(
        t06, r"(CONTROL DE ERRORES: se detectaron \d+ anomalias.+)$",
        "resumen de control de errores", re.M)
    d["anomalias_total"] = buscar(t06, r"se detectaron (\d+) anomalias",
                                  "total de anomalias")
    d["fallas_inyectadas"] = [
        ln.strip() for ln in bloque_entre(
            t06, "=== Se inyectan cuatro fallas distintas ===",
            "$ python3 src/gestor_incidencias.py",
            "fallas inyectadas").splitlines() if ln.strip()]
    d["gestor_exit_error"] = buscar(
        t06,
        r"\$ python3 src/gestor_incidencias\.py\s*\n"
        r"Codigo de salida: (\d+)\s+\(el gestor NO se detuvo\)",
        "codigo de salida con entrada danada")
    d["inventario_tras_error"] = buscar(t06, r"alertas conservadas: (\d+)",
                                        "alertas conservadas tras el error")
    d["ctl_informe"] = buscar(t06, r"(informe_CTL00_\S+\.txt)", "informe de control")
    clas = buscar_todos(t06, r"^\s*(sin_alertas|con_alertas|criticas)/\s+(\d+) informes$",
                        "clasificacion con el informe de control", re.M, minimo=3)
    d["clasificacion_ctl"] = ", ".join("%s/ %s" % (c, n) for c, n in clas)

    # --- 07 anti sobrescritura ----------------------------------------------
    d["copia_anti"] = buscar(t07, r"se ejecuta sobre la COPIA (\S+)",
                             "ruta de la copia usada en la anti-sobrescritura")
    d["ciclos"] = buscar_todos(
        t07, r"informes acumulados: (\d+) \| con sufijo de version: (\d+)",
        "ciclos de anti-sobrescritura", minimo=2)
    d["informes_iniciales_v"] = buscar(t07, r"Informes acumulados al inicio: (\d+)",
                                       "informes acumulados al inicio")
    d["ejemplo_versiones"] = bloque_entre(t07, "Ejemplo de convivencia sin perdida:",
                                          "Comprobacion:",
                                          "ejemplo de versiones").strip()
    d["esperados_v"] = buscar(t07, r"= (\d+) esperados", "informes esperados")
    d["encontrados_v"] = buscar(t07, r"encontrados:\s+(\d+)", "informes encontrados")
    # Ultimo ciclo registrado: cuantos de esos informes llevan sufijo _vN. NO
    # es lo mismo que el total acumulado, y confundirlos era exactamente el
    # error que declaraba "61 informes con sufijos" cuando con sufijo hay 40.
    d["con_sufijo_v"] = buscar_todos(
        t07, r"con sufijo de version:\s+(\d+)", "informes con sufijo de version")[-1]
    d["idempotencia"] = "\n".join(
        ln.strip() for ln in t07.splitlines()
        if re.match(r"^\s+corrida \d+ ->", ln))
    if len(d["idempotencia"].splitlines()) < 3:
        morir("se esperaban 3 corridas de idempotencia en la evidencia 07")

    # --- estado real del repositorio ---------------------------------------
    d["n_sin"] = contar_archivos(RAIZ / "gestion_ambiental" / "sin_alertas")
    d["n_con"] = contar_archivos(RAIZ / "gestion_ambiental" / "con_alertas")
    d["n_crit"] = contar_archivos(RAIZ / "gestion_ambiental" / "criticas")
    d["n_salida"] = contar_archivos(RAIZ / "salida", "informe_*.txt")
    d["n_alertas_log"] = contar_lineas(RAIZ / "alertas" / "alertas_detectadas.log")
    d["n_bitacora"] = contar_lineas(RAIZ / "logs" / "gestion_ambiental.log")
    d["bitacora_bytes"] = tam_bytes(RAIZ / "logs" / "gestion_ambiental.log")
    d["indicadores"] = []
    total_ind = 0
    for carpeta, archivo in (("temperatura", "alertas_temperatura.log"),
                             ("humedad", "alertas_humedad.log"),
                             ("pm25", "alertas_pm25.log"),
                             ("ruido", "alertas_ruido.log")):
        n = contar_lineas(RAIZ / "gestion_ambiental" / "alertas_por_indicador"
                          / carpeta / archivo)
        total_ind += n
        d["indicadores"].append((carpeta, archivo, n))
    d["total_indicadores"] = total_ind

    inv_path = RAIZ / "gestion_ambiental" / "inventario_ambiental.json"
    d["inventario"] = json.loads(leer(inv_path))
    d["inventario_bytes"] = tam_bytes(inv_path)
    d["resguardo_bytes"] = tam_bytes(RAIZ / "gestion_ambiental" / "resumen_resguardado.txt")
    d["resumen_bytes"] = tam_bytes(RAIZ / "salida" / "resumen_ambiental.txt")
    d["resguardo_identico"] = (
        (RAIZ / "salida" / "resumen_ambiental.txt").read_bytes()
        == (RAIZ / "gestion_ambiental" / "resumen_resguardado.txt").read_bytes())
    d["entradas"] = contar_archivos(RAIZ / "entrada", "*.jsonl")
    d["respaldo_parte1_real"] = len(sorted(
        (RAIZ / "evidencias" / "salida_parte1").glob("*.txt")))
    if not (RAIZ / "salida" / "LEEME.txt").exists():
        morir("falta salida/LEEME.txt, que explica por que salida/ solo trae el resumen")

    verificar_coherencia(d)
    return d


# ---------------------------------------------------------------------------
# Coherencia entre la evidencia congelada y el estado vivo del repositorio
# ---------------------------------------------------------------------------

def verificar_coherencia(d):
    """Aborta si la evidencia congelada y el arbol entregable no concuerdan.

    Motivo: un informe que mezcle una cifra leida en vivo con otra extraida de
    evidencias/ puede afirmar un par que nunca existio (por ejemplo "bitacora
    de 38 lineas y 2442 bytes", donde las lineas venian del repositorio y los
    bytes de una corrida distinta congelada en la evidencia). Aqui se exige
    que ambas fuentes digan lo mismo ANTES de emitir el documento; si difieren,
    no se emite nada.
    """
    inv = d["inventario"]
    problemas = []

    def cmp(etiqueta, vivo, evidencia):
        if str(vivo) != str(evidencia):
            problemas.append("%s: el repositorio dice %s y evidencias/ dice %s"
                             % (etiqueta, vivo, evidencia))

    # Clasificacion de informes
    cmp("informes en gestion_ambiental/sin_alertas", d["n_sin"], d["ev_n_sin"])
    cmp("informes en gestion_ambiental/con_alertas", d["n_con"], d["ev_n_con"])
    cmp("informes en gestion_ambiental/criticas", d["n_crit"], d["ev_n_crit"])
    cmp("archivos .jsonl en entrada/", d["entradas"], d["ev_entradas"])
    cmp("archivos respaldados en evidencias/salida_parte1/",
        d["respaldo_parte1_real"], d["respaldo_parte1"])

    # Bitacora: el par (lineas, bytes) tiene que venir de la MISMA corrida.
    cmp("lineas de logs/gestion_ambiental.log", d["n_bitacora"], d["ev_n_bitacora"])
    cmp("bytes de logs/gestion_ambiental.log", d["bitacora_bytes"], d["stat_size"])

    # Alertas: cinco fuentes independientes deben dar el mismo numero.
    for carpeta, _a, n in d["indicadores"]:
        cmp("alertas de " + carpeta, n, d["ev_indicadores"][carpeta])
    cmp("suma de los cuatro logs por indicador",
        d["total_indicadores"], d["ev_suma_logs"])
    cmp("lineas de alertas/alertas_detectadas.log",
        d["n_alertas_log"], d["ev_alertas_log"])
    cmp("total_alertas del inventario", inv["total_alertas"], d["ev_alertas_inventario"])
    cmp("alertas totales del resumen consolidado",
        d["resumen"]["Alertas totales"], d["ev_alertas_resumen"])
    cmp("alertas de la bitacora: secuencial contra concurrente",
        d["alertas_sec"], d["alertas_con"])

    # Cuadratura interna del inventario vivo
    cmp("informes declarados por el inventario", inv["resumen"]["total_informes"],
        d["n_sin"] + d["n_con"] + d["n_crit"])
    for cat, n in (("sin_alertas", d["n_sin"]), ("con_alertas", d["n_con"]),
                   ("criticas", d["n_crit"])):
        cmp("inventario/" + cat, inv["resumen"]["informes_por_categoria"][cat], n)
    cmp("total de alertas del inventario contra la suma por indicador",
        inv["total_alertas"], d["total_indicadores"])

    # El resumen entregado es el de la corrida concurrente; la corrida
    # secuencial (la que alimenta la tabla) debe declarar las mismas metricas.
    resumen_vivo = {}
    for ln in (RAIZ / "salida" / "resumen_ambiental.txt").read_text(
            encoding="utf-8", errors="replace").splitlines():
        if ":" in ln:
            k, v = ln.split(":", 1)
            resumen_vivo[sin_acentos(k).strip()] = limpiar(v).strip()
    for clave in ("Archivos procesados", "Lineas leidas", "Mediciones validas",
                  "Mediciones invalidas", "Alertas totales"):
        cmp("salida/resumen_ambiental.txt / " + clave,
            resumen_vivo.get(clave), d["resumen"][clave])
    cmp("tiempo interno de la corrida concurrente entregada",
        resumen_vivo.get("Tiempo total de ejecucion"),
        d["t_con"]["interno"] + " segundos")
    cmp("trabajadores de la corrida concurrente entregada",
        resumen_vivo.get("Cantidad de trabajadores"), d["t_con"]["trabajadores"])

    # salida/ debe contener solo el resumen y el LEEME (los informes los movio
    # el gestor); si aparecieran informes, salida/LEEME.txt estaria mintiendo.
    if d["n_salida"] != 0:
        problemas.append("salida/ conserva %d informes individuales, pero "
                         "salida/LEEME.txt afirma que el gestor los movio todos"
                         % d["n_salida"])
    if not d["resguardo_identico"]:
        problemas.append("salida/resumen_ambiental.txt y "
                         "gestion_ambiental/resumen_resguardado.txt difieren")

    # evidencias/comparacion_parte1.txt es prosa escrita a mano, pero cita
    # cifras de la misma corrida. Si se desincroniza, el corrector ve dos
    # documentos del entregable diciendo cosas distintas, que es exactamente
    # el defecto que se esta corrigiendo.
    comp = leer(RAIZ / "evidencias" / "comparacion_parte1.txt")
    for etiqueta, patron, esperado in (
            ("tiempo real de la secuencial",
             r"Secuencial\s*: real (\S+)", d["t_sec"]["real"]),
            ("tiempo real de la concurrente",
             r"Concurrente: real (\S+)", d["t_con"]["real"]),
            ("archivos de la carga amplificada",
             r"carga amplificada de (\d+) archivos", d["carga_archivos"]),
            ("cambios de contexto voluntarios",
             r"registro (\d+) cambios de contexto voluntarios", d["ctx_vol"]),
            ("alertas totales",
             r"Alertas detectadas\s+(\d+)", d["resumen"]["Alertas totales"]),
            ("trabajadores de la concurrente",
             r"Cantidad de trabajadores\s+1\s+(\d+)", d["t_con"]["trabajadores"])):
        m = re.search(patron, comp)
        if not m:
            problemas.append("evidencias/comparacion_parte1.txt ya no declara %s"
                             % etiqueta)
        elif m.group(1) != str(esperado):
            problemas.append(
                "evidencias/comparacion_parte1.txt dice %s = %s y "
                "evidencias/debian/ dice %s" % (etiqueta, m.group(1), esperado))

    # README.md es parte del entregable y la pauta lo evalua junto con el
    # informe ("Informe y README"). Hasta ahora este verificador NO lo leia, y
    # por eso pudo convivir un README que declaraba "du -sh entrego 528K" con
    # un informe que declaraba 624K a partir de la MISMA evidencia. Si los dos
    # documentos del entregable se contradicen, no se emite ninguno.
    readme = leer(RAIZ / "README.md")
    dfr = d["df_raiz"].split()
    met = {sin_acentos(f[0]).strip(): f[1].strip() for f in d["tabla_metricas"]}
    for etiqueta, patron, esperado in (
            ("tamano del proyecto (du -sh)",
             r"`du -sh` entrego \*\*(\S+?)\*\*", d["du_proyecto"]),
            ("tamano de la particion raiz",
             r"`/dev/sda1` con ([\d.]+ ?[KMGT]) totales", dfr[1]),
            ("espacio usado en la particion raiz",
             r"totales, ([\d.]+ ?[KMGT]) usados", dfr[2]),
            ("espacio disponible en la particion raiz",
             r"usados y ([\d.]+ ?[KMGT]) disponibles", dfr[3]),
            ("uso de la particion raiz", r"disponibles \((\d+) % de uso\)",
             dfr[4].rstrip("%")),
            ("archivos procesados", r"\| Archivos procesados \| (\d+) \|",
             d["resumen"]["Archivos procesados"]),
            ("mediciones validas", r"\| Mediciones validas \| (\d+) \|",
             d["resumen"]["Mediciones validas"]),
            ("mediciones invalidas", r"\| Mediciones invalidas \| (\d+) \|",
             d["resumen"]["Mediciones invalidas"]),
            ("alertas totales", r"\| Alertas totales \| (\d+) \|",
             d["resumen"]["Alertas totales"]),
            ("PM2.5 maximo global", r"\| PM2\.5 maximo global \| (.+?) \|",
             met["PM2.5 maximo global"]),
            ("ruido maximo global", r"\| Ruido maximo global \| (.+?) \|",
             met["Ruido maximo global"]),
            ("alertas de temperatura", r"\| Temperatura \| (\d+) \|",
             d["ev_indicadores"]["temperatura"]),
            ("alertas de humedad", r"\| Humedad \| (\d+) \|",
             d["ev_indicadores"]["humedad"]),
            ("alertas de PM2.5", r"\| PM2\.5 \| (\d+) \|",
             d["ev_indicadores"]["pm25"]),
            ("alertas de ruido", r"\| Ruido \| (\d+) \|",
             d["ev_indicadores"]["ruido"]),
            ("informes acumulados al inicio de la prueba anti-sobrescritura",
             r"Partiendo de (\d+) informes acumulados", d["informes_iniciales_v"]),
            ("informes acumulados al final de la prueba anti-sobrescritura",
             r"y luego a (\d+) \(", d["encontrados_v"]),
            ("informes con sufijo de version", r"y (\d+) de ellos con sufijo de version",
             d["con_sufijo_v"]),
            ("informes esperados en la prueba anti-sobrescritura",
             r"= (\d+) esperados", d["esperados_v"]),
            ("anomalias inyectadas", r"El gestor registro las (\d+) anomalias",
             d["anomalias_total"]),
            ("clasificacion de informes",
             r"# (\d+ / \d+ / \d+) informes",
             "%d / %d / %d" % (d["n_sin"], d["n_con"], d["n_crit"])),
    ):
        # El README lleva tildes y la evidencia no, de modo que la comparacion
        # se hace sobre el texto sin acentos. Los espacios tambien se ignoran:
        # df imprime "24G" y el README lo escribe "24 G", y esa diferencia
        # tipografica no es una contradiccion.
        m = re.search(patron, sin_acentos(limpiar(readme)))
        if not m:
            problemas.append("README.md ya no declara %s" % etiqueta)
        else:
            visto = m.group(1).strip().replace(" ", "")
            debe = str(esperado).strip().replace(" ", "")
            if visto != debe:
                problemas.append("README.md dice %s = %s y la evidencia dice %s"
                                 % (etiqueta, m.group(1).strip(), esperado))

    # Las capturas de la instalacion real deben existir.
    for nombre in FOTOS_INSTALACION_HYPERV:
        if not (FOTOS_HYPERV / nombre).exists():
            problemas.append("falta la captura de la instalacion real %s"
                             % (FOTOS_HYPERV / nombre))

    if problemas:
        morir("el estado del repositorio NO concuerda con evidencias/. No se "
              "emite el informe.\n  - " + "\n  - ".join(problemas))


# ---------------------------------------------------------------------------
# Modelo de contenido: bloques independientes del formato de salida
# ---------------------------------------------------------------------------

def h1(clave, titulo):
    """Titulo de nivel 1. El numero se asigna solo y queda registrado."""
    _CONTADOR["h1"] += 1
    _CONTADOR["h2"] = 0
    n = _registrar(clave, str(_CONTADOR["h1"]))
    return {"t": "h1", "x": n + ". " + titulo}


def h2(clave, titulo):
    """Titulo de nivel 2 dentro del ultimo nivel 1."""
    if _CONTADOR["h1"] == 0:
        morir("se pidio una subseccion antes de abrir la seccion que la contiene")
    _CONTADOR["h2"] += 1
    n = _registrar(clave, "%d.%d" % (_CONTADOR["h1"], _CONTADOR["h2"]))
    return {"t": "h2", "x": n + " " + titulo}


def p(t):
    return {"t": "p", "x": t}


def ul(items):
    return {"t": "ul", "x": list(items)}


def ol(items):
    return {"t": "ol", "x": list(items)}


def code(cmd, salida):
    return {"t": "code", "cmd": cmd, "x": salida}


def tabla(cabecera, filas, anchos=None):
    return {"t": "tabla", "head": cabecera, "rows": filas, "w": anchos}


# Numero de la ultima figura emitida. Igual que con las secciones, el numero
# NO se escribe a mano en el pie: lo asigna fig() por orden de aparicion. Asi,
# quitar o agregar una figura no puede dejar el resto de la numeracion mal.
_N_FIG = {"n": 0}


def fig(ruta, pie, alto=None, texto_pt=None):
    """Figura del informe. El pie se pasa SIN el "Figura N." inicial.

    Por omision la figura se coloca a ancho de columna completo, que es lo que
    corresponde a una captura de consola: una transcripcion de terminal
    reducida a un tercio de su tamano nativo deja de ser evidencia, porque no
    se puede leer.

    alto es un TOPE en centimetros. texto_pt fija a cuantos puntos debe quedar
    el texto de dentro de la captura; sirve para las figuras que no son
    transcripciones de terminal (una pantalla de menu, por ejemplo), donde
    ocupar media pagina no aporta nada.
    """
    _N_FIG["n"] += 1
    return {"t": "fig", "ruta": str(ruta), "alto": alto, "texto_pt": texto_pt,
            "n": _N_FIG["n"], "pie": "Figura %d. %s" % (_N_FIG["n"], pie)}


def ancho_de_figura(ruta, ancho_util, texto_pt=None):
    """Ancho con que colocar una figura, en las unidades de ancho_util.

    Se parte del ancho de columna completo y solo se reduce si el texto de
    dentro de la captura fuese a salir mas grande de lo necesario.
    """
    from PIL import Image
    with Image.open(ruta) as im:
        ancho_px = im.size[0]
    linea_px = altura_de_linea_px(ruta)
    objetivo = texto_pt if texto_pt is not None else TEXTO_MAXIMO_PT
    if linea_px is None:
        return ancho_util
    deseado = ancho_px * (objetivo / float(linea_px))
    factor = ancho_util / 481.89  # el ancho util del PDF, en las mismas unidades
    return min(ancho_util, deseado * factor)


def n_figuras():
    return _N_FIG["n"]


# Colocacion final de cada figura en el PDF. La llena render_pdf() y main() la
# imprime, para poder demostrar que ninguna captura quedo ilegible.
#
# El criterio de legibilidad NO es la escala: una captura de 1600x1200 al 31 %
# se lee perfectamente y una de 640x400 al 60 % puede no leerse. Lo que decide
# es a cuantos puntos queda el TEXTO de dentro de la figura una vez colocada.
# El umbral es 6.0 pt, la altura de la letra de los bloques de transcripcion de
# este mismo informe (Courier 6.6 pt): por encima de eso, lo que se lee en el
# cuerpo del informe se lee tambien dentro de la captura.
ESCALAS = []
TEXTO_MINIMO_PT = 6.0

# Tope por arriba: ninguna captura necesita que su texto salga mas grande que
# el del propio informe. Una captura recortada a 544x294 colocada a ancho de
# columna completo deja su letra en 10.6 pt, mas grande que el cuerpo del
# documento (9 pt) y casi el doble que sus bloques de transcripcion (6.6 pt);
# ese exceso solo gasta pagina. Las capturas se colocan a ancho de columna
# completo salvo que eso pase de este tope.
TEXTO_MAXIMO_PT = 8.5


def nota(t):
    return {"t": "nota", "x": t}


def salto():
    return {"t": "salto"}


def construir_documento(d, imgs_hv):
    inv = d["inventario"]
    res = d["resumen"]
    hv = d["hv"]
    ind = {c: n for c, _a, n in d["indicadores"]}
    b = []

    # ---------------------------------------------------------------- Portada
    b.append({"t": "portada"})
    b.append(salto())

    # ------------------------------------------- 1. Introduccion y objetivos
    b.append(h1("intro", "Introduccion y objetivos"))
    b.append(p(
        "Este informe documenta el traslado de la solucion de monitoreo ambiental del "
        "Equipo 07 a un ambiente Debian GNU/Linux, como exige la Parte 2 de la Solemne 01 "
        "practica (Forma B): instalacion de la maquina virtual, preparacion del entorno, "
        "ejecucion de las dos versiones de la Parte 1 sobre el mismo conjunto de datos, "
        "gestor de incidencias y observacion de procesos, memoria y sistema de archivos. "
        "Todas las cifras se midieron dentro de la maquina virtual en una UNICA corrida y "
        "quedaron en evidencias/. El documento lo genera scripts/generar_informe.py, que "
        "lee esos archivos, los compara con el estado real del arbol entregado y se "
        "detiene sin emitir nada si ambas fuentes no coinciden."))

    b.append(h2("objetivos", "Objetivo de cada componente"))
    b.append(ul([
        "src/secuencial.py: procesar un archivo JSONL completo antes de pasar al "
        "siguiente, sin hilos. Define el contrato de validacion y es la referencia contra "
        "la cual se compara el resultado concurrente.",
        "src/concurrente.py: procesar el mismo conjunto repartiendo los archivos entre "
        + d["t_con"]["trabajadores"] + " trabajadores threading.Thread mediante una "
        "queue.Queue compartida, protegiendo acumuladores globales y bitacora de alertas "
        "con mutex threading.Lock.",
        "src/gestor_incidencias.py: clasificar los informes por cantidad de alertas, "
        "resguardar el resumen, separar alertas por indicador, generar el inventario JSON "
        "y mantener una bitacora trazable sin detenerse ante entradas defectuosas.",
    ]))

    b.append(h2("flujo", "Flujo del sistema de principio a fin"))
    b.append(ol([
        "Generacion de datos: scripts/generar_archivos_entrada.py crea "
        + d["archivos_generados"] + " archivos estacion_CODIGO_AAAAMMDD.jsonl en "
        "entrada/ con semilla fija " + d["semilla"] + ", de modo que la entrada es "
        "identica en cualquier maquina.",
        "Ingesta y validacion: cada linea JSONL se valida en formato, tipos, rangos y "
        "timestamp; la que no cumple se descarta como invalida sin detener el proceso.",
        "Procesamiento: la version secuencial recorre los archivos uno tras otro; la "
        "concurrente los encola en una queue.Queue y los reparte entre "
        + d["t_con"]["trabajadores"] + " hilos que acumulan resultados locales y solo "
        "actualizan las variables globales dentro de una seccion critica.",
        "Deteccion de alertas: cada medicion valida se contrasta con los umbrales y la "
        "alerta se escribe en alertas/alertas_detectadas.log con el formato "
        "archivo;estacion;timestamp;indicador;valor;umbral.",
        "Salida: un informe por archivo en salida/informe_*.txt y el consolidado "
        "salida/resumen_ambiental.txt.",
        "Gestion de incidencias: los informes se MUEVEN a gestion_ambiental/sin_alertas/, "
        "con_alertas/ o criticas/ segun sus alertas; el resumen se COPIA como "
        "resumen_resguardado.txt; las alertas se separan por indicador; se generan "
        "inventario_ambiental.json y logs/gestion_ambiental.log. Por eso salida/ queda "
        "solo con el resumen: ver " + ref("leeme") + ".",
    ]))

    b.append(h2("validacion", "Reglas de validacion y umbrales de alerta"))
    b.append(tabla(
        ["Campo", "Rango valido", "Indicador", "Condicion de alerta"],
        [[RANGOS_VALIDOS[i][0], RANGOS_VALIDOS[i][1],
          UMBRALES[i][0] if i < len(UMBRALES) else "",
          UMBRALES[i][1] if i < len(UMBRALES) else ""]
         for i in range(len(RANGOS_VALIDOS))],
        anchos=[3.0, 5.0, 3.0, 5.5]))
    b.append(p(
        "Ambas versiones comparten estas constantes, y esa es la razon de fondo por la "
        "cual sus metricas deben coincidir hasta el ultimo decimal: la concurrencia "
        "cambia el orden en que se hace el trabajo, no el criterio con que se decide."))

    # ------------------------------------------------- 2. Instalacion Debian
    b.append(h1("instalacion", "Instalacion de Debian 13 en la maquina virtual"))

    b.append(h2("iso", "Descarga y verificacion del ISO"))
    b.append(p(
        "Se descargo la imagen de instalacion por red " + hv["iso_nombre"] + " ("
        + hv["iso_tamano"] + ") desde el sitio oficial de Debian y se verifico su "
        "integridad comparando su SHA256 con el publicado en SHA256SUMS, lo que garantiza "
        "que el medio no fue alterado ni quedo truncado. La salida es la del anfitrion y "
        "esta archivada en evidencias/hyperv/configuracion-vm-hyperv.txt:"))
    b.append(code("PS> Get-FileHash " + hv["iso_nombre"] + " -Algorithm SHA256",
                  "Hash calculado : " + hv["iso_sha_calculado"]
                  + "\nHash publicado : " + hv["iso_sha_publicado"]
                  + "\n  Fuente: " + hv["iso_fuente"]
                  + "\nCoincide       : " + hv["iso_coincide"]))

    b.append(h2("hipervisor", "Declaracion del cambio de hipervisor respecto del hito"))
    b.append(nota(
        "Declaracion explicita. En el hito el equipo planifico usar "
        + HIPERVISOR_PLANIFICADO + " y la implementacion final se hizo sobre "
        + hv["hipervisor_implementado"] + ". Se hace constar porque afecta lo declarado en "
        "el hito; la pauta admite la situacion, ya que la instalacion \"puede hacerlo en "
        "VirtualBox, VMware u otro hipervisor\". Los recursos comprometidos se respetaron "
        "exactamente, sin rebajar ninguno."))
    b.append(tabla(
        ["Recurso", "Minimo de la pauta", "Comprometido en el hito",
         "Implementado y verificado"],
        [["Hipervisor", "VirtualBox, VMware u otro", HIPERVISOR_PLANIFICADO,
          "Microsoft Hyper-V, Generacion " + hv["generacion"]],
         ["RAM", "2 GB", "4 GB",
          hv["ram_gb"] + " GB fijos (memoria dinamica: " + hv["ram_dinamica"] + "); "
          + d["memoria"].splitlines()[1].split()[1] + " utiles segun free -h"],
         ["Procesadores virtuales", "2", "4",
          hv["vcpu"] + " vCPU sobre " + d["modelo_cpu"]],
         ["Disco virtual", "20 GB", "25 GB",
          hv["vhd_gb"] + " GB " + hv["vhd_tipo"] + " (ocupa " + hv["vhd_real_gb"]
          + " GB reales en el anfitrion)"],
         ["Red", "NAT", "NAT con conectividad",
          hv["switch"] + " (" + hv["switch_tipo"] + "), IP " + d["ip_eth0"] + " por DHCP"],
         ["Sistema", "Debian 13, 64 bits", "Debian 13, 64 bits",
          d["pretty_name"] + ", version " + d["version_debian"]]],
        anchos=[2.5, 2.9, 2.9, 8.1]))
    b.append(p(
        "El propio sistema instalado confirma sobre que hipervisor corre, de modo que la "
        "declaracion es verificable y no depende de la palabra del equipo:"))
    b.append(code("$ systemd-detect-virt ; lscpu | grep 'Hypervisor vendor' ; uname -a",
                  d["hipervisor_detectado"]
                  + "\nHypervisor vendor:                       " + d["hipervisor_vendor"]
                  + "\n" + d["kernel"]))
    # La captura de la creacion de la VM en VirtualBox NO se incluye como
    # figura. No corresponde a la maquina entregada, y su propio contenido lo
    # delata: declara otro usuario y otro nombre de host que los de la maquina
    # definitiva. Ilustrar la configuracion de recursos con una pantalla de una
    # maquina distinta seria peor que no ilustrarla, porque los valores que
    # importan ya estan verificados arriba contra los cmdlets del hipervisor.
    # El archivo sigue en el repositorio para quien quiera revisar la etapa
    # previa: evidencias/fotos/instalacion/02-recursos-vm.png.
    b.append(p(
        "La etapa previa sobre " + HIPERVISOR_PLANIFICADO + " queda archivada en "
        "evidencias/fotos/instalacion/, donde consta la creacion de la maquina con esos "
        "mismos recursos. No se reproduce como figura porque no es la maquina entregada y "
        "sus valores ya estan verificados arriba contra los cmdlets del hipervisor."))

    b.append(h2("instalador", "Instalacion, usuario y particionado del disco virtual"))
    b.append(p(
        "Las capturas de esta seccion y de la siguiente son de la instalacion REAL de la "
        "maquina definitiva sobre Hyper-V. Se creo de Generacion " + hv["generacion"]
        + ", con firmware BIOS heredado, y su orden de arranque ("
        + hv["orden_arranque"] + ") explica que el instalador arranque en modo BIOS y no "
        "UEFI. Se creo el usuario sin privilegios " + d["usuario_vm"] + ", con acceso a "
        "sudo, y el host quedo como " + d["hostname_vm"] + ", de modo que cualquier salida "
        "de este informe se atribuye sin ambiguedad a la maquina del equipo."))
    b.append(fig(imgs_hv["01-menu-instalador-debian13.png"],
                 "Menu del instalador de Debian 13 en modo BIOS, en la maquina "
                 "de Generacion " + hv["generacion"] + " de Hyper-V. Instalacion real de "
                 "la maquina definitiva.", texto_pt=7.0))
    b.append(fig(imgs_hv["02-linea-de-arranque-preseed.png"],
                 "Linea de arranque del instalador con el archivo de "
                 "preconfiguracion (preseed), que fija idioma, teclado, zona horaria, "
                 "particionado guiado y el conjunto minimo de tareas. Usar preseed hace "
                 "la instalacion reproducible: el mismo archivo produce la misma "
                 "maquina."))
    b.append(fig(imgs_hv["03-instalacion-sistema-base.png"],
                 "Instalacion del sistema base: descarga e instalacion de paquetes desde "
                 "la replica de red. La etapa final del instalador -gestor de arranque y "
                 "copia de la configuracion de red al sistema instalado- es la misma "
                 "pantalla de progreso y esta archivada en "
                 "evidencias/fotos/instalacion-hyperv/04-instalacion-final.png."))
    b.append(p("El particionado se aplico sobre el disco de " + hv["vhd_gb"]
               + " GB con el esquema guiado \"todo en una particion\":"))
    b.append(code("$ lsblk", d["lsblk"]))
    b.append(ul([
        "sda1 (23.7 GB, ext4) es la particion raiz montada en /; sda2 es solo el "
        "contenedor logico de la extendida y sda5 (1.3 GB) es el area de intercambio, "
        "que el nucleo usa para descargar paginas cuando la memoria fisica escasea.",
        "sr0 es la unidad optica virtual desde la que se monto el ISO de instalacion; "
        "terminada la instalacion el ISO fue expulsado, como consta en la salida de "
        "Get-VMDvdDrive del anfitrion.",
    ]))

    b.append(h2("primer_inicio", "Primer inicio y conectividad"))
    b.append(fig(imgs_hv["05-primer-inicio.png"],
                 "Primer inicio del sistema recien instalado: el nucleo llega a "
                 "la consola tty1 y presenta el indicador de acceso del equipo "
                 + d["hostname_vm"] + ". La captura muestra el indicador de acceso, no la "
                 "sesion ya iniciada; que el acceso con el usuario " + d["usuario_vm"]
                 + " funciona lo demuestra el resto del informe, cuyas salidas fueron "
                 "obtenidas en esa sesion."))
    b.append(p(
        "eth0 obtiene su direccion por DHCP desde el conmutador " + hv["switch"]
        + " y la ruta por omision apunta a " + d["gateway"] + ". " + hv["nat"]
        + " El anfitrion ve la misma direccion (" + hv["ip_hyperv"] + ", MAC "
        + hv["mac"] + "), lo que cierra la trazabilidad entre lo que declara Hyper-V y lo "
        "que observa el huesped. Con esa configuracion la maquina alcanza las replicas de "
        "Debian, como confirman la resolucion de deb.debian.org y las lineas \"Hit\" de "
        "apt de la seccion siguiente."))
    b.append(code("$ ip -4 -br addr show ; ip route ; getent hosts deb.debian.org",
                  d["red"] + "\n" + d["dns"]))

    # --------------------------------------------- 3. Preparacion del ambiente
    b.append(h1("ambiente", "Preparacion del ambiente"))
    b.append(p(
        "El punto 2.a de la pauta exige actualizar el sistema antes de trabajar. La "
        "salida confirma que los tres origenes de paquetes estan accesibles (trixie, "
        "trixie-security y trixie-updates) y que el sistema quedo al dia."))
    # Se transcriben las lineas evaluables (origenes accesibles y estado final)
    # y se marca la elision del resto con [...], como en el resto del informe.
    b.append(code("$ sudo apt update ; sudo apt upgrade -y",
                  "\n".join(ln for ln in d["apt_update"].splitlines()
                            if ln.startswith(("Hit:", "All packages")))
                  + "\n\n[...]\n"
                  + "\n".join(d["apt_upgrade"].splitlines()[-3:])))
    b.append(p(
        "El punto 2.b pide instalar el entorno del lenguaje usado en la Parte 1. El "
        "proyecto esta en Python y Debian 13 ya trae " + d["python_version"] + " en el "
        "sistema base, de modo que no hubo que compilar ni agregar repositorios externos. "
        + d["sin_dependencias"] + ": sin pip, sin entorno virtual y sin archivo de "
        "requisitos, decision deliberada porque hace reproducible la ejecucion incluso sin "
        "red. " + d["compilacion"]))
    # OJO: d["compilacion"] es una frase de prosa del archivo de evidencia, no
    # una linea que haya escrito ninguna terminal. Va en el parrafo anterior y
    # NO dentro del bloque de transcripcion: un bloque presentado como salida
    # de consola solo puede contener lo que la consola imprimio. py_compile,
    # cuando todo compila, no imprime nada.
    b.append(code("$ python3 --version ; apt list --installed | "
                  "grep -E 'python3|openssh-server|time'\n"
                  "$ python3 -m py_compile src/*.py scripts/*.py   "
                  "# sin salida: los cuatro modulos compilan",
                  d["python_version"] + "\n" + d["apt_list"]))

    # ------------------------------------------ 4. Ejecucion de la Parte 1
    b.append(h1("parte1", "Ejecucion de la Parte 1 en Debian"))
    b.append(p(
        "El punto 2.c exige ejecutar ambas versiones sobre el mismo conjunto de archivos "
        ".jsonl y comprobar que seis metricas coinciden. Para que \"el mismo conjunto\" sea "
        "verificable, la entrada se genera con semilla fija y se le calcula una suma de "
        "verificacion antes de cada corrida."))
    b.append(code("$ cd " + d["dir_trabajo"] + "\n"
                  "$ python3 scripts/generar_archivos_entrada.py\n"
                  "$ cat entrada/*.jsonl | sha256sum\n"
                  "$ time python3 src/secuencial.py\n"
                  "$ time python3 src/concurrente.py",
                  "Archivos .jsonl generados: " + d["archivos_generados"]
                  + "\n" + d["checksum_entrada"] + "  -"
                  + "\n\nsecuencial   real " + d["t_sec"]["real"]
                  + "   user " + d["t_sec"]["user"] + "   sys " + d["t_sec"]["sys"]
                  + "   (tiempo interno " + d["t_sec"]["interno"] + " s, "
                  + d["t_sec"]["trabajadores"] + " trabajador)"
                  + "\nconcurrente  real " + d["t_con"]["real"]
                  + "   user " + d["t_con"]["user"] + "   sys " + d["t_con"]["sys"]
                  + "   (tiempo interno " + d["t_con"]["interno"] + " s, "
                  + d["t_con"]["trabajadores"] + " trabajadores)"))

    b.append(h2("metricas", "Las seis metricas exigidas por la pauta"))
    b.append(tabla(["Metrica", "Secuencial", "Concurrente", "Coincide"],
                   d["tabla_metricas"], anchos=[5.0, 4.0, 4.0, 3.5]))
    b.append(p("Metricas comparadas: " + d["metricas_comparadas"] + ". "
               + d["resultado_consistencia"] + ". El resumen consolidado agrega "
               + res["Lineas leidas"] + " lineas leidas, temperatura promedio "
               + res.get("Temperatura promedio global", "?") + ", humedad promedio "
               + res.get("Humedad promedio global", "?") + " y estacion con mas alertas "
               + res.get("Estacion con mayor cantidad de alertas", "?")
               + ". El entregable conserva el resumen de la corrida concurrente en "
               "salida/resumen_ambiental.txt."))

    b.append(h2("consistencia", "Comprobaciones adicionales de consistencia"))
    b.append(p(
        "Seis totales podrian coincidir por casualidad si dos errores se compensaran, de "
        "modo que se agregaron dos comprobaciones mas exigentes que las que pide la pauta: "
        "los " + d["informes_comparados"] + " informes individuales comparados uno por "
        "uno, con " + d["informes_con_diferencias"] + " diferencias, y el contenido "
        "completo de la bitacora de alertas, con " + d["alertas_sec"] + " alertas en el "
        "secuencial y " + d["alertas_con"] + " en el concurrente. El orden de escritura "
        "del concurrente varia por diseno, porque " + d["t_con"]["trabajadores"]
        + " hilos vuelcan sus buferes en momentos distintos; por eso el volcado final se "
        "ordena por nombre de archivo. " + d["veredicto_log"]))
    b.append(nota(
        "Resultado incomodo que se declara tal cual se midio: la concurrente fue mas "
        "LENTA que la secuencial en este conjunto (" + d["t_sec"]["real"] + " frente a "
        + d["t_con"]["real"] + " de tiempo real; " + d["t_sec"]["interno"] + " s frente a "
        + d["t_con"]["interno"] + " s de tiempo interno). La explicacion esta en "
        + ref("concurrencia") + "."))
    # El pie de esta figura NO se escribe a ciegas: se comprueba que la captura
    # muestre de verdad las filas de la tabla. Si la tabla salio vacia en la
    # pantalla (ha ocurrido), el pie no puede prometer "las seis metricas
    # coinciden", porque la imagen no lo muestra; la afirmacion sigue
    # respaldada por la tabla del cuerpo, y el pie lo dice asi.
    captura_7 = recortar_consola(FOTOS / "debian-01-consistencia-parte1.png")
    filas_7 = filas_visibles_de_la_tabla(captura_7)
    if filas_7 is not None and filas_7 >= 4:
        pie_7 = ("Comprobacion de consistencia en la consola de la maquina "
                 "virtual: las seis metricas coinciden, los " + d["informes_comparados"]
                 + " informes individuales son identicos y la bitacora de alertas tiene "
                 + d["alertas_sec"] + " alertas en ambas versiones.")
    else:
        pie_7 = ("Comprobacion de consistencia ejecutada en la consola de la "
                 "maquina virtual. En esta captura el cuerpo de la tabla comparativa no "
                 "alcanzo a quedar en pantalla: se ven el encabezado, los separadores y "
                 "el resultado final (los " + d["informes_comparados"] + " informes "
                 "individuales identicos y las " + d["alertas_sec"] + " alertas de la "
                 "bitacora en ambas versiones). La comparacion metrica por metrica es la "
                 "de la tabla de " + ref("metricas") + ", tomada de "
                 "evidencias/debian/03-ejecucion-parte1.txt.")
    # Esta captura es un recorte pequeno (544x294): a ancho de columna completo
    # se ampliaria hasta dejar su letra por encima de la del cuerpo del informe
    # y, sobre todo, desbordaria la pagina dejando un tercio en blanco. Se
    # coloca al tamano en que su texto queda a 7.2 pt, por encima del umbral.
    b.append(fig(captura_7, pie_7, texto_pt=7.2))

    # ----------------------------------------------- 5. Gestor de incidencias
    b.append(h1("gestor", "Gestor de incidencias"))
    b.append(p(
        "El punto 2.d enumera trece requisitos para src/gestor_incidencias.py. La tabla "
        "los recorre uno por uno con la evidencia que los respalda; todas las cifras "
        "provienen del estado real del proyecto y de los archivos de evidencia, y el "
        "generador comprueba que ambas fuentes coincidan antes de escribir esta pagina."))
    b.append(tabla(
        ["N", "Requisito de la pauta", "Evidencia obtenida"],
        [["1", "Detectar informes informe_CODIGO_AAAAMMDD.txt",
          "Patron ^informe_[A-Za-z0-9]+_\\d{8}(_v\\d+)?\\.txt$; "
          + str(inv["resumen"]["total_informes"]) + " informes detectados"],
         ["2", "Leer la cantidad de alertas de cada informe",
          "Linea \"Alertas detectadas: N\" de cada informe; suma "
          + str(inv["total_alertas"]) + " alertas"],
         ["3", "Mover informes con 0 alertas a sin_alertas/",
          "Rama implementada; queda en " + str(d["n_sin"]) + " informes (ver "
          + ref("sin_alertas") + ") y se demostro con " + d["ctl_informe"]],
         ["4", "Mover informes con 1 a 3 alertas a con_alertas/",
          str(d["n_con"]) + " informes en gestion_ambiental/con_alertas/"],
         ["5", "Mover informes con 4 o mas alertas a criticas/",
          str(d["n_crit"]) + " informes en gestion_ambiental/criticas/"],
         ["6", "Evitar sobrescritura mediante nombres _v1, _v2",
          "Tres corridas acumulan " + d["encontrados_v"] + " informes, "
          + d["con_sufijo_v"] + " de ellos con sufijo _v1 o _v2, sin perdidas (ver "
          + ref("anti") + ")"],
         ["7", "Copiar el resumen como resumen_resguardado.txt",
          "Copia de " + str(d["resguardo_bytes"])
          + " bytes, identica al original que permanece en salida/"],
         ["8", "Leer alertas/alertas_detectadas.log",
          str(d["n_alertas_log"]) + " lineas del formato "
          "archivo;estacion;timestamp;indicador;valor;umbral"],
         ["9", "Separar alertas por indicador",
          "temperatura " + str(ind["temperatura"]) + ", humedad " + str(ind["humedad"])
          + ", pm25 " + str(ind["pm25"]) + ", ruido " + str(ind["ruido"])],
         ["10", "Registrar alertas invalidas o desconocidas",
          d["anomalias_total"] + " anomalias registradas al inyectar fallas (ver "
          + ref("errores") + ")"],
         ["11", "Generar inventario_ambiental.json",
          "Archivo de " + str(d["inventario_bytes"]) + " bytes, "
          + d["json_valido"] + " con " + d["n_claves_inventario"]
          + " claves de primer nivel"],
         ["12", "Generar logs/gestion_ambiental.log",
          "Bitacora de " + d["ev_n_bitacora"] + " lineas y " + d["stat_size"]
          + " bytes, con marca de tiempo por evento"],
         ["13", "Controlar al menos un error sin detenerse",
          "Con cuatro fallas el gestor termino con codigo " + d["gestor_exit_error"]
          + " y conservo las " + d["inventario_tras_error"] + " alertas (ver "
          + ref("errores") + ")"]],
        anchos=[0.8, 6.2, 9.5]))

    b.append(h2("estructura", "Estructura generada"))
    b.append(code("$ find gestion_ambiental logs -type d | sort",
                  d["find_dirs"]))
    b.append(p("Bajo gestion_ambiental/ quedaron " + d["find_files_total"]
               + " archivos: los " + str(inv["resumen"]["total_informes"])
               + " informes clasificados, los cuatro registros por indicador, el "
               "inventario y el resumen resguardado. El gestor termino con codigo "
               + d["gestor_exit"] + " en su unica ejecucion sobre el arbol entregado."))
    b.append(fig(recortar_consola(FOTOS / "debian-03-estructura-gestor.png"),
                 "Estructura generada por el gestor y cuadratura de las alertas por "
                 "indicador, en la maquina virtual."))

    b.append(h2("cuadratura", "Inventario y cuadratura de las alertas"))
    b.append(p(
        "El valor del inventario no esta en existir sino en cuadrar: la suma de las "
        "alertas por indicador debe ser igual al total del inventario, al numero de lineas "
        "de la bitacora de alertas y al total del resumen consolidado. El generador "
        "verifica esa igualdad y ademas que el numero leido en el repositorio sea el mismo "
        "de la evidencia congelada; si alguna comprobacion falla, se detiene."))
    b.append(tabla(
        ["Fuente", "Alertas", "Fuente", "Alertas"],
        [["alertas_por_indicador/temperatura", str(ind["temperatura"]),
          "Suma de los cuatro indicadores", str(d["total_indicadores"])],
         ["alertas_por_indicador/humedad", str(ind["humedad"]),
          "alertas/alertas_detectadas.log", str(d["n_alertas_log"])],
         ["alertas_por_indicador/pm25", str(ind["pm25"]),
          "inventario_ambiental.json", str(inv["total_alertas"])],
         ["alertas_por_indicador/ruido", str(ind["ruido"]),
          "salida/resumen_ambiental.txt", res["Alertas totales"]]],
        anchos=[5.6, 1.8, 5.6, 1.8]))
    b.append(code("$ python3 -m json.tool gestion_ambiental/inventario_ambiental.json",
                  json.dumps(inv["resumen"], indent=2, ensure_ascii=False)
                  + "\n...\n"
                  + "\n".join(d["bitacora_final"].splitlines()[-3:])))

    b.append(h2("sin_alertas", "Por que sin_alertas/ quedo vacia"))
    b.append(p(
        "La carpeta sin_alertas/ contiene " + str(d["n_sin"]) + " informes y no es un "
        "defecto del gestor: la pauta de la Parte 1 exige que cada archivo de entrada "
        "produzca al menos una alerta, de modo que ningun informe puede tener cero y la "
        "rama queda legitimamente vacia. La rama existe y funciona: con el informe de "
        "control " + d["ctl_informe"] + ", construido con cero alertas, el gestor lo movio "
        "a sin_alertas/ y la clasificacion quedo en " + d["clasificacion_ctl"] + ". Esa "
        "demostracion se hizo sobre una COPIA, por lo que el arbol entregado conserva sus "
        + str(d["n_sin"]) + " informes en esa rama."))

    b.append(h2("leeme", "Por que salida/ solo conserva el resumen"))
    b.append(p(
        "En el arbol entregado, salida/ contiene solo resumen_ambiental.txt y LEEME.txt. "
        "No falta nada: los " + d["informes_generados"] + " informes individuales SI se "
        "generaron ahi (src/concurrente.py) y el gestor los MOVIO a sus carpetas de "
        "clasificacion, que es lo que exigen los puntos 3, 4 y 5 de la pauta. Mover y no "
        "copiar se explica en " + ref("mover_copiar") + ". Antes de ejecutar el gestor se "
        "respaldo salida/ completo, de modo que el estado que dejo la Parte 1 tambien es "
        "auditable:"))
    b.append(ul([
        "salida/resumen_ambiental.txt (" + str(d["resumen_bytes"]) + " bytes): el "
        "consolidado que la pauta pide mantener en su lugar.",
        "salida/LEEME.txt: explica esta situacion dentro del propio entregable e indica "
        "como regenerar los informes.",
        "evidencias/salida_parte1/: " + d["respaldo_parte1"] + " archivos, es decir los "
        + d["informes_generados"] + " informes individuales mas el resumen, tal como los "
        "dejo la Parte 1 antes de la clasificacion.",
    ]))

    # -------------------------------- 6. Observacion de procesos y archivos
    b.append(h1("procesos", "Observacion de procesos y del sistema de archivos"))
    b.append(nota(
        "Nota metodologica sobre la carga amplificada. El conjunto oficial de "
        + d["archivos_generados"] + " archivos se procesa en milisegundos y ps no alcanza "
        "a tomar una muestra util. Por eso la observacion se hizo en dos partes: (A) el "
        "consumo de la corrida OFICIAL con /usr/bin/time -v, que no necesita muestreo; y "
        "(B) la observacion en vivo con ps sobre una CARGA AMPLIFICADA de "
        + d["carga_archivos"] + " archivos (" + d["carga_lineas"] + " lineas), los mismos "
        "archivos oficiales replicados con otras fechas validas. " + d["copias_demo"]
        + " Toda metrica entregable sale del conjunto oficial; la carga amplificada solo "
        "hace que el proceso viva lo suficiente para observarlo ("
        + d["muestras_ps"] + " muestras)."))

    b.append(h2("recursos", "Consumo de recursos de la corrida oficial"))
    b.append(code("$ /usr/bin/time -v python3 src/concurrente.py", d["time_v"]))
    b.append(ul([
        "Percent of CPU " + d["pct_cpu_oficial"] + " sobre centesimas de segundo: con "
        "solo " + d["archivos_generados"] + " archivos el proceso apenas alcanza a "
        "repartir trabajo antes de terminar.",
        "Maximum resident set size " + d["max_rss"] + " KB (unos "
        + "%.1f" % (int(d["max_rss"]) / 1024.0) + " MB de memoria fisica): el costo del "
        "interprete de Python mas los acumuladores del programa.",
        d["ctx_vol"] + " cambios de contexto voluntarios y " + d["ctx_invol"]
        + " involuntarios. Los voluntarios son el proceso cediendo la CPU por su cuenta "
        "al esperar entrada/salida o un mutex; los involuntarios son el planificador "
        "expulsandolo al agotar su cuanto. " + d["fs_outputs"] + " bloques escritos "
        "corresponden a los informes, el resumen y la bitacora de alertas.",
    ]))

    b.append(h2("identificacion", "Identificacion del proceso concurrente"))
    b.append(code(d["ps_cmd"], d["ps_cabecera"] + "\n" + d["ps_linea"]))
    b.append(tabla(
        ["Campo", "Valor", "Que significa en este proceso concreto"],
        [["PID", d["pid"],
          "Identificador que el nucleo asigno al proceso que ejecuta "
          "python3 src/concurrente.py; es unico mientras el proceso vive."],
         ["PPID", d["ppid"],
          "Proceso padre: " + d["padre"].splitlines()[-1].strip()
          + ". Es el shell desde el que se lanzo la ejecucion."],
         ["STAT", d["stat"],
          "S: en el instante de la muestra el proceso dormia en espera interrumpible "
          "(entrada/salida o mutex); la l final significa multihilo. En otras muestras "
          "aparece R, ejecutandose en CPU."],
         ["%CPU", d["pcpu"] + " (max. " + d["pcpu_max"] + ")",
          "Que supere el 100 % es la prueba de trabajo simultaneo en mas de un nucleo: "
          "un proceso de un solo hilo no puede pasar de 100 %."],
         ["%MEM", d["pmem"],
          "Fraccion de los " + hv["ram_gb"] + " GB de memoria fisica ocupada por el "
          "proceso; es baja porque el programa procesa por flujo y no carga todo en "
          "memoria."],
         ["RSS", d["rss"] + " KB",
          "Memoria residente: la parte del proceso efectivamente cargada en RAM, que es "
          "lo que cuesta de verdad en memoria fisica."],
         ["VSZ", d["vsz"] + " KB",
          "Memoria virtual reservada, unas trece veces el RSS: incluye bibliotecas "
          "compartidas, pilas de hilos y regiones que nunca se tocaron, por lo que no "
          "debe leerse como consumo real."],
         ["NLWP", d["nlwp"],
          "Hilos vivos: 1 principal + " + d["t_con"]["trabajadores"]
          + " trabajadores, exactamente los que declara el programa."]],
        anchos=[1.5, 2.6, 12.4]))

    b.append(h2("hilos", "Hilos y vista del nucleo"))
    b.append(code("$ ps -L -o pid,tid,stat,%cpu,comm -p " + d["pid"] + "\n"
                  "$ grep -E 'State|Threads|VmRSS' /proc/" + d["pid"] + "/status",
                  d["ps_hilos"] + "\n" + d["proc_status"]))
    b.append(p(
        "Los cuatro identificadores de hilo comparten el mismo PID pero tienen TID "
        "distintos: el primero coincide con el PID y es el hilo principal, que encola y "
        "espera; los otros " + d["t_con"]["trabajadores"] + " consumen de la cola. El "
        "nucleo confirma Threads: " + d["proc_threads"] + " y VmRSS: " + d["proc_vmrss"]
        + " kB, coherente con ps, y State: " + d["proc_estado"] + ". Que el estado sea "
        "\"sleeping\" en una muestra y \"running\" en otra no es contradictorio: los hilos "
        "alternan entre calculo y espera por entrada/salida."))
    b.append(code("$ ps -o pid,ppid,stat,%cpu,%mem,rss,vsz,nlwp,cmd -p <PID>   "
                  "# muestras sucesivas", d["ps_evolucion"]))
    b.append(p(
        "La primera muestra fue tomada antes de que arrancaran los trabajadores: un solo "
        "hilo, 0.0 % de CPU y unos 7 MB residentes. En las siguientes NLWP sube a "
        + d["nlwp"] + ", el %CPU pasa de 100 y el RSS crece de forma sostenida a medida "
        "que se llenan los acumuladores. Esa progresion es la concurrencia hecha visible."))
    b.append(fig(recortar_consola(FOTOS / "debian-02-observacion-procesos.png"),
                 "Observacion en vivo del proceso concurrente en la consola de la "
                 "maquina virtual. Es otra corrida de la misma prueba, con su propia "
                 "carga amplificada -la que declara el encabezado de la pantalla, "
                 "distinta de la transcrita arriba- y por eso muestra otro PID y otro "
                 "RSS; coincide en estructura y en orden de magnitud: un proceso con "
                 + d["nlwp"] + " hilos vivos, %CPU por encima de 100 y el mismo "
                 "du -sh de " + d["du_proyecto"] + " del proyecto."))

    b.append(h2("memoria_fs",
                "Memoria del sistema, espacio disponible y tamano del proyecto"))
    mem = d["free_h"].splitlines()[1].split()
    dfr = d["df_raiz"].split()
    b.append(code("$ free -h ; df -h / ; du -sh ~/solemne_so_equipo07",
                  d["free_h"] + "\n\n"
                  + "Filesystem      Size  Used Avail Use% Mounted on\n" + d["df_raiz"]
                  + "\n\n" + d["du_proyecto"] + "\t" + RUTA_PROYECTO_VM))
    b.append(p(
        "La maquina tiene " + mem[1] + " de memoria total, con " + mem[2] + " en uso y "
        + mem[3] + " libres. La columna buff/cache (" + mem[5] + ") no es memoria perdida "
        "sino cache de disco que el nucleo devuelve en cuanto un proceso la necesita, y "
        "por eso la disponible (" + mem[6] + ") supera a la libre; el intercambio queda en "
        "0 B usados, senal de que el trabajo nunca presiono la memoria fisica. El proyecto "
        "vive en " + dfr[0] + " montado en /, con " + dfr[1] + " de capacidad, " + dfr[2]
        + " usados y " + dfr[3] + " disponibles (" + dfr[4] + " de uso), y ocupa "
        + d["du_proyecto"] + "; ese total supera la suma de sus subcarpetas por el "
        "redondeo de du a bloques de " + d["stat_ioblock"] + " bytes."))

    b.append(h2("stat", "Metadatos de la bitacora con stat"))
    b.append(code("$ stat logs/gestion_ambiental.log ; ls -lah logs/gestion_ambiental.log",
                  d["stat_bitacora"] + "\n" + d["ls_bitacora"]))
    b.append(tabla(
        ["Campo de stat", "Valor", "Lectura"],
        [["Tipo y tamano",
          "regular file, " + d["stat_size"] + " B en " + d["stat_bloques"] + " bloques",
          "Archivo ordinario de datos. El tamano logico es menor que el espacio "
          "reservado, porque el sistema de archivos asigna bloques completos de "
          + d["stat_ioblock"] + " bytes."],
         ["Inodo y enlaces",
          d["stat_inodo"] + " (dispositivo " + d["stat_dispositivo"] + "), "
          + d["stat_enlaces"] + " enlace",
          "El nombre gestion_ambiental.log es solo una entrada de directorio que apunta a "
          "ese inodo. Un enlace duro subiria el contador a 2 sin duplicar un byte."],
         ["Permisos", d["stat_permisos"],
          "Lectura y escritura para el propietario y su grupo, solo lectura para el "
          "resto: la bitacora es auditable por terceros pero no modificable."],
         ["Propietario / grupo", d["stat_uid"] + " / " + d["stat_gid"],
          "Pertenece al usuario sin privilegios del equipo, no a root: el gestor no "
          "necesita permisos elevados para operar."],
         ["Access / Modify / Change / Birth", "iguales al segundo",
          "Ultima lectura, ultima escritura de contenido, ultimo cambio de metadatos del "
          "inodo y creacion: coinciden porque el archivo se creo y se escribio en la "
          "misma corrida."]],
        anchos=[3.0, 3.9, 9.6]))
    b.append(p(
        "Esa bitacora tiene " + d["ev_n_bitacora"] + " lineas y " + d["stat_size"]
        + " bytes. Ambas cifras corresponden a la misma corrida: el generador de este "
        "informe compara el conteo de lineas del archivo entregado con el que declara la "
        "evidencia, y el tamano del archivo entregado con el que devolvio stat, y se "
        "detiene si alguno de los dos pares no calza."))

    # ---------------------------------------- 7. Mover vs copiar y errores
    b.append(h1("mover", "Mover frente a copiar, y control de errores"))

    b.append(h2("mover_copiar", "Por que los informes se mueven y el resumen se copia"))
    b.append(p(
        "Los informes se MUEVEN desde salida/ a su carpeta de clasificacion porque el "
        "traslado es un cambio de estado definitivo: un informe ya clasificado no debe "
        "seguir figurando como pendiente. Copiarlos duplicaria cada informe y se perderia "
        "la propiedad mas util: que salida/ sin informes significa \"no queda nada por "
        "clasificar\"."))
    b.append(p(
        "Mover dentro de una misma particion (aqui origen y destino estan ambos en "
        + dfr[0] + ") es la llamada rename(2): no copia datos, elimina la entrada de "
        "directorio antigua y crea otra que apunta al MISMO inodo. Los bloques no se "
        "tocan, el numero de inodo no cambia y solo se actualiza el ctime; por eso es casi "
        "instantanea y su costo no depende del tamano. Entre particiones distintas el "
        "nucleo no podria renombrar y shutil.move degradaria a copiar y borrar."))
    b.append(p(
        "El resumen consolidado tiene otro papel: es la vista general que el modulo de "
        "monitoreo consulta en su ruta de siempre. Por eso salida/resumen_ambiental.txt se "
        "CONSERVA y ademas se COPIA como gestion_ambiental/resumen_resguardado.txt con "
        "shutil.copy. Copiar si crea un inodo nuevo, con su contenido en otros bloques: "
        "desde ese momento los dos archivos son independientes y modificar uno no altera "
        "al otro, que es lo que se espera de un respaldo."))
    b.append(tabla(
        ["", "Mover (rename)", "Copiar (copy)"],
        [["Se aplica a", "informe_*.txt", "resumen_ambiental.txt"],
         ["Efecto en el inodo", "conserva el mismo inodo", "crea un inodo nuevo"],
         ["Datos en disco", "no se duplican", "se duplican"],
         ["Original", "deja de existir en el origen", "permanece en salida/"],
         ["Proposito", "reclasificacion definitiva", "respaldo y auditoria"]],
        anchos=[3.6, 6.5, 6.4]))
    b.append(code("# Comprobacion posterior a la ejecucion del gestor", d["mover_copiar"]))
    b.append(p(
        "El resumen original (" + str(d["resumen_bytes"]) + " bytes) sigue en salida/, la "
        "copia resguardada pesa " + str(d["resguardo_bytes"]) + " bytes y su contenido es "
        "identico byte a byte: "
        + ("comprobado" if d["resguardo_identico"] else "NO coincide") + ". En salida/ "
        "quedan " + str(d["n_salida"]) + " informes individuales, porque todos fueron "
        "movidos y no copiados; esto es lo que documenta " + ref("leeme") + "."))

    b.append(h2("errores", "Control de errores sin detener el procesamiento"))
    b.append(nota(
        "Esta demostracion es DESTRUCTIVA: inyecta datos danados. Por eso NO se ejecuto "
        "sobre el arbol entregado sino sobre una COPIA completa en " + d["copia_control"]
        + ". De ahi que la bitacora entregada cierre con \"no se detectaron anomalias en "
        "esta corrida\": el entregable viene de una corrida limpia y las cifras de esta "
        "subseccion pertenecen a la copia."))
    b.append(p(
        "El requisito 13 pide controlar al menos un error sin detener el procesamiento. "
        "Se inyectaron cuatro fallas distintas, para probar tanto el analisis de la "
        "bitacora de alertas como la lectura de los informes: una linea sin separadores, "
        "otra con campos insuficientes, un indicador desconocido (\"Radiacion\") y un "
        "informe sin la linea \"Alertas detectadas: N\"."))
    b.append(code("$ python3 src/gestor_incidencias.py ; echo \"Codigo de salida: $?\"\n"
                  "$ grep 'ANOMALIA\\|CONTROL DE ERRORES' logs/gestion_ambiental.log",
                  "Codigo de salida: " + d["gestor_exit_error"] + "\n\n"
                  + "\n".join(d["anomalias_lineas"][:4]) + "\n" + d["control_errores"]))
    b.append(p(
        "El comportamiento es el correcto: las " + d["anomalias_total"] + " anomalias "
        "quedaron registradas con su numero de linea y su contenido, el gestor NO se "
        "detuvo, termino con codigo " + d["gestor_exit_error"] + ", conservo las "
        + d["inventario_tras_error"] + " alertas validas y genero el inventario igual. "
        "Descartar el dato defectuoso y continuar es preferible a abortar: una linea "
        "corrupta no puede invalidar una jornada de monitoreo, pero tampoco puede "
        "desaparecer en silencio."))
    b.append(fig(recortar_consola(FOTOS / "debian-04-stat-y-control-de-errores.png"),
                 "Metadatos de la bitacora con stat y evidencia del control de "
                 "errores en la consola de la maquina virtual. La captura corresponde a "
                 "una corrida de control anterior, con dos alertas danadas en vez de las "
                 "cuatro fallas transcritas arriba; por eso declara dos anomalias y otro "
                 "tamano de bitacora. El resultado es el mismo: registra cada anomalia, no "
                 "se detiene, termina con codigo " + d["gestor_exit_error"] + " y completa "
                 "el inventario con las " + d["inventario_tras_error"] + " alertas. La "
                 "ultima linea queda cortada por el borde de la consola; su texto completo "
                 "esta en evidencias/debian/06-control-de-errores.txt."))

    b.append(h2("anti", "Idempotencia y proteccion contra sobrescritura"))
    b.append(nota(
        "Esta demostracion tambien es destructiva, porque duplica informes a proposito, y "
        "por eso se ejecuto sobre la misma COPIA " + d["copia_anti"] + ". El arbol "
        "entregado conserva sus " + str(inv["resumen"]["total_informes"]) + " informes "
        "sin sufijos de version, que es el estado de una unica corrida limpia."))
    b.append(code("# Tres corridas consecutivas del gestor sobre la copia\n"
                  "$ ls -lah gestion_ambiental/criticas/informe_STG04_20240101*",
                  d["idempotencia"] + "\n\n" + d["ejemplo_versiones"]))
    b.append(p(
        "Las alertas no se duplican entre corridas porque los registros por indicador se "
        "reconstruyen de forma atomica sobre archivos temporales que se renombran al "
        "final. Al reprocesar, los informes ya clasificados no se pierden ni se pisan: "
        "obtener_nombre_seguro() comprueba la existencia del destino y agrega un sufijo "
        "correlativo. Partiendo de " + d["informes_iniciales_v"] + " informes, tras dos "
        "ciclos adicionales conviven " + d["encontrados_v"] + " donde se esperaban "
        + d["esperados_v"] + " (" + d["con_sufijo_v"] + " con sufijo): ningun archivo fue "
        "sobrescrito ni eliminado."))

    # ------------------------------------------------------ 8. Conclusiones
    b.append(h1("conclusiones", "Conclusiones"))

    b.append(h2("concurrencia",
                "Sobre concurrencia: el resultado que no favorece la hipotesis"))
    b.append(p(
        "La conclusion mas importante es tambien la mas incomoda: en este caso concreto "
        "la concurrencia NO acelero el procesamiento. La version concurrente tardo "
        + d["t_con"]["real"] + " frente a " + d["t_sec"]["real"] + " de la secuencial, y "
        "su tiempo interno fue " + d["t_con"]["interno"] + " s frente a "
        + d["t_sec"]["interno"] + " s. El dato no se maquilla, se explica:"))
    b.append(ul([
        "El trabajo util es demasiado pequeno: " + d["archivos_generados"]
        + " archivos con " + res["Lineas leidas"] + " lineas, procesados en centesimas de "
        "segundo.",
        "El costo fijo de la concurrencia no depende del tamano del trabajo: crear "
        + d["t_con"]["trabajadores"] + " hilos, encolar los archivos y tomar y soltar los "
        "mutex en cada actualizacion pesa mas que lo que se ahorra al repartir un trabajo "
        "tan breve.",
        "El trabajo esta dominado por entrada/salida sobre archivos diminutos que el "
        "nucleo ya tiene en cache, de modo que hay poca espera que solapar; y en CPython "
        "el bloqueo global del interprete impide que dos hilos ejecuten codigo Python puro "
        "a la vez.",
    ]))
    b.append(p(
        "La concurrencia si se observa cuando el trabajo crece: con la carga amplificada "
        "de " + d["carga_archivos"] + " archivos el proceso alcanzo " + d["pcpu_max"]
        + " % de CPU, por encima del 100 % que jamas superaria un proceso de un solo hilo. "
        "El mecanismo funciona; lo que falta en el conjunto oficial es trabajo suficiente "
        "para amortizar su costo de entrada."))

    b.append(h2("sincronizacion", "Sobre sincronizacion y sistema de archivos"))
    b.append(p(
        "Que las " + d["metricas_comparadas"] + " metricas coincidan y que los "
        + d["informes_comparados"] + " informes sean identicos es la evidencia de que la "
        "sincronizacion es correcta: con " + d["t_con"]["trabajadores"] + " hilos "
        "actualizando contadores compartidos, sin mutex cada corrida habria dado un "
        "resultado distinto. El patron fue acumular en variables locales y entrar a la "
        "seccion critica una sola vez por archivo; queue.Queue garantiza que cada archivo "
        "lo toma un unico trabajador. El unico efecto observable de la concurrencia es el "
        "ORDEN de las lineas de la bitacora."))
    b.append(p(
        "Mover y copiar obligo a distinguir el nombre de un archivo del archivo mismo: un "
        "nombre es una entrada de directorio que apunta a un inodo. Las herramientas "
        "completaron el cuadro: stat dio inodo, enlaces, permisos y marcas de tiempo; "
        "ls -lah y du -sh, la diferencia entre tamano logico y bloques ocupados; df -h, el "
        "sistema de archivos y su espacio libre; free -h, que la cache no es memoria "
        "perdida; y ps con /proc, la estructura interna del proceso."))

    b.append(h2("honestidad", "Declaraciones de honestidad tecnica"))
    b.append(ul([
        "Hipervisor: el hito declaro " + HIPERVISOR_PLANIFICADO + " y la implementacion "
        "final se hizo sobre " + hv["hipervisor_implementado"] + ", con los recursos "
        "comprometidos intactos; la pauta admite otro hipervisor. TODAS las figuras de "
        "instalacion son de la maquina real en Hyper-V; la etapa previa sobre "
        + HIPERVISOR_PLANIFICADO + " queda archivada, no presentada como evidencia de la "
        "maquina entregada.",
        "Rendimiento: la version concurrente resulto mas lenta que la secuencial en el "
        "conjunto oficial; se informa tal como se midio.",
        "Clasificacion: sin_alertas/ quedo con " + str(d["n_sin"]) + " informes porque la "
        "pauta exige al menos una alerta por archivo; la rama se demostro aparte.",
        "Demostraciones destructivas: el control de errores (" + ref("errores") + ") y la "
        "anti-sobrescritura (" + ref("anti") + ") se ejecutaron sobre una COPIA, no sobre "
        "el arbol entregado; por eso la bitacora entregada no registra esas anomalias.",
        "Trazabilidad: cada cifra se lee al generar el informe desde evidencias/ y del "
        "estado real del proyecto, y el generador compara ambas fuentes -y las cifras del "
        "README- antes de escribir nada; si discrepan, no se emite el documento.",
    ]))

    b.append(h2("indice", "Indice de evidencias que respaldan este informe"))
    b.append(p(
        "Cada seccion puede auditarse contra el archivo que la origina; el generador "
        "falla si alguno cambia de forma incompatible."))
    b.append(tabla(
        ["Archivo de evidencia", "Contenido", "Secciones"],
        [["evidencias/hyperv/configuracion-vm-hyperv.txt",
          "cmdlets de Hyper-V y verificacion SHA256 del ISO",
          ref("iso") + ", " + ref("hipervisor")],
         ["evidencias/debian/01-preparar-ambiente.txt",
          "apt update, apt upgrade y entorno de Python", ref("ambiente")],
         ["evidencias/debian/02-entorno.txt",
          "distribucion, kernel, hipervisor, CPU, memoria, disco y red",
          ref("instalacion")],
         ["evidencias/debian/03-ejecucion-parte1.txt",
          "ambas corridas, las seis metricas y la comparacion de informes y alertas",
          ref("parte1")],
         ["evidencias/debian/04-observacion-procesos.txt",
          "/usr/bin/time -v, ps, ps -L, /proc, free -h, df -h y du -sh", ref("procesos")],
         ["evidencias/debian/05-gestor-incidencias.txt",
          "estructura, inventario, stat de la bitacora, mover frente a copiar",
          ref("gestor") + ", " + ref("mover_copiar")],
         ["evidencias/debian/06-control-de-errores.txt",
          "anomalias inyectadas y rama sin_alertas (sobre una copia)",
          ref("sin_alertas") + ", " + ref("errores")],
         ["evidencias/debian/07-anti-sobrescritura.txt",
          "sufijos _v1 y _v2 e idempotencia (sobre una copia)", ref("anti")],
         ["evidencias/salida_parte1/",
          d["respaldo_parte1"] + " archivos: salida/ tal como la dejo la Parte 1",
          ref("leeme")],
         ["evidencias/fotos/",
          "instalacion real en Hyper-V y consola de la maquina definitiva",
          "Figuras 1 a %d" % n_figuras()],
         ["gestion_ambiental/, alertas/, logs/, salida/",
          "estado real: informes, alertas por indicador, inventario y bitacora",
          ref("gestor") + ", " + ref("mover")],
         ["docs/Solemne01PracticaParte2FormaB.pdf",
          "pauta oficial contra la cual se verifico cada requisito", "todas"]],
        anchos=[6.4, 7.2, 2.9]))

    return b

# ---------------------------------------------------------------------------
# Render DOCX
# ---------------------------------------------------------------------------

def render_docx(bloques, destino):
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    def campo(par, instruccion):
        """Inserta un campo de Word (PAGE, NUMPAGES) en un parrafo."""
        run = par.add_run()
        ini = OxmlElement("w:fldChar")
        ini.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instruccion
        fin = OxmlElement("w:fldChar")
        fin.set(qn("w:fldCharType"), "end")
        run._r.append(ini)
        run._r.append(instr)
        run._r.append(fin)
        return run

    doc = Document()

    # El DOCX debe salir en A4 igual que el PDF. python-docx crea el documento
    # en tamano Carta, que no es el formato de la entrega, y ademas deja el pie
    # de pagina vacio: el documento perdia la numeracion de paginas.
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.6)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        section.footer_distance = Cm(1.0)
        section.different_first_page_header_footer = True

        pie = section.footer.paragraphs[0]
        pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
        izq = pie.add_run(limpiar(EVALUACION + "  |  " + EQUIPO + "  |  "
                                  + SECCION + "  |  Pagina "))
        izq.font.size = Pt(8)
        izq.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
        num = campo(pie, " PAGE ")
        num.font.size = Pt(8)
        num.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
        de = pie.add_run(" de ")
        de.font.size = Pt(8)
        de.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
        tot = campo(pie, " NUMPAGES ")
        tot.font.size = Pt(8)
        tot.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    # Metadatos: sin esto el archivo se entrega firmado por "python-docx".
    props = doc.core_properties
    props.title = limpiar("Informe final - " + EVALUACION)
    props.subject = limpiar(ASIGNATURA + " - " + EVALUACION)
    props.author = limpiar(EQUIPO + " (" + ", ".join(INTEGRANTES) + ")")
    props.last_modified_by = limpiar(EQUIPO)
    props.category = limpiar(UNIVERSIDAD + " - " + ASIGNATURA)
    props.comments = limpiar(
        "Generado de forma reproducible por scripts/generar_informe.py a partir "
        "de los archivos de evidencia del repositorio.")
    props.keywords = limpiar("Debian, concurrencia, gestor de incidencias, "
                             + EQUIPO + ", " + SECCION)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.06

    def sombrear(par, color="F2F3F5"):
        pr = par._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        pr.append(shd)

    def borde_izquierdo(par, color="6E7B8B"):
        pr = par._p.get_or_add_pPr()
        bordes = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "18")
        left.set(qn("w:space"), "6")
        left.set(qn("w:color"), color)
        bordes.append(left)
        pr.append(bordes)

    for blk in bloques:
        t = blk["t"]

        if t == "portada":
            for _ in range(3):
                doc.add_paragraph()
            def centrado(texto, size, negrita=False, color=None, espacio=6):
                par = doc.add_paragraph()
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                par.paragraph_format.space_after = Pt(espacio)
                run = par.add_run(limpiar(texto))
                run.font.size = Pt(size)
                run.bold = negrita
                if color:
                    run.font.color.rgb = RGBColor(*color)
                return par
            centrado(UNIVERSIDAD, 16, True)
            centrado(ASIGNATURA, 13, False, espacio=24)
            centrado(EVALUACION, 20, True, (0x1F, 0x3A, 0x5F))
            centrado(SUBTITULO, 13, False, espacio=30)
            centrado(EQUIPO + "  -  " + SECCION, 13, True)
            centrado(LENGUAJE, 11, False, espacio=20)
            centrado("Integrantes", 11.5, True, espacio=4)
            for nombre in INTEGRANTES:
                centrado(nombre, 11, False, espacio=2)
            doc.add_paragraph()
            centrado(FECHA, 11)

        elif t == "salto":
            par = doc.add_paragraph()
            par.add_run().add_break(WD_BREAK.PAGE)

        elif t == "h1":
            par = doc.add_paragraph()
            par.paragraph_format.space_before = Pt(4)
            par.paragraph_format.space_after = Pt(8)
            run = par.add_run(limpiar(blk["x"]))
            run.bold = True
            run.font.size = Pt(15)
            run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

        elif t == "h2":
            par = doc.add_paragraph()
            par.paragraph_format.space_before = Pt(8)
            par.paragraph_format.space_after = Pt(4)
            run = par.add_run(limpiar(blk["x"]))
            run.bold = True
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x2E, 0x5A, 0x88)

        elif t == "p":
            par = doc.add_paragraph(limpiar(blk["x"]))
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        elif t in ("ul", "ol"):
            estilo = "List Bullet" if t == "ul" else "List Number"
            for item in blk["x"]:
                par = doc.add_paragraph(limpiar(item), style=estilo)
                par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                par.paragraph_format.space_after = Pt(3)

        elif t == "code":
            if blk["cmd"]:
                par = doc.add_paragraph()
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.space_before = Pt(4)
                sombrear(par, "E8EAED")
                run = par.add_run(limpiar(blk["cmd"]))
                run.font.name = "Consolas"
                run.font.size = Pt(7.5)
                run.bold = True
            for linea in limpiar(blk["x"]).splitlines() or [""]:
                par = doc.add_paragraph()
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.line_spacing = 1.0
                sombrear(par, "F4F5F7")
                run = par.add_run(linea)
                run.font.name = "Consolas"
                run.font.size = Pt(7.5)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

        elif t == "nota":
            par = doc.add_paragraph()
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            par.paragraph_format.left_indent = Cm(0.3)
            par.paragraph_format.space_before = Pt(6)
            par.paragraph_format.space_after = Pt(8)
            sombrear(par, "FBF3E2")
            borde_izquierdo(par, "B7791F")
            run = par.add_run(limpiar(blk["x"]))
            run.font.size = Pt(10)

        elif t == "tabla":
            tab = doc.add_table(rows=1, cols=len(blk["head"]))
            tab.style = "Table Grid"
            hdr = tab.rows[0].cells
            for i, texto in enumerate(blk["head"]):
                hdr[i].text = ""
                par = hdr[i].paragraphs[0]
                par.paragraph_format.space_after = Pt(2)
                run = par.add_run(limpiar(texto))
                run.bold = True
                run.font.size = Pt(8.5)
            for fila in blk["rows"]:
                celdas = tab.add_row().cells
                for i, texto in enumerate(fila):
                    celdas[i].text = ""
                    par = celdas[i].paragraphs[0]
                    par.paragraph_format.space_after = Pt(2)
                    run = par.add_run(limpiar(str(texto)))
                    run.font.size = Pt(8)
            if blk.get("w"):
                total = sum(blk["w"])
                disponible = 17.0
                for i, ancho in enumerate(blk["w"]):
                    cm = Cm(ancho / total * disponible)
                    for fila in tab.rows:
                        fila.cells[i].width = cm
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

        elif t == "fig":
            from PIL import Image
            ruta = blk["ruta"]
            with Image.open(ruta) as im:
                ancho_px, alto_px = im.size
            ancho = ancho_de_figura(ruta, 17.0, blk.get("texto_pt"))
            alto = ancho * alto_px / ancho_px
            tope = blk.get("alto")
            if tope is not None and alto > tope:
                alto = tope
                ancho = alto * ancho_px / alto_px
            par = doc.add_paragraph()
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.space_after = Pt(2)
            par.add_run().add_picture(ruta, width=Cm(ancho))
            pie = doc.add_paragraph()
            pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pie.paragraph_format.space_after = Pt(8)
            run = pie.add_run(limpiar(blk["pie"]))
            run.italic = True
            run.font.size = Pt(8.5)

        else:
            morir("bloque desconocido en el modelo: %r" % t)

    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))


# ---------------------------------------------------------------------------
# Render PDF
# ---------------------------------------------------------------------------

def render_pdf(bloques, destino):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (BaseDocTemplate, Frame, Image, KeepTogether,
                                    PageBreak, PageTemplate, Paragraph, Spacer,
                                    Table, TableStyle)
    from xml.sax.saxutils import escape

    AZUL = colors.HexColor("#1F3A5F")
    AZUL2 = colors.HexColor("#2E5A88")
    GRIS = colors.HexColor("#F4F5F7")
    GRIS2 = colors.HexColor("#E8EAED")
    CREMA = colors.HexColor("#FBF3E2")
    AMBAR = colors.HexColor("#B7791F")

    base = getSampleStyleSheet()
    S = {}
    S["p"] = ParagraphStyle("p", parent=base["BodyText"], fontName="Helvetica",
                            fontSize=9.0, leading=10.9, alignment=TA_JUSTIFY,
                            spaceAfter=3.5)
    S["h1"] = ParagraphStyle("h1", parent=S["p"], fontName="Helvetica-Bold",
                             fontSize=13, leading=15.5, textColor=AZUL,
                             spaceBefore=3, spaceAfter=5, alignment=0)
    S["h2"] = ParagraphStyle("h2", parent=S["p"], fontName="Helvetica-Bold",
                             fontSize=10.2, leading=12.4, textColor=AZUL2,
                             spaceBefore=5, spaceAfter=2, alignment=0)
    S["li"] = ParagraphStyle("li", parent=S["p"], leftIndent=0.55 * cm,
                             bulletIndent=0.15 * cm, spaceAfter=2.5)
    S["code"] = ParagraphStyle("code", parent=S["p"], fontName="Courier",
                               fontSize=6.6, leading=7.3, alignment=0,
                               spaceAfter=0, spaceBefore=0)
    S["cmd"] = ParagraphStyle("cmd", parent=S["code"], fontName="Courier-Bold")
    S["nota"] = ParagraphStyle("nota", parent=S["p"], fontSize=8.8, leading=10.9,
                               leftIndent=0.25 * cm, rightIndent=0.15 * cm)
    S["pie"] = ParagraphStyle("pie", parent=S["p"], fontName="Helvetica-Oblique",
                              fontSize=7.4, leading=8.7, alignment=TA_CENTER,
                              spaceBefore=1.5, spaceAfter=4)
    S["td"] = ParagraphStyle("td", parent=S["p"], fontSize=7.5, leading=8.6,
                             spaceAfter=0, alignment=0)
    S["th"] = ParagraphStyle("th", parent=S["td"], fontName="Helvetica-Bold",
                             textColor=colors.white)
    S["portada_u"] = ParagraphStyle("pu", parent=S["p"], alignment=TA_CENTER,
                                    fontName="Helvetica-Bold", fontSize=16,
                                    leading=20, spaceAfter=4)
    S["portada_a"] = ParagraphStyle("pa", parent=S["p"], alignment=TA_CENTER,
                                    fontSize=12.5, leading=16, spaceAfter=26)
    S["portada_t"] = ParagraphStyle("pt", parent=S["p"], alignment=TA_CENTER,
                                    fontName="Helvetica-Bold", fontSize=19,
                                    leading=24, textColor=AZUL, spaceAfter=4)
    S["portada_s"] = ParagraphStyle("ps", parent=S["p"], alignment=TA_CENTER,
                                    fontSize=12.5, leading=16, spaceAfter=32)
    S["portada_e"] = ParagraphStyle("pe", parent=S["p"], alignment=TA_CENTER,
                                    fontName="Helvetica-Bold", fontSize=12.5,
                                    leading=16, spaceAfter=3)
    S["portada_n"] = ParagraphStyle("pn", parent=S["p"], alignment=TA_CENTER,
                                    fontSize=10.5, leading=14, spaceAfter=2)

    ancho_util = A4[0] - 2 * 2.0 * cm

    def esc(texto):
        return escape(limpiar(texto))

    historia = []

    def pie_pagina(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.6)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        if doc_.page > 1:
            canvas.drawString(2.0 * cm, 0.95 * cm,
                              limpiar(EVALUACION + "  |  " + EQUIPO + "  |  " + SECCION))
            canvas.drawRightString(A4[0] - 2.0 * cm, 0.95 * cm,
                                   "Pagina %d" % doc_.page)
            canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
            canvas.line(2.0 * cm, 1.3 * cm, A4[0] - 2.0 * cm, 1.3 * cm)
        canvas.restoreState()

    for blk in bloques:
        t = blk["t"]

        if t == "portada":
            historia.append(Spacer(1, 3.6 * cm))
            historia.append(Paragraph(esc(UNIVERSIDAD), S["portada_u"]))
            historia.append(Paragraph(esc(ASIGNATURA), S["portada_a"]))
            historia.append(Paragraph(esc(EVALUACION), S["portada_t"]))
            historia.append(Paragraph(esc(SUBTITULO), S["portada_s"]))
            historia.append(Paragraph(esc(EQUIPO + "  -  " + SECCION), S["portada_e"]))
            historia.append(Paragraph(esc(LENGUAJE), S["portada_n"]))
            historia.append(Spacer(1, 0.8 * cm))
            historia.append(Paragraph("<b>" + esc("Integrantes") + "</b>", S["portada_n"]))
            for nombre in INTEGRANTES:
                historia.append(Paragraph(esc(nombre), S["portada_n"]))
            historia.append(Spacer(1, 1.0 * cm))
            historia.append(Paragraph(esc(FECHA), S["portada_n"]))

        elif t == "salto":
            historia.append(PageBreak())

        elif t == "h1":
            historia.append(Paragraph(esc(blk["x"]), S["h1"]))

        elif t == "h2":
            historia.append(Paragraph(esc(blk["x"]), S["h2"]))

        elif t == "p":
            historia.append(Paragraph(esc(blk["x"]), S["p"]))

        elif t == "ul":
            for item in blk["x"]:
                historia.append(Paragraph(esc(item), S["li"], bulletText="•"))

        elif t == "ol":
            for i, item in enumerate(blk["x"], 1):
                historia.append(Paragraph(esc(item), S["li"], bulletText="%d." % i))

        elif t == "code":
            filas = []
            if blk["cmd"]:
                for linea in limpiar(blk["cmd"]).splitlines():
                    filas.append([Paragraph(escape(linea) or "&nbsp;", S["cmd"])])
            n_cmd = len(filas)
            for linea in limpiar(blk["x"]).splitlines() or [""]:
                # Las lineas cortas conservan la alineacion de columnas usando
                # espacios duros; las muy largas se dejan partir para no salirse
                # del ancho util de la pagina.
                texto = escape(linea)
                if len(linea) <= 116:
                    texto = texto.replace(" ", "&nbsp;")
                filas.append([Paragraph(texto or "&nbsp;", S["code"])])
            tab = Table(filas, colWidths=[ancho_util])
            estilo = [
                ("BACKGROUND", (0, 0), (-1, -1), GRIS),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 0.4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0.4),
                ("LINEBEFORE", (0, 0), (0, -1), 1.6, colors.HexColor("#9AA5B1")),
            ]
            if n_cmd:
                estilo.append(("BACKGROUND", (0, 0), (-1, n_cmd - 1), GRIS2))
                estilo.append(("TOPPADDING", (0, 0), (-1, 0), 3))
                estilo.append(("BOTTOMPADDING", (0, n_cmd - 1), (-1, n_cmd - 1), 3))
            estilo.append(("BOTTOMPADDING", (0, -1), (-1, -1), 3))
            tab.setStyle(TableStyle(estilo))
            historia.append(Spacer(1, 2))
            historia.append(tab)
            historia.append(Spacer(1, 5))

        elif t == "nota":
            tab = Table([[Paragraph(esc(blk["x"]), S["nota"])]], colWidths=[ancho_util])
            tab.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), CREMA),
                ("LINEBEFORE", (0, 0), (0, -1), 2.4, AMBAR),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            historia.append(Spacer(1, 2))
            historia.append(tab)
            historia.append(Spacer(1, 5))

        elif t == "tabla":
            pesos = blk.get("w") or [1] * len(blk["head"])
            total = float(sum(pesos))
            anchos = [ancho_util * w / total for w in pesos]
            filas = [[Paragraph(esc(c), S["th"]) for c in blk["head"]]]
            for fila in blk["rows"]:
                filas.append([Paragraph(esc(str(c)), S["td"]) for c in fila])
            tab = Table(filas, colWidths=anchos, repeatRows=1)
            tab.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), AZUL2),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                 [colors.white, colors.HexColor("#F5F7FA")]),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C8CDD4")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))
            historia.append(Spacer(1, 3))
            historia.append(tab)
            historia.append(Spacer(1, 7))

        elif t == "fig":
            from PIL import Image as PILImage
            with PILImage.open(blk["ruta"]) as im:
                ancho_px, alto_px = im.size
            ancho = ancho_de_figura(blk["ruta"], ancho_util, blk.get("texto_pt"))
            alto = ancho * alto_px / float(ancho_px)
            tope = blk.get("alto")
            if tope is not None and alto > tope * cm:
                alto = tope * cm
                ancho = alto * ancho_px / float(alto_px)
            escala = ancho / float(ancho_px)
            linea_px = altura_de_linea_px(blk["ruta"])
            ESCALAS.append({
                "n": blk.get("n"),
                "ruta": Path(blk["ruta"]).name,
                "nativo": (ancho_px, alto_px),
                "colocada": (ancho, alto),
                "escala": escala,
                "texto_pt": None if linea_px is None else linea_px * escala,
            })
            img = Image(blk["ruta"], width=ancho, height=alto)
            img.hAlign = "CENTER"
            historia.append(Spacer(1, 2))
            historia.append(KeepTogether([img, Paragraph(esc(blk["pie"]), S["pie"])]))

        else:
            morir("bloque desconocido en el modelo: %r" % t)

    # Un Spacer al final del relato hace que reportlab abra una pagina mas y
    # la deje en blanco. Se descartan los espaciadores colgantes.
    while historia and isinstance(historia[-1], Spacer):
        historia.pop()

    destino.parent.mkdir(parents=True, exist_ok=True)
    doc = BaseDocTemplate(str(destino), pagesize=A4,
                          leftMargin=2.0 * cm, rightMargin=2.0 * cm,
                          topMargin=1.5 * cm, bottomMargin=1.6 * cm,
                          title=limpiar("Informe final - " + EVALUACION),
                          author=limpiar(EQUIPO + " (" + ", ".join(INTEGRANTES) + ")"),
                          subject=limpiar(ASIGNATURA + " - " + EVALUACION),
                          creator=limpiar(UNIVERSIDAD + " - " + EQUIPO))
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height,
                  id="cuerpo")
    doc.addPageTemplates([PageTemplate(id="normal", frames=[frame],
                                       onPage=pie_pagina)])
    doc.build(historia)


# ---------------------------------------------------------------------------

def verificar_titulos(bloques):
    """Comprueba que los numeros de seccion emitidos existan de verdad.

    resolver_referencias() ya garantiza que ninguna referencia apunte a una
    clave inexistente. Esto cierra el circulo por el otro lado: recorre los
    titulos realmente escritos en el documento y exige que el registro de
    secciones y los titulos coincidan exactamente.
    """
    numerados = []
    for blk in bloques:
        if blk.get("t") in ("h1", "h2"):
            m = re.match(r"^(\d+(?:\.\d+)?)[. ]", blk["x"])
            if not m:
                morir("el titulo %r no lleva numero de seccion" % blk["x"])
            numerados.append(m.group(1))
    faltan = sorted(set(SECCIONES.values()) - set(numerados))
    sobran = sorted(set(numerados) - set(SECCIONES.values()))
    if faltan or sobran:
        morir("los numeros de seccion registrados y los titulos escritos no "
              "coinciden (registrados sin titulo: %s; titulos sin registro: %s)"
              % (faltan, sobran))
    if len(numerados) != len(set(numerados)):
        morir("hay numeros de seccion repetidos en el documento")


def contar_paginas(pdf):
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    return len(PdfReader(str(pdf)).pages)


def main():
    print("[generar_informe] Leyendo evidencia del repositorio...")
    datos = recolectar_datos()
    print("[generar_informe] Preparando capturas...")
    # Las capturas de la etapa previa ya no ilustran ninguna figura, pero se
    # siguen extrayendo y cacheando porque el informe las cita por su ruta.
    extraer_capturas_instalacion()
    imgs_hv = capturas_instalacion_real()
    print("[generar_informe] Construyendo el documento...")
    bloques = construir_documento(datos, imgs_hv)
    bloques = resolver_referencias(bloques)
    verificar_titulos(bloques)
    print("[generar_informe] Escribiendo DOCX...")
    render_docx(bloques, SALIDA_DOCX)
    print("[generar_informe] Escribiendo PDF...")
    render_pdf(bloques, SALIDA_PDF)

    paginas = contar_paginas(SALIDA_PDF)
    print("[generar_informe] Listo.")
    print("  " + str(SALIDA_DOCX))
    print("  " + str(SALIDA_PDF))
    if paginas is not None:
        estado = "dentro del rango 8-12" if 8 <= paginas <= 12 else "FUERA DEL RANGO 8-12"
        print("  Paginas del PDF: %d (%s)" % (paginas, estado))

    # Escala de cada figura. Una captura de terminal colocada muy por debajo de
    # su tamano nativo deja de ser evidencia, porque no se puede leer; este
    # listado permite comprobarlo sin abrir el PDF.
    print("\n[generar_informe] Colocacion de las figuras en el PDF "
          "(texto minimo legible: %.1f pt):" % TEXTO_MINIMO_PT)
    bajas = []
    for f in ESCALAS:
        if f["texto_pt"] is None:
            marca = "   (sin texto de terminal)"
        elif f["texto_pt"] >= TEXTO_MINIMO_PT:
            marca = "   texto %.1f pt" % f["texto_pt"]
        else:
            marca = "   texto %.1f pt  <-- ILEGIBLE" % f["texto_pt"]
            bajas.append(f["n"])
        print("  Fig %-2d  %-42s %5dx%-5d px -> %5.1f x %-5.1f pt   escala %.3f%s"
              % (f["n"], f["ruta"], f["nativo"][0], f["nativo"][1],
                 f["colocada"][0], f["colocada"][1], f["escala"], marca))
    if bajas:
        print("  ATENCION: figuras con texto por debajo de %.1f pt: %s"
              % (TEXTO_MINIMO_PT, ", ".join(str(n) for n in bajas)))
    else:
        print("  Ninguna figura quedo con su texto por debajo de %.1f pt."
              % TEXTO_MINIMO_PT)


if __name__ == "__main__":
    main()
