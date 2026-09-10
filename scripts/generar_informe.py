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

Las unicas figuras del informe son archivos PNG del repositorio: la instalacion
real sobre Hyper-V (evidencias/fotos/instalacion-hyperv/) y la consola de la
maquina definitiva (evidencias/fotos/debian-0*.png).

Dependencias: python-docx, reportlab, Pillow.

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
FOTOS_HYPERV = FOTOS / "instalacion-hyperv"
CACHE_RECORTES = FOTOS / "recortes"
DOCS = RAIZ / "docs"

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
UNIVERSIDAD = "Universidad San Sebastián"
ASIGNATURA = "Sistemas Operativos"
EVALUACION = "Solemne 01 Práctico - Parte 2 - Forma B"
SUBTITULO = "Informe final asincrónico"
EQUIPO = "Equipo 07"
SECCION = "Sección NRC 18897"
LENGUAJE = "Lenguaje utilizado en la Parte 1: Python"
INTEGRANTES = [
    "Benjamín Zamora",
    "José Palma",
    "Franco Maripil",
    "Nicolás Portilla",
    "Thomas Márquez",
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

# %CPU maximo que muestra EN PANTALLA la captura de la observacion en vivo
# (evidencias/fotos/debian-02-observacion-procesos.png). Es el unico dato del
# informe que se lee de una imagen y no de un archivo de texto, porque esa
# corrida no dejo transcripcion: la captura ES su evidencia. Se declara aqui,
# a la vista, para que quien revise la figura pueda contrastarlo de un vistazo
# y corregirlo en un solo lugar si la captura cambiara.
FIG_OBSERVACION_PCPU = "107"

# Marcador con el que el control de versiones conserva una carpeta vacia. No lo
# produjo el gestor, y por eso el arbol entregado trae un archivo mas que la
# evidencia congelada bajo gestion_ambiental/. verificar_coherencia() exige que
# la diferencia entre ambos sea EXACTAMENTE este archivo y ningun otro.
MARCADOR_CARPETA_VACIA = "gestion_ambiental/sin_alertas/.gitkeep"

# Significado de la letra de estado que ps imprime en la columna STAT. La
# explicacion del informe NO se escribe a mano: se arma con la letra que traiga
# la muestra transcrita, de modo que no pueda describir un estado distinto del
# que la propia transcripcion muestra.
ESTADOS_PS = {
    "R": "ejecutándose en CPU",
    "S": "durmiendo en espera interrumpible (entrada/salida o un mutex)",
    "D": "en espera ininterrumpible por entrada/salida",
}

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


# Tres lineas de los archivos de evidencia NO son salida de ningun comando:
# son frases que el equipo escribio en la bitacora de la sesion y que el
# informe intercala en sus propios parrafos. La consola de la maquina virtual
# se uso sin acentos, de modo que llegan sin ellos y quedarian como tres
# palabras sin tilde en medio de prosa acentuada. Se les restituye la tilde
# -solo la ortografia, ni una palabra ni una cifra cambia- y la lista queda
# aqui, a la vista, para que se pueda auditar exactamente que se toco. Las
# transcripciones de salida de comandos NO pasan por aqui: van literales.
TILDES_DE_EVIDENCIA = {
    "estandar": "estándar",
    "modulos": "módulos",
    "arbol": "árbol",
}


def acentuar_prosa_de_evidencia(texto):
    def uno(m):
        return TILDES_DE_EVIDENCIA[m.group(0).lower()]
    return re.sub(r"\b(?:%s)\b" % "|".join(TILDES_DE_EVIDENCIA), uno, texto)


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
# Rios de justificacion
#
# El cuerpo del informe va justificado a los dos margenes y esta lleno de rutas
# y nombres de archivo -scripts/generar_archivos_entrada.py,
# estacion_CODIGO_AAAAMMDD.jsonl, gestion_ambiental/resumen_resguardado.txt-.
# Una palabra de esas no cabe casi nunca en el hueco que queda al final de la
# linea y no se puede partir por ningun sitio, asi que el justificador no tiene
# mas remedio que repartir todo el sobrante entre los pocos espacios de la
# linea: salian lineas con huecos entre palabras de tres y cuatro veces lo
# normal, que el ojo lee como un rio blanco bajando por la columna.
#
# La solucion no es tocar el texto sino darle al justificador mas sitios por
# donde cortar. Una ruta se puede partir despues de una barra o de un guion
# bajo, y antes del punto de la extension, sin anadir ningun signo y sin que se
# lea peor -es como se parten las direcciones web-, y eso es lo unico que se
# hace aqui, en el momento de maquetar. El MODELO del documento no se toca: ni
# una palabra cambia.
#
# Las dos salidas lo aplican por caminos distintos porque cada motor parte las
# lineas a su manera:
#
# - El PDF no lleva ningun caracter nuevo. A reportlab se le ensena a partir
#   estas palabras (ver rutas_partibles_en_pdf), de modo que el texto extraido
#   del PDF sigue siendo exactamente el mismo.
# - El DOCX si lleva un espacio de ancho cero (U+200B) en cada punto de corte:
#   es el caracter que Word entiende como "aqui se puede partir", no ocupa
#   nada y no se imprime.
# ---------------------------------------------------------------------------

# Por debajo de esta longitud una palabra ya cabe en casi cualquier hueco y no
# es la que provoca el rio; partirla solo anadiria cortes inutiles.
CORTE_MINIMO = 10

# Una ruta o un nombre de archivo: solo letras sin tilde, cifras y los signos
# que aparecen en un nombre de archivo. Las palabras del idioma quedan fuera
# porque llevan tilde o porque no contienen ninguna barra ni guion bajo.
RUTA_LARGA = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_./\~@+-]*[A-Za-z0-9/]$")


def puntos_de_corte(palabra):
    """Posiciones de una palabra donde se puede partir la linea.

    Devuelve indices i tales que palabra[:i] queda al final de una linea y
    palabra[i:] empieza la siguiente. Se corta DESPUES de una barra o de un
    guion bajo -el separador se queda arriba, que es como se parten las
    direcciones web- y ANTES de un punto, para que la extension viaje entera
    y la linea no termine en un punto que se leeria como final de frase.
    Nunca se deja menos de dos caracteres a un lado.

    El corte ante el punto no es un adorno: sin el, la unica division posible
    de estacion_CODIGO_AAAAMMDD.jsonl dejaba AAAAMMDD.jsonl entero para la
    linea siguiente, seguian sobrando 66 pt en la primera y el rio se
    mantenia. Partiendo tambien ahi, lo que sobra baja a 14 pt.
    """
    # reportlab puede llegar aqui con bytes o con una subclase de str.
    if isinstance(palabra, bytes):
        palabra = palabra.decode("utf-8", "replace")
    palabra = str(palabra)
    if len(palabra) < CORTE_MINIMO or not RUTA_LARGA.match(palabra):
        return []
    # Sin barra ni guion bajo no es una ruta: es una palabra con un punto
    # dentro -queue.Queue, threading.Lock- que no conviene partir.
    if "/" not in palabra and "_" not in palabra and "\\" not in palabra:
        return []
    cortes = set()
    for i, ch in enumerate(palabra):
        if not (2 <= i <= len(palabra) - 3):
            continue
        if ch in "/_\\":
            cortes.add(i + 1)
        elif ch == ".":
            cortes.add(i)
    return sorted(cortes)


def con_puntos_de_corte(texto, marca="\u200b"):
    """El mismo texto con una marca invisible en cada punto de corte."""
    salida = []
    for palabra in re.split(r"(\s+)", texto):
        cortes = puntos_de_corte(palabra)
        if cortes:
            trozos = []
            anterior = 0
            for i in cortes:
                trozos.append(palabra[anterior:i])
                anterior = i
            trozos.append(palabra[anterior:])
            palabra = marca.join(trozos)
        salida.append(palabra)
    return "".join(salida)


def rutas_partibles_en_pdf():
    """Ensena a reportlab a partir rutas largas por sus separadores.

    reportlab ya sabe partir una direccion web al final de una linea, y lo hace
    sin anadir ningun caracter: parte la palabra y sigue en la siguiente. Lo
    hace desde _uri_split_pairs(), que solo reconoce como partible lo que
    encaja en su patron de URI -exige esquema o nombre de dominio-, de modo que
    scripts/generar_archivos_entrada.py no entraba. Aqui se amplia esa unica
    funcion para que tambien reconozca rutas y nombres de archivo, y se deja
    intacto todo lo demas: es reportlab quien sigue decidiendo cuando conviene
    partir y por donde.
    """
    from reportlab.platypus import paragraph as _rl
    if getattr(_rl, "_rutas_partibles", False):
        return
    original = _rl._uri_split_pairs

    def parejas(palabra):
        conocido = original(palabra)
        if conocido is not None:
            return conocido
        cortes = puntos_de_corte(palabra)
        if not cortes:
            return None
        texto = palabra.decode("utf-8", "replace") if isinstance(palabra, bytes)             else str(palabra)
        # De cabeza mas larga a mas corta: reportlab se queda con la primera
        # que le quepa, y lo que se busca es aprovechar la linea al maximo.
        return [(texto[:i], texto[i:]) for i in reversed(cortes)]

    _rl._uri_split_pairs = parejas
    _rl._rutas_partibles = True


# ---------------------------------------------------------------------------
# Preparacion de imagenes
# ---------------------------------------------------------------------------

def capturas_instalacion_real():
    """Capturas de la instalacion definitiva, hecha sobre Hyper-V.

    Son la UNICA evidencia grafica de instalacion del entregable, y archivos
    del repositorio. Si falta alguna, el generador se detiene.
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


def recortar_tras_hueco(origen, hueco_min=60, margen=8):
    """Deja solo el primer bloque de contenido de una captura ya recortada.

    recortar_consola() recorta al rectangulo que contiene TODO lo que no es
    fondo, y en las pantallas del instalador eso incluye adornos que estan
    lejos del texto: el menu de 01-menu-instalador-debian13.png termina en la
    fila 261 de 650 y el resto es fondo azul con el logotipo "debian 13", que
    no aporta nada al informe pero se lleva un cuarto de la pagina 3.

    En vez de recortar por una fraccion escrita a mano -que dejaria de valer si
    la captura cambia- se busca el primer hueco vertical de al menos hueco_min
    pixeles sin contenido y se corta ahi. Asi el criterio es "lo que sigue tras
    un vacio grande es decoracion, no la misma pantalla de texto", y se ajusta
    solo al tamano real de cada imagen. Si no hay ningun hueco de ese tamano la
    imagen se devuelve tal cual.
    """
    from PIL import Image

    origen = Path(origen)
    CACHE_RECORTES.mkdir(parents=True, exist_ok=True)
    destino = CACHE_RECORTES / ("util-" + origen.name)
    if destino.exists() and destino.stat().st_mtime >= origen.stat().st_mtime:
        return destino

    bandas = bandas_de_contenido(origen)
    corte = None
    for anterior, siguiente in zip(bandas, bandas[1:]):
        if siguiente[0] - anterior[1] >= hueco_min:
            corte = anterior[1]
            break
    im = Image.open(origen).convert("RGB")
    if corte is not None:
        im = im.crop((0, 0, im.size[0], min(im.size[1], corte + 1 + margen)))
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

    Se toma la mediana de las bandas de contenido cuyo alto es el de una linea
    de texto. La horquilla es 8-40 px por los dos extremos:

    - por abajo se descartan las bandas de 2 a 4 px, que NO son texto sino el
      cursor de bloque, los subrayados y las lineas de separacion. Sin ese
      filtro, la tira de una sola linea de la Figura 2 -que tiene una banda de
      texto de 20 px y otra de 4 px del cursor- daba una mediana de 12 px y se
      colocaba al doble de su tamano correcto.
    - por arriba se descartan los bloques graficos (barras de progreso,
      recuadros, logotipos), que no son lineas de texto.

    Basta UNA banda de texto para tener la medida: una captura de una sola
    linea es tan medible como una transcripcion de veinte. Si no hay ninguna
    -el cuadro de progreso del instalador es un unico bloque grafico- se
    devuelve None, y entonces la figura se coloca por densidad pura.
    """
    import statistics
    alturas = sorted(b[1] - b[0] + 1 for b in bandas_de_contenido(ruta)
                     if 8 <= (b[1] - b[0] + 1) <= 40)
    if not alturas:
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

    # Frase del informe, no transcripcion: la escribe este generador a partir
    # de la generacion leida arriba, y por eso va acentuada como el resto de la
    # prosa. Lo que se transcribe literal son las salidas de los cmdlets.
    h["hipervisor_implementado"] = (
        "Microsoft Hyper-V (máquina de Generación " + h["generacion"]
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
    d["sin_dependencias"] = acentuar_prosa_de_evidencia(buscar(
        t01, r"^(El proyecto usa exclusivamente la biblioteca estandar de Python)",
        "declaracion de dependencias", re.M))
    d["compilacion"] = acentuar_prosa_de_evidencia(
        buscar(t01, r"^(Los \w+ modulos compilan sin errores\.)$",
               "comprobacion de compilacion", re.M))

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
    # La fila que el informe imprime bajo el comando de identificacion tiene
    # que salir del bloque de identificacion, y no de cualquier parte de la
    # evidencia: buscarla por patron en todo el archivo hacia que se colara la
    # primera muestra del bloque "Evolucion durante la ejecucion", cuyos
    # valores son de otro instante y describen otro estado del proceso.
    bloque_ident = bloque_entre(
        t04, "--- Identificacion y recursos del proceso concurrente ---",
        "Interpretacion de cada campo", "identificacion del proceso")
    d["ps_cmd"] = buscar(bloque_ident, r"^(\$ ps -o [^\n]+)$", "comando ps", re.M)
    d["ps_cabecera"] = buscar(bloque_ident, r"^(\s*PID\s+PPID STAT.*)$",
                              "cabecera de ps", re.M)
    linea_ps = buscar(
        bloque_ident,
        r"^(\s*\d+\s+\d+\s+\S+\s+[\d.]+\s+[\d.]+\s+\d+\s+\d+\s+\d+\s+\S+\s+python3.*)$",
        "muestra de ps de la identificacion", re.M)
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
    # La explicacion del campo STAT se arma con la letra que trae ESTA muestra,
    # no con una escrita a mano: asi no puede describir un estado que la fila
    # transcrita no muestra.
    letra = d["stat"][0]
    if letra not in ESTADOS_PS:
        morir("la muestra de ps trae el estado %r y el informe no sabe explicarlo"
              % d["stat"])
    if not d["stat"].endswith("l"):
        morir("la muestra de ps trae el estado %r, sin la marca de multihilo"
              % d["stat"])
    otra = "S" if letra == "R" else "R"
    d["stat_explicacion"] = (
        d["stat"] + ": en el instante de la muestra el proceso estaba "
        + ESTADOS_PS[letra] + "; la l final indica que es multihilo. En otras "
        "muestras aparece " + otra + ", " + ESTADOS_PS[otra] + ".")

    d["ps_evolucion"] = bloque_entre(
        t04, "--- Evolucion durante la ejecucion",
        "--- Hilos del proceso", "evolucion de CPU").strip("\n")
    filas_evo = [ln.split() for ln in d["ps_evolucion"].splitlines() if ln.strip()]
    if len(filas_evo) < 2:
        morir("se esperaban al menos dos muestras sucesivas en el bloque de evolucion")
    d["evo_pcpu_inicial"] = filas_evo[0][3]
    d["evo_rss_inicial_mb"] = "%.1f" % (int(filas_evo[0][5]) / 1024.0)
    d["evo_nlwp_inicial"] = filas_evo[0][7]
    d["evo_pcpu_final"] = filas_evo[-1][3]
    d["evo_nlwp_final"] = filas_evo[-1][7]
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
    d["copias_demo"] = acentuar_prosa_de_evidencia(buscar(
        t04, r"^(Ambas(?: partes)? se ejecutan sobre COPIAS del proyecto.+)$",
        "declaracion de uso de copias", re.M))

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
    d["cierre_inventario"] = buscar(d["bitacora_final"],
                                    r"^(.*Inventario generado en .+)$",
                                    "linea de cierre del inventario", re.M)
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
    # Archivos que hay DE VERDAD bajo gestion_ambiental/ en el arbol entregado.
    # La cifra no se fija a mano ni se copia de la evidencia: el gestor produjo
    # los suyos y el control de versiones agrego ademas el marcador de la
    # carpeta vacia, de modo que el arbol trae uno mas. verificar_coherencia()
    # exige que esa sea toda la diferencia.
    d["archivos_gestion"] = sorted(
        ruta.relative_to(RAIZ).as_posix()
        for ruta in (RAIZ / "gestion_ambiental").rglob("*") if ruta.is_file())
    d["n_archivos_gestion"] = str(len(d["archivos_gestion"]))
    d["marcador_vacio"] = MARCADOR_CARPETA_VACIA.split("/", 1)[1]

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

    # Arbol entregado contra el listado congelado de la evidencia. Lo unico que
    # el arbol puede traer de mas es el marcador de la carpeta vacia; cualquier
    # otro archivo -o cualquier ausencia- significa que el informe describiria
    # una estructura que no es la que se entrega.
    faltan_gest = sorted(set(d["find_files"]) - set(d["archivos_gestion"]))
    sobran_gest = sorted(set(d["archivos_gestion"]) - set(d["find_files"]))
    if faltan_gest:
        problemas.append("la evidencia registra bajo gestion_ambiental/ archivos que "
                         "el arbol entregado no tiene: " + ", ".join(faltan_gest))
    if sobran_gest != [MARCADOR_CARPETA_VACIA]:
        problemas.append("bajo gestion_ambiental/ el arbol entregado trae %s de mas y "
                         "solo se admite %s"
                         % (sobran_gest or "nada", MARCADOR_CARPETA_VACIA))

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
    met = {sin_acentos(f[0]).strip(): f[1].strip() for f in d["tabla_metricas"]}
    # El README es una guia de ejecucion corta, no un duplicado del informe: solo
    # declara las cifras que un lector necesita para saber si su corrida salio bien.
    # Por eso aqui se comprueban esas y nada mas. Los datos finos -du -sh, el df de
    # la particion, los conteos de la prueba anti-sobrescritura y las anomalias
    # inyectadas- viven en el informe y en evidencias/, y este mismo verificador ya
    # los contrasta contra la evidencia mas arriba.
    for etiqueta, patron, esperado in (
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

    La figura se coloca por DENSIDAD, no por ancho. Colocar buscando el ancho
    de columna dejaba resoluciones efectivas dispares -de 72 a 240 dpi- entre
    capturas vecinas: la de 72 dpi salia dentada y la tira de una sola linea
    salia estirada de margen a margen, con su letra al doble que la de las
    figuras de su misma pagina.

    Hay dos formas de fijar la densidad, y se usa la mas informada de las dos:

    1. Si la captura tiene texto medible, se coloca al ancho en que la letra
       de dentro queda a texto_pt puntos. Esa es la densidad que ve el lector:
       dos capturas tomadas a resoluciones distintas acaban con la letra del
       mismo tamano, que es lo que hace homogeneo un bloque de figuras.
    2. Si no tiene texto medible -un cuadro de progreso, un recuadro grafico-
       no hay letra que igualar, y se coloca a DPI_DE_CAPTURA puntos por
       pulgada de imagen.

    Sobre las dos se aplican los mismos dos topes de siempre: el ancho de
    columna y el tamano nativo. El tope nativo impide el defecto contrario al
    de encoger: una captura pequena -05-primer-inicio.png son 374x68 px-
    colocada a ancho de columna completo queda a escala 1.288 y sale
    visiblemente pixelada. Ampliar no anade informacion a una imagen: solo
    estira pixeles. La equivalencia usada es un pixel de la imagen = un punto
    del PDF, que es la misma con la que se calcula ESCALAS, y por eso el tope
    nativo es exactamente la colocacion a 72 dpi.
    """
    from PIL import Image
    with Image.open(ruta) as im:
        ancho_px = im.size[0]
    factor = ancho_util / 481.89  # el ancho util del PDF, en las mismas unidades
    nativo = ancho_px * factor
    linea_px = altura_de_linea_px(ruta)
    if linea_px is None:
        deseado = ancho_px * (72.0 / DPI_DE_CAPTURA)
    else:
        objetivo = texto_pt if texto_pt is not None else TEXTO_DE_CAPTURA_PT
        deseado = ancho_px * (objetivo / float(linea_px))
    return min(ancho_util, nativo, deseado * factor)


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

# Tamano al que queda la letra de DENTRO de una captura una vez colocada. Es
# el de los bloques de transcripcion del propio informe (Courier 6.6 pt): la
# captura y el bloque de codigo que tiene al lado se leen igual. Queda con
# holgura por encima del minimo de 6.0 pt, de modo que la legibilidad no
# depende de este ajuste.
#
# Igualar la letra es lo que iguala la densidad APARENTE. Las capturas se
# tomaron a resoluciones distintas -la linea de una consola mide 12 px en unas
# y 22 px en otras-, asi que un mismo dpi para todas daria letras de tamanos
# distintos, que es justo el defecto que se quiere evitar. Fijando la letra, el
# dpi de cada figura sale derivado: dpi = 72 * linea_px / TEXTO_DE_CAPTURA_PT.
TEXTO_DE_CAPTURA_PT = 6.2

# Densidad con la que se colocan las capturas SIN texto medible, donde no hay
# letra que igualar (el cuadro de progreso del instalador). 140 dpi es el punto
# medio de la horquilla en la que una captura de pantalla ya no se ve dentada y
# todavia no gasta pagina de mas.
#
# Este numero es el suelo de densidad del informe, no su valor unico. Con la
# regla de la letra, las capturas quedan entre 131 dpi -las de la consola de la
# maquina, cuya linea mide 12 px- y 240 dpi -las del instalador, cuya linea
# mide 22 px-. Esa horquilla es inevitable y es la correcta: las capturas se
# tomaron a resoluciones distintas, y lo que el lector compara no es el dpi
# sino el tamano de la letra, que si queda igual en todas. Forzar un mismo dpi
# para las ocho solo se podria hacer sacando la letra de unas al doble que la
# de otras, o bajando a 3.6 pt las que ahora estan en 6.6, por debajo del
# minimo legible. Lo que si desaparece es el extremo de 72 dpi, que era el
# unico que se veia dentado.
DPI_DE_CAPTURA = 140.0

# ---------------------------------------------------------------------------
# Reglas de maquetacion compartidas por el PDF y el DOCX
#
# LINEAS_CODIGO_JUNTAS: hasta cuantas lineas de transcripcion se mantienen sin
# partir entre dos paginas. Una linea de codigo ocupa 8.1 pt (Courier 6.6 con
# interlinea 7.3 mas 0.8 de relleno), asi que 22 lineas son 178 pt: 6.3 cm, la
# cuarta parte del alto de texto de una A4 (26.6 cm). El limite sale de ahi. Un
# bloque que quepa en un cuarto de pagina se empuja entero antes que partirse,
# porque el hueco que deja nunca pasa de ese cuarto; uno mas largo SI se parte,
# porque empujarlo dejaria media pagina o mas en blanco y ademas reportlab
# acabaria arrastrandolo de pagina en pagina. Con este umbral, de los 19
# bloques del informe solo el volcado del inventario (24 lineas) puede partirse.
#
# LINEAS_CODIGO_MINIMAS: cuando un bloque largo se parte, cuantas lineas de
# salida tienen que quedar como minimo junto al comando para que la primera
# mitad signifique algo. Seis es lo que ocupa la salida corta tipica de este
# informe: un comando solo al pie de una pagina no es evidencia de nada.
#
# FILAS_TABLA_MINIMAS: por debajo de esto una tabla no se considera partible,
# y se empuja entera antes que dejar una cabecera con una fila colgando.
# ---------------------------------------------------------------------------
LINEAS_CODIGO_JUNTAS = 22
LINEAS_CODIGO_MINIMAS = 6
FILAS_TABLA_MINIMAS = 3

# Tablas de hasta esta altura (cabecera incluida) no se parten entre paginas.
# Ocho filas de las tablas de este informe miden unos 4 cm, que es el mismo
# limite con el que el PDF decide si una tabla se empuja entera o se deja
# partir. Por encima de eso partirla es preferible: la cabecera se repite y lo
# que se pierde es menos que la media pagina en blanco que costaria empujarla.
FILAS_TABLA_JUNTAS = 8

# Que parte del hueco sobrante de la portada va ARRIBA del bloque. El resto va
# abajo. 0.45 y no 0.5 porque el centro optico de una hoja esta algo por encima
# del centro geometrico: un bloque repartido al 50 % se ve caido. Las dos
# salidas usan el mismo numero, que es lo que hace que la portada del PDF y la
# del DOCX se parezcan.
ALTO_PORTADA = 0.45

# Alto de la caja de texto de una A4 con los margenes de este informe
# (29.7 - 1.5 - 1.6 cm), en puntos. El PDF lo obtiene de reportlab; el DOCX no
# tiene forma de medir, asi que lo necesita escrito.
ALTO_DE_TEXTO_PT = (29.7 - 1.5 - 1.6) / 2.54 * 72


def extracto_inventario(inv):
    """Extracto del inventario que respeta la ESTRUCTURA real del archivo.

    total_informes, informes_por_categoria, anomalias_registradas y
    errores_de_proceso NO son claves de primer nivel: cuelgan de "resumen".
    Presentarlas como si lo fueran describia un archivo distinto del entregado
    y contradecia al propio comando que el informe muestra encima. Aqui se
    vuelca el objeto completo y solo se eliden los tres subarboles largos, con
    la misma marca [...] que el resto del informe, de modo que las claves de
    primer nivel quedan TODAS a la vista.
    """
    marca = "__ELISION__"
    resumido = {}
    for clave, valor in inv.items():
        if clave in ("informes", "alertas_por_indicador", "anomalias"):
            resumido[clave] = marca
        else:
            resumido[clave] = valor
    texto = json.dumps(resumido, indent=2, ensure_ascii=False)
    return texto.replace('"' + marca + '"', "{ [...] }")


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
    b.append(h1("intro", "Introducción y objetivos"))
    b.append(p(
        "Este informe documenta el traslado de la solución de monitoreo ambiental del "
        "Equipo 07 a un ambiente Debian GNU/Linux, como exige la Parte 2 de la Solemne 01 "
        "práctica (Forma B): instalación de la máquina virtual, preparación del entorno, "
        "ejecución de las dos versiones de la Parte 1 sobre el mismo conjunto de datos, "
        "gestor de incidencias y observación de procesos, memoria y sistema de archivos. "
        "Todas las cifras se midieron dentro de la máquina virtual en una ÚNICA corrida y "
        "quedaron en evidencias/. El documento lo genera scripts/generar_informe.py, que "
        "lee esos archivos, los compara con el estado real del árbol entregado y se "
        "detiene sin emitir nada si ambas fuentes no coinciden."))

    b.append(h2("objetivos", "Objetivo de cada componente"))
    b.append(ul([
        "src/secuencial.py: procesar un archivo JSONL completo antes de pasar al "
        "siguiente, sin hilos. Define el contrato de validación y es la referencia contra "
        "la cual se compara el resultado concurrente.",
        "src/concurrente.py: procesar el mismo conjunto repartiendo los archivos entre "
        + d["t_con"]["trabajadores"] + " trabajadores threading.Thread mediante una "
        "queue.Queue compartida, protegiendo acumuladores globales y bitácora de alertas "
        "con mutex threading.Lock.",
        "src/gestor_incidencias.py: clasificar los informes por cantidad de alertas, "
        "resguardar el resumen, separar alertas por indicador, generar el inventario JSON "
        "y mantener una bitácora trazable sin detenerse ante entradas defectuosas.",
    ]))

    b.append(h2("flujo", "Flujo del sistema de principio a fin"))
    b.append(ol([
        "Generación de datos: scripts/generar_archivos_entrada.py crea "
        + d["archivos_generados"] + " archivos estacion_CODIGO_AAAAMMDD.jsonl en "
        "entrada/ con semilla fija " + d["semilla"] + ", de modo que la entrada es "
        "idéntica en cualquier máquina.",
        "Ingesta y validación: cada línea JSONL se valida en formato, tipos, rangos y "
        "timestamp; la que no cumple se descarta como inválida sin detener el proceso.",
        "Procesamiento: la versión secuencial recorre los archivos uno tras otro; la "
        "concurrente los encola en una queue.Queue y los reparte entre "
        + d["t_con"]["trabajadores"] + " hilos que acumulan resultados locales y sólo "
        "actualizan las variables globales dentro de una sección crítica.",
        "Detección de alertas: cada medición válida se contrasta con los umbrales y la "
        "alerta se escribe en alertas/alertas_detectadas.log con el formato "
        "archivo;estacion;timestamp;indicador;valor;umbral.",
        "Salida: un informe por archivo en salida/informe_*.txt y el consolidado "
        "salida/resumen_ambiental.txt.",
        "Gestión de incidencias: los informes se MUEVEN a gestion_ambiental/sin_alertas/, "
        "con_alertas/ o criticas/ según sus alertas; el resumen se COPIA como "
        "resumen_resguardado.txt; las alertas se separan por indicador; se generan "
        "inventario_ambiental.json y logs/gestion_ambiental.log. Por eso salida/ queda "
        "sólo con el resumen: ver " + ref("leeme") + ".",
    ]))

    b.append(h2("validacion", "Reglas de validación y umbrales de alerta"))
    b.append(tabla(
        ["Campo", "Rango válido", "Indicador", "Condición de alerta"],
        [[RANGOS_VALIDOS[i][0], RANGOS_VALIDOS[i][1],
          UMBRALES[i][0] if i < len(UMBRALES) else "",
          UMBRALES[i][1] if i < len(UMBRALES) else ""]
         for i in range(len(RANGOS_VALIDOS))],
        anchos=[3.0, 5.0, 3.0, 5.5]))
    b.append(p(
        "Ambas versiones comparten estas constantes, y ésa es la razón de fondo por la "
        "cual sus métricas deben coincidir hasta el último decimal: la concurrencia "
        "cambia el orden en que se hace el trabajo, no el criterio con que se decide."))

    # ------------------------------------------------- 2. Instalacion Debian
    b.append(h1("instalacion", "Instalación de Debian 13 en la máquina virtual"))

    b.append(h2("iso", "Descarga y verificación del ISO"))
    b.append(p(
        "Se descargó la imagen de instalación por red " + hv["iso_nombre"] + " ("
        + hv["iso_tamano"] + ") desde el sitio oficial de Debian y se verificó su "
        "integridad comparando su SHA256 con el publicado en SHA256SUMS, lo que garantiza "
        "que el medio no fue alterado ni quedó truncado. La salida es la del anfitrión y "
        "está archivada en evidencias/hyperv/configuracion-vm-hyperv.txt:"))
    b.append(code("PS> Get-FileHash " + hv["iso_nombre"] + " -Algorithm SHA256",
                  "Hash calculado : " + hv["iso_sha_calculado"]
                  + "\nHash publicado : " + hv["iso_sha_publicado"]
                  + "\n  Fuente: " + hv["iso_fuente"]
                  + "\nCoincide       : " + hv["iso_coincide"]))

    b.append(h2("hipervisor", "Declaración del cambio de hipervisor respecto del hito"))
    b.append(nota(
        "Declaración explícita. En el hito el equipo planificó usar "
        + HIPERVISOR_PLANIFICADO + " y la implementación final se hizo sobre "
        + hv["hipervisor_implementado"] + ". Se hace constar porque afecta lo declarado en "
        "el hito; la pauta admite la situación, ya que la instalación \"puede hacerlo en "
        "VirtualBox, VMware u otro hipervisor\". Los recursos comprometidos se respetaron "
        "exactamente, sin rebajar ninguno."))
    b.append(tabla(
        ["Recurso", "Mínimo de la pauta", "Comprometido en el hito",
         "Implementado y verificado"],
        [["Hipervisor", "VirtualBox, VMware u otro", HIPERVISOR_PLANIFICADO,
          "Microsoft Hyper-V, Generación " + hv["generacion"]],
         ["RAM", "2 GB", "4 GB",
          hv["ram_gb"] + " GB fijos (memoria dinámica: " + hv["ram_dinamica"] + "); "
          + d["memoria"].splitlines()[1].split()[1] + " útiles según free -h"],
         ["Procesadores virtuales", "2", "4",
          hv["vcpu"] + " vCPU sobre " + d["modelo_cpu"]],
         ["Disco virtual", "20 GB", "25 GB",
          hv["vhd_gb"] + " GB " + hv["vhd_tipo"] + " (ocupa " + hv["vhd_real_gb"]
          + " GB reales en el anfitrión)"],
         ["Red", "NAT", "NAT con conectividad",
          hv["switch"] + " (" + hv["switch_tipo"] + "), IP " + d["ip_eth0"] + " por DHCP"],
         ["Sistema", "Debian 13, 64 bits", "Debian 13, 64 bits",
          d["pretty_name"] + ", versión " + d["version_debian"]]],
        anchos=[2.5, 2.9, 2.9, 8.1]))
    b.append(p(
        "El propio sistema instalado confirma sobre qué hipervisor corre, de modo que la "
        "declaración es verificable y no depende de la palabra del equipo:"))
    b.append(code("$ systemd-detect-virt ; lscpu | grep 'Hypervisor vendor' ; uname -a",
                  d["hipervisor_detectado"]
                  + "\nHypervisor vendor:                       " + d["hipervisor_vendor"]
                  + "\n" + d["kernel"]))
    # De la etapa previa sobre VirtualBox NO se conserva ninguna captura en el
    # entregable. Las que habia correspondian a otra maquina -otro usuario, otro
    # nombre de host- y sus cifras contradecian a las de la maquina entregada,
    # de modo que ilustrar la configuracion de recursos con ellas habria sido
    # peor que no ilustrarla. El cambio de hipervisor se sigue declarando en
    # prosa, que es lo que exige la honestidad, y los recursos quedan
    # verificados contra los cmdlets del hipervisor en la tabla de arriba.
    b.append(p(
        "De la etapa previa sobre " + HIPERVISOR_PLANIFICADO + " no se conserva ninguna "
        "captura en el entregable, y por eso esta subsección no lleva figura: las "
        "imágenes de esa etapa correspondían a otra máquina, con otro usuario y otro "
        "nombre de host, de modo que presentarlas habría contradicho lo que documenta el "
        "resto del informe. La única evidencia gráfica de instalación es la de la máquina "
        "definitiva sobre " + hv["hipervisor_implementado"] + " (" + ref("instalador")
        + " y " + ref("primer_inicio") + "); los recursos comprometidos en el hito ya "
        "quedaron verificados arriba contra los cmdlets del hipervisor."))

    b.append(h2("instalador", "Instalación, usuario y particionado del disco virtual"))
    b.append(p(
        "Las capturas de esta sección y de la siguiente son de la instalación REAL de la "
        "máquina definitiva sobre Hyper-V. Se creó de Generación " + hv["generacion"]
        + ", con firmware BIOS heredado, y su orden de arranque ("
        + hv["orden_arranque"] + ") explica que el instalador arranque en modo BIOS y no "
        "UEFI. Se creó el usuario sin privilegios " + d["usuario_vm"] + ", con acceso a "
        "sudo, y el host quedó como " + d["hostname_vm"] + ", de modo que cualquier salida "
        "de este informe se atribuye sin ambigüedad a la máquina del equipo."))
    b.append(fig(recortar_tras_hueco(imgs_hv["01-menu-instalador-debian13.png"]),
                 "Menú del instalador de Debian 13 en modo BIOS, en la máquina "
                 "de Generación " + hv["generacion"] + " de Hyper-V. Instalación real de "
                 "la máquina definitiva.", texto_pt=TEXTO_DE_CAPTURA_PT))
    b.append(fig(imgs_hv["02-linea-de-arranque-preseed.png"],
                 "Línea de arranque del instalador con el archivo de "
                 "preconfiguración (preseed), que fija idioma, teclado, zona horaria, "
                 "particionado guiado y el conjunto mínimo de tareas. Usar preseed hace "
                 "la instalación reproducible: el mismo archivo produce la misma "
                 "máquina."))
    b.append(fig(imgs_hv["03-instalacion-sistema-base.png"],
                 "Instalación del sistema base: descarga e instalación de paquetes desde "
                 "la réplica de red. La etapa final del instalador -gestor de arranque y "
                 "copia de la configuración de red al sistema instalado- es la misma "
                 "pantalla de progreso y está archivada en "
                 "evidencias/fotos/instalacion-hyperv/04-instalacion-final.png."))
    b.append(p("El particionado se aplicó sobre el disco de " + hv["vhd_gb"]
               + " GB con el esquema guiado \"todo en una partición\":"))
    b.append(code("$ lsblk", d["lsblk"]))
    b.append(ul([
        "sda1 (23.7 GB, ext4) es la partición raíz montada en /; sda2 es sólo el "
        "contenedor lógico de la extendida y sda5 (1.3 GB) es el área de intercambio, "
        "que el núcleo usa para descargar páginas cuando la memoria física escasea.",
        "sr0 es la unidad óptica virtual desde la que se montó el ISO de instalación; "
        "terminada la instalación el ISO fue expulsado, como consta en la salida de "
        "Get-VMDvdDrive del anfitrión.",
    ]))

    b.append(h2("primer_inicio", "Primer inicio y conectividad"))
    b.append(fig(imgs_hv["05-primer-inicio.png"],
                 "Primer inicio del sistema recién instalado: el núcleo llega a "
                 "la consola tty1 y presenta el indicador de acceso del equipo "
                 + d["hostname_vm"] + ". La captura muestra el indicador de acceso, no la "
                 "sesión ya iniciada; que el acceso con el usuario " + d["usuario_vm"]
                 + " funciona lo demuestra el resto del informe, cuyas salidas fueron "
                 "obtenidas en esa sesión."))
    b.append(p(
        "eth0 obtiene su dirección por DHCP desde el conmutador " + hv["switch"]
        + " y la ruta por omisión apunta a " + d["gateway"] + ". " + hv["nat"]
        + " El anfitrión ve la misma dirección (" + hv["ip_hyperv"] + ", MAC "
        + hv["mac"] + "), lo que cierra la trazabilidad entre lo que declara Hyper-V y lo "
        "que observa el huésped. Con esa configuración la máquina alcanza las réplicas de "
        "Debian, como confirman la resolución de deb.debian.org y las líneas \"Hit\" de "
        "apt de la sección siguiente."))
    b.append(code("$ ip -4 -br addr show ; ip route ; getent hosts deb.debian.org",
                  d["red"] + "\n" + d["dns"]))

    # --------------------------------------------- 3. Preparacion del ambiente
    b.append(h1("ambiente", "Preparación del ambiente"))
    b.append(p(
        "El punto 2.a de la pauta exige actualizar el sistema antes de trabajar. La "
        "salida confirma que los tres orígenes de paquetes están accesibles (trixie, "
        "trixie-security y trixie-updates) y que el sistema quedó al día."))
    # Se transcriben las lineas evaluables (origenes accesibles y estado final)
    # y se marca la elision del resto con [...], como en el resto del informe.
    b.append(code("$ sudo apt update ; sudo apt upgrade -y",
                  "\n".join(ln for ln in d["apt_update"].splitlines()
                            if ln.startswith(("Hit:", "All packages")))
                  + "\n\n[...]\n"
                  + "\n".join(d["apt_upgrade"].splitlines()[-3:])))
    b.append(p(
        "El punto 2.b pide instalar el entorno del lenguaje usado en la Parte 1. El "
        "proyecto está en Python y Debian 13 ya trae " + d["python_version"] + " en el "
        "sistema base, de modo que no hubo que compilar ni agregar repositorios externos. "
        + d["sin_dependencias"] + ": sin pip, sin entorno virtual y sin archivo de "
        "requisitos, decisión deliberada porque hace reproducible la ejecución incluso sin "
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
    b.append(h1("parte1", "Ejecución de la Parte 1 en Debian"))
    b.append(p(
        "El punto 2.c exige ejecutar ambas versiones sobre el mismo conjunto de archivos "
        ".jsonl y comprobar que seis métricas coinciden. Para que \"el mismo conjunto\" sea "
        "verificable, la entrada se genera con semilla fija y se le calcula una suma de "
        "verificación antes de cada corrida."))
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

    b.append(h2("metricas", "Las seis métricas exigidas por la pauta"))
    b.append(tabla(["Métrica", "Secuencial", "Concurrente", "Coincide"],
                   d["tabla_metricas"], anchos=[5.0, 4.0, 4.0, 3.5]))
    b.append(p("Métricas comparadas: " + d["metricas_comparadas"] + ". "
               + d["resultado_consistencia"] + ". El resumen consolidado agrega "
               + res["Lineas leidas"] + " líneas leídas, temperatura promedio "
               + res.get("Temperatura promedio global", "?") + ", humedad promedio "
               + res.get("Humedad promedio global", "?") + " y estación con más alertas "
               + res.get("Estacion con mayor cantidad de alertas", "?")
               + ". El entregable conserva el resumen de la corrida concurrente en "
               "salida/resumen_ambiental.txt."))

    b.append(h2("consistencia", "Comprobaciones adicionales de consistencia"))
    b.append(p(
        "Seis totales podrían coincidir por casualidad si dos errores se compensaran, de "
        "modo que se agregaron dos comprobaciones más exigentes que las que pide la pauta: "
        "los " + d["informes_comparados"] + " informes individuales comparados uno por "
        "uno, con " + d["informes_con_diferencias"] + " diferencias, y el contenido "
        "completo de la bitácora de alertas, con " + d["alertas_sec"] + " alertas en el "
        "secuencial y " + d["alertas_con"] + " en el concurrente. El orden de escritura "
        "del concurrente varía por diseño, porque " + d["t_con"]["trabajadores"]
        + " hilos vuelcan sus búferes en momentos distintos; por eso el volcado final se "
        "ordena por nombre de archivo. " + d["veredicto_log"]))
    b.append(nota(
        "Resultado incómodo que se declara tal cual se midió: la concurrente fue más "
        "LENTA que la secuencial en este conjunto (" + d["t_sec"]["real"] + " frente a "
        + d["t_con"]["real"] + " de tiempo real; " + d["t_sec"]["interno"] + " s frente a "
        + d["t_con"]["interno"] + " s de tiempo interno). La explicación está en "
        + ref("concurrencia") + "."))
    # El pie de esta figura NO se escribe a ciegas: se comprueba que la captura
    # muestre de verdad las filas de la tabla. Si la tabla salio vacia en la
    # pantalla (ha ocurrido), el pie no puede prometer "las seis metricas
    # coinciden", porque la imagen no lo muestra; la afirmacion sigue
    # respaldada por la tabla del cuerpo, y el pie lo dice asi.
    captura_7 = recortar_consola(FOTOS / "debian-01-consistencia-parte1.png")
    filas_7 = filas_visibles_de_la_tabla(captura_7)
    if filas_7 is not None and filas_7 >= 4:
        pie_7 = ("Comprobación de consistencia en la consola de la máquina "
                 "virtual: las seis métricas coinciden, los " + d["informes_comparados"]
                 + " informes individuales son idénticos y la bitácora de alertas tiene "
                 + d["alertas_sec"] + " alertas en ambas versiones.")
    else:
        pie_7 = ("Comprobación de consistencia ejecutada en la consola de la "
                 "máquina virtual. En esta captura el cuerpo de la tabla comparativa no "
                 "alcanzó a quedar en pantalla: se ven el encabezado, los separadores y "
                 "el resultado final (los " + d["informes_comparados"] + " informes "
                 "individuales idénticos y las " + d["alertas_sec"] + " alertas de la "
                 "bitácora en ambas versiones). La comparación métrica por métrica es la "
                 "de la tabla de " + ref("metricas") + ", tomada de "
                 "evidencias/debian/03-ejecucion-parte1.txt.")
    # Esta captura es un recorte pequeno (544x294): a ancho de columna completo
    # se ampliaria hasta dejar su letra por encima de la del cuerpo del informe
    # y, sobre todo, desbordaria la pagina dejando un tercio en blanco. Se
    # coloca al tamano en que su texto queda a 7.2 pt, por encima del umbral.
    b.append(fig(captura_7, pie_7, texto_pt=TEXTO_DE_CAPTURA_PT))

    # ----------------------------------------------- 5. Gestor de incidencias
    b.append(h1("gestor", "Gestor de incidencias"))
    b.append(p(
        "El punto 2.d enumera trece requisitos para src/gestor_incidencias.py. La tabla "
        "los recorre uno por uno con la evidencia que los respalda; todas las cifras "
        "provienen del estado real del proyecto y de los archivos de evidencia, y el "
        "generador comprueba que ambas fuentes coincidan antes de escribir esta página."))
    b.append(tabla(
        ["N", "Requisito de la pauta", "Evidencia obtenida"],
        [["1", "Detectar informes informe_CODIGO_AAAAMMDD.txt",
          "Patrón ^informe_[A-Za-z0-9]+_\\d{8}(_v\\d+)?\\.txt$; "
          + str(inv["resumen"]["total_informes"]) + " informes detectados"],
         ["2", "Leer la cantidad de alertas de cada informe",
          "Línea \"Alertas detectadas: N\" de cada informe; suma "
          + str(inv["total_alertas"]) + " alertas"],
         ["3", "Mover informes con 0 alertas a sin_alertas/",
          "Rama implementada; queda en " + str(d["n_sin"]) + " informes (ver "
          + ref("sin_alertas") + ") y se demostró con " + d["ctl_informe"]],
         ["4", "Mover informes con 1 a 3 alertas a con_alertas/",
          str(d["n_con"]) + " informes en gestion_ambiental/con_alertas/"],
         ["5", "Mover informes con 4 o más alertas a criticas/",
          str(d["n_crit"]) + " informes en gestion_ambiental/criticas/"],
         ["6", "Evitar sobrescritura mediante nombres _v1, _v2",
          "Tres corridas acumulan " + d["encontrados_v"] + " informes, "
          + d["con_sufijo_v"] + " de ellos con sufijo _v1 o _v2, sin pérdidas (ver "
          + ref("anti") + ")"],
         ["7", "Copiar el resumen como resumen_resguardado.txt",
          "Copia de " + str(d["resguardo_bytes"])
          + " bytes, idéntica al original que permanece en salida/"],
         ["8", "Leer alertas/alertas_detectadas.log",
          str(d["n_alertas_log"]) + " líneas del formato "
          "archivo;estacion;timestamp;indicador;valor;umbral"],
         ["9", "Separar alertas por indicador",
          "temperatura " + str(ind["temperatura"]) + ", humedad " + str(ind["humedad"])
          + ", pm25 " + str(ind["pm25"]) + ", ruido " + str(ind["ruido"])],
         ["10", "Registrar alertas inválidas o desconocidas",
          d["anomalias_total"] + " anomalías registradas al inyectar fallas (ver "
          + ref("errores") + ")"],
         ["11", "Generar inventario_ambiental.json",
          "Archivo de " + str(d["inventario_bytes"]) + " bytes, "
          + d["json_valido"] + " con " + d["n_claves_inventario"]
          + " claves de primer nivel"],
         ["12", "Generar logs/gestion_ambiental.log",
          "Bitácora de " + d["ev_n_bitacora"] + " líneas y " + d["stat_size"]
          + " bytes, con marca de tiempo por evento"],
         ["13", "Controlar al menos un error sin detenerse",
          "Con cuatro fallas el gestor terminó con código " + d["gestor_exit_error"]
          + " y conservó las " + d["inventario_tras_error"] + " alertas (ver "
          + ref("errores") + ")"]],
        anchos=[0.8, 6.2, 9.5]))

    b.append(h2("estructura", "Estructura generada"))
    b.append(code("$ find gestion_ambiental logs -type d | sort",
                  d["find_dirs"]))
    # La cifra se lee del arbol entregado, no se fija a mano: el gestor produjo
    # los archivos de la corrida y el control de versiones agrego ademas el
    # marcador sin_alertas/.gitkeep, sin el cual git no conservaria una carpeta
    # que quedo legitimamente vacia. Por eso el arbol trae uno mas que la
    # evidencia congelada, y recolectar_datos() comprueba que la diferencia sea
    # exactamente ese marcador.
    b.append(p("Bajo gestion_ambiental/ quedaron " + d["n_archivos_gestion"]
               + " archivos: los " + d["find_files_total"] + " que produjo el gestor -los "
               + str(inv["resumen"]["total_informes"]) + " informes clasificados, los "
               "cuatro registros por indicador, el inventario y el resumen resguardado, "
               "los mismos que registra la evidencia- más el marcador "
               + d["marcador_vacio"] + ", con el que el control de versiones conserva una "
               "carpeta que quedó legítimamente vacía (" + ref("sin_alertas")
               + "). El gestor terminó con código " + d["gestor_exit"] + " en su única "
               "ejecución sobre el árbol entregado."))
    b.append(fig(recortar_consola(FOTOS / "debian-03-estructura-gestor.png"),
                 "Estructura generada por el gestor y cuadratura de las alertas por "
                 "indicador, en la máquina virtual.", texto_pt=TEXTO_DE_CAPTURA_PT))

    b.append(h2("cuadratura", "Inventario y cuadratura de las alertas"))
    b.append(p(
        "El valor del inventario no está en existir sino en cuadrar: la suma de las "
        "alertas por indicador debe ser igual al total del inventario, al número de líneas "
        "de la bitácora de alertas y al total del resumen consolidado. El generador "
        "verifica esa igualdad y además que el número leído en el repositorio sea el mismo "
        "de la evidencia congelada; si alguna comprobación falla, se detiene."))
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
    # El extracto reproduce la ESTRUCTURA real del archivo: las cifras de
    # cuadratura cuelgan de "resumen", no del primer nivel. Confundirlo hacia
    # que el informe mostrara como claves de primer nivel cuatro que no lo son.
    # Los tres subarboles largos se eliden con la misma marca [...] que el
    # resto del informe; las ocho claves de primer nivel quedan todas a la
    # vista, que es lo que el requisito 11 pide poder comprobar.
    b.append(code("$ python3 -m json.tool gestion_ambiental/inventario_ambiental.json",
                  extracto_inventario(inv)))
    b.append(p(
        "total_alertas aparece dos veces porque el inventario lo declara dentro de "
        "resumen y también en la raíz; ambos valores son el mismo número que cuadra la "
        "tabla anterior, y la bitácora cierra la corrida declarando ese mismo par:"))
    b.append(code("$ grep 'Inventario generado' logs/gestion_ambiental.log",
                  d["cierre_inventario"]))

    b.append(h2("sin_alertas", "Por qué sin_alertas/ quedó vacía"))
    b.append(p(
        "La carpeta sin_alertas/ contiene " + str(d["n_sin"]) + " informes y no es un "
        "defecto del gestor: la pauta de la Parte 1 exige que cada archivo de entrada "
        "produzca al menos una alerta, de modo que ningún informe puede tener cero y la "
        "rama queda legítimamente vacía. La rama existe y funciona: con el informe de "
        "control " + d["ctl_informe"] + ", construido con cero alertas, el gestor lo movió "
        "a sin_alertas/ y la clasificación quedó en " + d["clasificacion_ctl"] + ". Esa "
        "demostración se hizo sobre una COPIA, por lo que el árbol entregado conserva sus "
        + str(d["n_sin"]) + " informes en esa rama."))

    b.append(h2("leeme", "Por qué salida/ sólo conserva el resumen"))
    b.append(p(
        "En el árbol entregado, salida/ contiene sólo resumen_ambiental.txt y LEEME.txt. "
        "No falta nada: los " + d["informes_generados"] + " informes individuales SÍ se "
        "generaron ahí (src/concurrente.py) y el gestor los MOVIÓ a sus carpetas de "
        "clasificación, que es lo que exigen los puntos 3, 4 y 5 de la pauta. Mover y no "
        "copiar se explica en " + ref("mover_copiar") + ". Antes de ejecutar el gestor se "
        "respaldó salida/ completo, de modo que el estado que dejó la Parte 1 también es "
        "auditable:"))
    b.append(ul([
        "salida/resumen_ambiental.txt (" + str(d["resumen_bytes"]) + " bytes): el "
        "consolidado que la pauta pide mantener en su lugar.",
        "salida/LEEME.txt: explica esta situación dentro del propio entregable e indica "
        "cómo regenerar los informes.",
        "evidencias/salida_parte1/: " + d["respaldo_parte1"] + " archivos, es decir los "
        + d["informes_generados"] + " informes individuales más el resumen, tal como los "
        "dejó la Parte 1 antes de la clasificación.",
    ]))

    # -------------------------------- 6. Observacion de procesos y archivos
    b.append(h1("procesos", "Observación de procesos y del sistema de archivos"))
    b.append(nota(
        "Nota metodológica sobre la carga amplificada. El conjunto oficial de "
        + d["archivos_generados"] + " archivos se procesa en milisegundos y ps no alcanza "
        "a tomar una muestra útil. Por eso la observación se hizo en dos partes: (A) el "
        "consumo de la corrida OFICIAL con /usr/bin/time -v, que no necesita muestreo; y "
        "(B) la observación en vivo con ps sobre una CARGA AMPLIFICADA de "
        + d["carga_archivos"] + " archivos (" + d["carga_lineas"] + " líneas), los mismos "
        "archivos oficiales replicados con otras fechas válidas. " + d["copias_demo"]
        + " Toda métrica entregable sale del conjunto oficial; la carga amplificada sólo "
        "hace que el proceso viva lo suficiente para observarlo ("
        + d["muestras_ps"] + " muestras)."))

    b.append(h2("recursos", "Consumo de recursos de la corrida oficial"))
    b.append(code("$ /usr/bin/time -v python3 src/concurrente.py", d["time_v"]))
    b.append(ul([
        "Percent of CPU " + d["pct_cpu_oficial"] + " sobre centésimas de segundo: con "
        "sólo " + d["archivos_generados"] + " archivos el proceso apenas alcanza a "
        "repartir trabajo antes de terminar.",
        "Maximum resident set size " + d["max_rss"] + " KB (unos "
        + "%.1f" % (int(d["max_rss"]) / 1024.0) + " MB de memoria física): el costo del "
        "intérprete de Python más los acumuladores del programa.",
        d["ctx_vol"] + " cambios de contexto voluntarios y " + d["ctx_invol"]
        + " involuntarios. Los voluntarios son el proceso cediendo la CPU por su cuenta "
        "al esperar entrada/salida o un mutex; los involuntarios son el planificador "
        "expulsándolo al agotar su cuanto. " + d["fs_outputs"] + " bloques escritos "
        "corresponden a los informes, el resumen y la bitácora de alertas.",
    ]))

    b.append(h2("identificacion", "Identificación del proceso concurrente"))
    b.append(code(d["ps_cmd"], d["ps_cabecera"] + "\n" + d["ps_linea"]))
    b.append(tabla(
        ["Campo", "Valor", "Qué significa en este proceso concreto"],
        [["PID", d["pid"],
          "Identificador que el núcleo asignó al proceso que ejecuta "
          "python3 src/concurrente.py; es único mientras el proceso vive."],
         ["PPID", d["ppid"],
          "Proceso padre: " + d["padre"].splitlines()[-1].strip()
          + ". Es el intérprete de órdenes desde el que se lanzó la ejecución."],
         ["STAT", d["stat"], d["stat_explicacion"]],
         ["%CPU", d["pcpu"] + " (máx. " + d["pcpu_max"] + ")",
          "ps lo promedia sobre toda la vida del proceso, y esta muestra aún arrastra el "
          "arranque. El trabajo simultáneo en más de un núcleo lo prueba el máximo "
          "observado, " + d["pcpu_max"] + ": un proceso de un solo hilo no pasa de 100 %."],
         ["%MEM", d["pmem"],
          "Fracción de los " + hv["ram_gb"] + " GB de memoria física ocupada por el "
          "proceso; es baja porque el programa procesa por flujo y no carga todo en "
          "memoria."],
         ["RSS", d["rss"] + " KB",
          "Memoria residente: la parte del proceso efectivamente cargada en RAM, que es "
          "lo que cuesta de verdad en memoria física."],
         ["VSZ", d["vsz"] + " KB",
          "Memoria virtual reservada, unas trece veces el RSS: incluye bibliotecas "
          "compartidas, pilas de hilos y regiones que nunca se tocaron, por lo que no "
          "debe leerse como consumo real."],
         ["NLWP", d["nlwp"],
          "Hilos vivos: 1 principal + " + d["t_con"]["trabajadores"]
          + " trabajadores, exactamente los que declara el programa."]],
        anchos=[1.5, 2.6, 12.4]))

    b.append(h2("hilos", "Hilos y vista del núcleo"))
    b.append(code("$ ps -L -o pid,tid,stat,%cpu,comm -p " + d["pid"] + "\n"
                  "$ grep -E 'State|Threads|VmRSS' /proc/" + d["pid"] + "/status",
                  d["ps_hilos"] + "\n" + d["proc_status"]))
    b.append(p(
        "Los cuatro identificadores de hilo comparten el mismo PID pero tienen TID "
        "distintos: el primero coincide con el PID y es el hilo principal, que encola y "
        "espera; los otros " + d["t_con"]["trabajadores"] + " consumen de la cola. El "
        "núcleo confirma Threads: " + d["proc_threads"] + " y VmRSS: " + d["proc_vmrss"]
        + " kB, coherente con ps, y State: " + d["proc_estado"] + ". Que el estado sea "
        "\"sleeping\" en una muestra y \"running\" en otra no es contradictorio: los hilos "
        "alternan entre cálculo y espera por entrada/salida."))
    # La evolucion es un bloque APARTE, con su propio encabezado: sus filas
    # pertenecen a muestras sucesivas del mismo proceso y no a la muestra de
    # identificacion de 6.2. Mezclarlas era lo que hacia que el informe
    # imprimiera bajo el comando de identificacion una fila que la evidencia
    # registra en este otro bloque.
    b.append(h2("evolucion", "Evolución del proceso durante la ejecución"))
    b.append(code(d["ps_cmd"] + "   # repetido sobre el mismo PID",
                  d["ps_cabecera"] + "\n" + d["ps_evolucion"]))
    b.append(p(
        "La primera muestra fue tomada antes de que arrancaran los trabajadores: un solo "
        "hilo (NLWP " + d["evo_nlwp_inicial"] + "), " + d["evo_pcpu_inicial"]
        + " % de CPU y " + d["evo_rss_inicial_mb"] + " MB residentes. En las siguientes "
        "NLWP sube a " + d["nlwp"] + ", el %CPU pasa de 100 y el RSS crece de forma "
        "sostenida a medida que se llenan los acumuladores; en la última los trabajadores "
        "ya terminaron y NLWP vuelve a " + d["evo_nlwp_final"] + " con el %CPU acumulado "
        "en " + d["evo_pcpu_final"] + ". Esa progresión es la concurrencia hecha visible."))
    # Pie honesto de esta figura: la carga es la MISMA que la transcrita (los
    # mismos archivos y lineas que declara el encabezado de la pantalla); lo
    # que cambia es la corrida, y con ella el PID, el RSS y el %CPU maximo.
    # Omitir el %CPU -que es justo el valor mas llamativo- para enumerar solo
    # PID y RSS descargaba de mas al lector.
    b.append(fig(recortar_consola(FOTOS / "debian-02-observacion-procesos.png"),
                 "Observación en vivo del proceso concurrente en la consola de la "
                 "máquina virtual. Es otra corrida de la misma prueba y con la MISMA "
                 "carga amplificada: los " + d["carga_archivos"] + " archivos y "
                 + d["carga_lineas"] + " líneas que declara el encabezado de la pantalla "
                 "son los de la transcripción de arriba. Lo que cambia es la corrida, y "
                 "por eso difieren el PID, el RSS y el %CPU máximo, que aquí llega a "
                 + FIG_OBSERVACION_PCPU + " frente a los " + d["pcpu_max"]
                 + " transcritos. Coincide en lo que importa: " + d["nlwp"] + " hilos "
                 "vivos, %CPU por encima de 100 y el mismo du -sh de "
                 + d["du_proyecto"] + " del proyecto.", texto_pt=TEXTO_DE_CAPTURA_PT))

    b.append(h2("memoria_fs",
                "Memoria del sistema, espacio disponible y tamaño del proyecto"))
    mem = d["free_h"].splitlines()[1].split()
    dfr = d["df_raiz"].split()
    b.append(code("$ free -h ; df -h / ; du -sh ~/solemne_so_equipo07",
                  d["free_h"] + "\n\n"
                  + "Filesystem      Size  Used Avail Use% Mounted on\n" + d["df_raiz"]
                  + "\n\n" + d["du_proyecto"] + "\t" + RUTA_PROYECTO_VM))
    b.append(p(
        "La máquina tiene " + mem[1] + " de memoria total, con " + mem[2] + " en uso y "
        + mem[3] + " libres. La columna buff/cache (" + mem[5] + ") no es memoria perdida "
        "sino caché de disco que el núcleo devuelve en cuanto un proceso la necesita, y "
        "por eso la disponible (" + mem[6] + ") supera a la libre; el intercambio queda en "
        "0 B usados, señal de que el trabajo nunca presionó la memoria física. El proyecto "
        "vive en " + dfr[0] + " montado en /, con " + dfr[1] + " de capacidad, " + dfr[2]
        + " usados y " + dfr[3] + " disponibles (" + dfr[4] + " de uso), y ocupa "
        + d["du_proyecto"] + "; ese total supera la suma de sus subcarpetas por el "
        "redondeo de du a bloques de " + d["stat_ioblock"] + " bytes."))

    b.append(h2("stat", "Metadatos de la bitácora con stat"))
    b.append(code("$ stat logs/gestion_ambiental.log ; ls -lah logs/gestion_ambiental.log",
                  d["stat_bitacora"] + "\n" + d["ls_bitacora"]))
    b.append(tabla(
        ["Campo de stat", "Valor", "Lectura"],
        [["Tipo y tamaño",
          "regular file, " + d["stat_size"] + " B en " + d["stat_bloques"] + " bloques",
          "Archivo ordinario de datos. El tamaño lógico es menor que el espacio "
          "reservado, porque el sistema de archivos asigna bloques completos de "
          + d["stat_ioblock"] + " bytes."],
         ["Inodo y enlaces",
          d["stat_inodo"] + " (dispositivo " + d["stat_dispositivo"] + "), "
          + d["stat_enlaces"] + " enlace",
          "El nombre gestion_ambiental.log es sólo una entrada de directorio que apunta a "
          "ese inodo. Un enlace duro subiría el contador a 2 sin duplicar un byte."],
         ["Permisos", d["stat_permisos"],
          "Lectura y escritura para el propietario y su grupo, sólo lectura para el "
          "resto: la bitácora es auditable por terceros pero no modificable."],
         ["Propietario / grupo", d["stat_uid"] + " / " + d["stat_gid"],
          "Pertenece al usuario sin privilegios del equipo, no a root: el gestor no "
          "necesita permisos elevados para operar."],
         ["Access / Modify / Change / Birth", "iguales al segundo",
          "Última lectura, última escritura de contenido, último cambio de metadatos del "
          "inodo y creación: coinciden porque el archivo se creó y se escribió en la "
          "misma corrida."]],
        anchos=[3.0, 3.9, 9.6]))
    b.append(p(
        "Esa bitácora tiene " + d["ev_n_bitacora"] + " líneas y " + d["stat_size"]
        + " bytes. Ambas cifras corresponden a la misma corrida: el generador de este "
        "informe compara el conteo de líneas del archivo entregado con el que declara la "
        "evidencia, y el tamaño del archivo entregado con el que devolvió stat, y se "
        "detiene si alguno de los dos pares no calza."))

    # ---------------------------------------- 7. Mover vs copiar y errores
    b.append(h1("mover", "Mover frente a copiar, y control de errores"))

    b.append(h2("mover_copiar", "Por qué los informes se mueven y el resumen se copia"))
    b.append(p(
        "Los informes se MUEVEN desde salida/ a su carpeta de clasificación porque el "
        "traslado es un cambio de estado definitivo: un informe ya clasificado no debe "
        "seguir figurando como pendiente. Copiarlos duplicaría cada informe y se perdería "
        "la propiedad más útil: que salida/ sin informes significa \"no queda nada por "
        "clasificar\"."))
    b.append(p(
        "Mover dentro de una misma partición (aquí origen y destino están ambos en "
        + dfr[0] + ") es la llamada rename(2): no copia datos, elimina la entrada de "
        "directorio antigua y crea otra que apunta al MISMO inodo. Los bloques no se "
        "tocan, el número de inodo no cambia y sólo se actualiza el ctime; por eso es casi "
        "instantánea y su costo no depende del tamaño. Entre particiones distintas el "
        "núcleo no podría renombrar y shutil.move degradaría a copiar y borrar."))
    b.append(p(
        "El resumen consolidado tiene otro papel: es la vista general que el módulo de "
        "monitoreo consulta en su ruta de siempre. Por eso salida/resumen_ambiental.txt se "
        "CONSERVA y además se COPIA como gestion_ambiental/resumen_resguardado.txt con "
        "shutil.copy. Copiar sí crea un inodo nuevo, con su contenido en otros bloques: "
        "desde ese momento los dos archivos son independientes y modificar uno no altera "
        "al otro, que es lo que se espera de un respaldo."))
    b.append(tabla(
        ["", "Mover (rename)", "Copiar (copy)"],
        [["Se aplica a", "informe_*.txt", "resumen_ambiental.txt"],
         ["Efecto en el inodo", "conserva el mismo inodo", "crea un inodo nuevo"],
         ["Datos en disco", "no se duplican", "se duplican"],
         ["Original", "deja de existir en el origen", "permanece en salida/"],
         ["Propósito", "reclasificación definitiva", "respaldo y auditoría"]],
        anchos=[3.6, 6.5, 6.4]))
    b.append(code("# Comprobacion posterior a la ejecucion del gestor", d["mover_copiar"]))
    b.append(p(
        "El resumen original (" + str(d["resumen_bytes"]) + " bytes) sigue en salida/, la "
        "copia resguardada pesa " + str(d["resguardo_bytes"]) + " bytes y su contenido es "
        "idéntico byte a byte: "
        + ("comprobado" if d["resguardo_identico"] else "NO coincide") + ". En salida/ "
        "quedan " + str(d["n_salida"]) + " informes individuales, porque todos fueron "
        "movidos y no copiados; esto es lo que documenta " + ref("leeme") + "."))

    b.append(h2("errores", "Control de errores sin detener el procesamiento"))
    b.append(nota(
        "Esta demostración es DESTRUCTIVA: inyecta datos dañados. Por eso NO se ejecutó "
        "sobre el árbol entregado sino sobre una COPIA completa en " + d["copia_control"]
        + ". De ahí que la bitácora entregada cierre con \"no se detectaron anomalias en "
        "esta corrida\": el entregable viene de una corrida limpia y las cifras de esta "
        "subsección pertenecen a la copia."))
    b.append(p(
        "El requisito 13 pide controlar al menos un error sin detener el procesamiento. "
        "Se inyectaron cuatro fallas distintas, para probar tanto el análisis de la "
        "bitácora de alertas como la lectura de los informes: una línea sin separadores, "
        "otra con campos insuficientes, un indicador desconocido (\"Radiacion\") y un "
        "informe sin la línea \"Alertas detectadas: N\"."))
    b.append(code("$ python3 src/gestor_incidencias.py ; echo \"Codigo de salida: $?\"\n"
                  "$ grep 'ANOMALIA\\|CONTROL DE ERRORES' logs/gestion_ambiental.log",
                  "Codigo de salida: " + d["gestor_exit_error"] + "\n\n"
                  + "\n".join(d["anomalias_lineas"][:4]) + "\n" + d["control_errores"]))
    b.append(p(
        "El comportamiento es el correcto: las " + d["anomalias_total"] + " anomalías "
        "quedaron registradas con su número de línea y su contenido, el gestor NO se "
        "detuvo, terminó con código " + d["gestor_exit_error"] + ", conservó las "
        + d["inventario_tras_error"] + " alertas válidas y generó el inventario igual. "
        "Descartar el dato defectuoso y continuar es preferible a abortar: una línea "
        "corrupta no puede invalidar una jornada de monitoreo, pero tampoco puede "
        "desaparecer en silencio."))
    # Pie honesto: la captura NO corresponde a la corrida de cuatro fallas que
    # se transcribe arriba, sino a otra de control con dos alertas danadas. Su
    # ultima linea queda cortada por el borde de la consola y NO figura en
    # evidencias/debian/06-control-de-errores.txt, que registra la otra corrida
    # (cuatro anomalias, otra marca de tiempo): remitir alli era mandar al
    # lector a una comprobacion que falla. Lo que la captura si demuestra esta
    # completo en sus propias lineas anteriores, y eso es lo que declara el pie.
    b.append(fig(recortar_consola(FOTOS / "debian-04-stat-y-control-de-errores.png"),
                 "Metadatos de la bitácora con stat y control de errores en la consola "
                 "de la máquina virtual. Arriba, el stat del archivo entregado: los "
                 "mismos " + d["stat_size"] + " bytes e inodo " + d["stat_inodo"]
                 + " que transcribe " + ref("stat") + ". Abajo, una corrida de control "
                 "DISTINTA de la de " + ref("errores") + ", con dos alertas dañadas en "
                 "vez de cuatro fallas; sus líneas legibles ya dan el resultado completo "
                 "(código " + d["gestor_exit_error"] + ", dos anomalías y "
                 + d["inventario_tras_error"] + " alertas conservadas). La última línea, "
                 "el resumen CONTROL DE ERRORES de esa misma corrida, es la que corta el "
                 "borde de la consola: no está transcrita en ninguna evidencia, porque la "
                 "corrida archivada en evidencias/debian/06-control-de-errores.txt es la "
                 "otra.", texto_pt=TEXTO_DE_CAPTURA_PT))

    b.append(h2("anti", "Idempotencia y protección contra sobrescritura"))
    b.append(nota(
        "Esta demostración también es destructiva, porque duplica informes a propósito, y "
        "por eso se ejecutó sobre la misma COPIA " + d["copia_anti"] + ". El árbol "
        "entregado conserva sus " + str(inv["resumen"]["total_informes"]) + " informes "
        "sin sufijos de versión, que es el estado de una única corrida limpia."))
    b.append(code("# Tres corridas consecutivas del gestor sobre la copia\n"
                  "$ ls -lah gestion_ambiental/criticas/informe_STG04_20240101*",
                  d["idempotencia"] + "\n\n" + d["ejemplo_versiones"]))
    b.append(p(
        "Las alertas no se duplican entre corridas porque los registros por indicador se "
        "reconstruyen de forma atómica sobre archivos temporales que se renombran al "
        "final. Al reprocesar, los informes ya clasificados no se pierden ni se pisan: "
        "obtener_nombre_seguro() comprueba la existencia del destino y agrega un sufijo "
        "correlativo. Partiendo de " + d["informes_iniciales_v"] + " informes, tras dos "
        "ciclos adicionales conviven " + d["encontrados_v"] + " donde se esperaban "
        + d["esperados_v"] + " (" + d["con_sufijo_v"] + " con sufijo): ningún archivo fue "
        "sobrescrito ni eliminado."))

    # ------------------------------------------------------ 8. Conclusiones
    b.append(h1("conclusiones", "Conclusiones"))

    b.append(h2("concurrencia",
                "Sobre concurrencia: el resultado que no favorece la hipótesis"))
    b.append(p(
        "La conclusión más importante es también la más incómoda: en este caso concreto "
        "la concurrencia NO aceleró el procesamiento. La versión concurrente tardó "
        + d["t_con"]["real"] + " frente a " + d["t_sec"]["real"] + " de la secuencial, y "
        "su tiempo interno fue " + d["t_con"]["interno"] + " s frente a "
        + d["t_sec"]["interno"] + " s. El dato no se maquilla, se explica:"))
    b.append(ul([
        "El trabajo útil es demasiado pequeño: " + d["archivos_generados"]
        + " archivos con " + res["Lineas leidas"] + " líneas, procesados en centésimas de "
        "segundo.",
        "El costo fijo de la concurrencia no depende del tamaño del trabajo: crear "
        + d["t_con"]["trabajadores"] + " hilos, encolar los archivos y tomar y soltar los "
        "mutex en cada actualización pesa más que lo que se ahorra al repartir un trabajo "
        "tan breve.",
        "El trabajo está dominado por entrada/salida sobre archivos diminutos que el "
        "núcleo ya tiene en caché, de modo que hay poca espera que solapar; y en CPython "
        "el bloqueo global del intérprete impide que dos hilos ejecuten código Python puro "
        "a la vez.",
    ]))
    b.append(p(
        "La concurrencia sí se observa cuando el trabajo crece: con la carga amplificada "
        "de " + d["carga_archivos"] + " archivos el proceso alcanzó " + d["pcpu_max"]
        + " % de CPU, por encima del 100 % que jamás superaría un proceso de un solo hilo. "
        "El mecanismo funciona; lo que falta en el conjunto oficial es trabajo suficiente "
        "para amortizar su costo de entrada."))

    b.append(h2("sincronizacion", "Sobre sincronización y sistema de archivos"))
    b.append(p(
        "Que las " + d["metricas_comparadas"] + " métricas coincidan y que los "
        + d["informes_comparados"] + " informes sean idénticos es la evidencia de que la "
        "sincronización es correcta: con " + d["t_con"]["trabajadores"] + " hilos "
        "actualizando contadores compartidos, sin mutex cada corrida habría dado un "
        "resultado distinto. El patrón fue acumular en variables locales y entrar a la "
        "sección crítica una sola vez por archivo; queue.Queue garantiza que cada archivo "
        "lo toma un único trabajador. El único efecto observable de la concurrencia es el "
        "ORDEN de las líneas de la bitácora."))
    b.append(p(
        "Mover y copiar obligó a distinguir el nombre de un archivo del archivo mismo: un "
        "nombre es una entrada de directorio que apunta a un inodo. Las herramientas "
        "completaron el cuadro: stat dio inodo, enlaces, permisos y marcas de tiempo; "
        "ls -lah y du -sh, la diferencia entre tamaño lógico y bloques ocupados; df -h, el "
        "sistema de archivos y su espacio libre; free -h, que la caché no es memoria "
        "perdida; y ps con /proc, la estructura interna del proceso."))

    b.append(h2("honestidad", "Declaraciones de honestidad técnica"))
    b.append(ul([
        "Hipervisor: el hito declaró " + HIPERVISOR_PLANIFICADO + " y la implementación "
        "final se hizo sobre " + hv["hipervisor_implementado"] + ", con los recursos "
        "comprometidos intactos; la pauta admite otro hipervisor. TODAS las figuras de "
        "instalación son de la máquina real en Hyper-V: de la etapa previa sobre "
        + HIPERVISOR_PLANIFICADO + " no se conserva ninguna captura en el entregable, "
        "porque no correspondía a la máquina entregada.",
        "Rendimiento: la versión concurrente resultó más lenta que la secuencial en el "
        "conjunto oficial; se informa tal como se midió.",
        "Clasificación: sin_alertas/ quedó con " + str(d["n_sin"]) + " informes porque la "
        "pauta exige al menos una alerta por archivo; la rama se demostró aparte.",
        "Demostraciones destructivas: el control de errores (" + ref("errores") + ") y la "
        "anti-sobrescritura (" + ref("anti") + ") se ejecutaron sobre una COPIA, no sobre "
        "el árbol entregado; por eso la bitácora entregada no registra esas anomalías.",
        "Trazabilidad: cada cifra se lee al generar el informe desde evidencias/ y del "
        "estado real del proyecto, y el generador compara ambas fuentes -y las cifras del "
        "README- antes de escribir nada; si discrepan, no se emite el documento.",
    ]))

    b.append(h2("indice", "Índice de evidencias que respaldan este informe"))
    b.append(p(
        "Cada sección puede auditarse contra el archivo que la origina; el generador "
        "falla si alguno cambia de forma incompatible."))
    b.append(tabla(
        ["Archivo de evidencia", "Contenido", "Secciones"],
        [["evidencias/hyperv/configuracion-vm-hyperv.txt",
          "cmdlets de Hyper-V y verificación SHA256 del ISO",
          ref("iso") + ", " + ref("hipervisor")],
         ["evidencias/debian/01-preparar-ambiente.txt",
          "apt update, apt upgrade y entorno de Python", ref("ambiente")],
         ["evidencias/debian/02-entorno.txt",
          "distribución, kernel, hipervisor, CPU, memoria, disco y red",
          ref("instalacion")],
         ["evidencias/debian/03-ejecucion-parte1.txt",
          "ambas corridas, las seis métricas y la comparación de informes y alertas",
          ref("parte1")],
         ["evidencias/debian/04-observacion-procesos.txt",
          "/usr/bin/time -v, ps, ps -L, /proc, free -h, df -h y du -sh", ref("procesos")],
         ["evidencias/debian/05-gestor-incidencias.txt",
          "estructura, inventario, stat de la bitácora, mover frente a copiar",
          ref("gestor") + ", " + ref("mover_copiar")],
         ["evidencias/debian/06-control-de-errores.txt",
          "anomalías inyectadas y rama sin_alertas (sobre una copia)",
          ref("sin_alertas") + ", " + ref("errores")],
         ["evidencias/debian/07-anti-sobrescritura.txt",
          "sufijos _v1 y _v2 e idempotencia (sobre una copia)", ref("anti")],
         ["evidencias/salida_parte1/",
          d["respaldo_parte1"] + " archivos: salida/ tal como la dejó la Parte 1",
          ref("leeme")],
         ["evidencias/fotos/",
          "instalación real en Hyper-V y consola de la máquina definitiva",
          "Figuras 1 a %d" % n_figuras()],
         ["gestion_ambiental/, alertas/, logs/, salida/",
          "estado real: informes, alertas por indicador, inventario y bitácora",
          ref("gestor") + ", " + ref("mover")],
         ["docs/Solemne01PracticaParte2FormaB.pdf",
          "pauta oficial contra la cual se verificó cada requisito", "todas"]],
        anchos=[6.4, 7.2, 2.9]))

    return b

# ---------------------------------------------------------------------------
# Render DOCX
# ---------------------------------------------------------------------------

def render_docx(bloques, destino):
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
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

        # Mismo pie que el PDF: linea de separacion, identificacion del
        # trabajo pegada al margen izquierdo y numero de pagina pegado al
        # derecho. El DOCX ponia antes "Página N de 12" centrado y sin linea,
        # de modo que los dos archivos de la misma entrega no se parecian
        # entre si; el criterio elegido es el del PDF.
        GRIS_PIE = RGBColor(0x6B, 0x72, 0x80)

        # La linea va en un parrafo propio, minimo, y no como borde superior
        # del parrafo del texto. Puestos en el mismo parrafo, Word deja de
        # aplicar el tabulador derecho y el numero de pagina se pega al texto
        # de la izquierda; comprobado convirtiendo el DOCX con el propio Word.
        raya = section.footer.paragraphs[0]
        borde = OxmlElement("w:pBdr")
        arriba = OxmlElement("w:top")
        arriba.set(qn("w:val"), "single")
        arriba.set(qn("w:sz"), "4")
        arriba.set(qn("w:space"), "0")
        arriba.set(qn("w:color"), "D1D5DB")
        borde.append(arriba)
        raya._p.get_or_add_pPr().append(borde)
        raya.paragraph_format.space_before = Pt(0)
        raya.paragraph_format.space_after = Pt(0)
        raya.paragraph_format.line_spacing = Pt(5)
        raya.add_run("").font.size = Pt(1)

        pie = section.footer.add_paragraph()
        pie.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pie.paragraph_format.space_before = Pt(0)
        pie.paragraph_format.space_after = Pt(0)
        # El tabulador derecho va al final de la caja de texto (21 - 2 - 2 cm),
        # que es donde el PDF alinea su "Pagina N".
        pie.paragraph_format.tab_stops.add_tab_stop(Cm(17.0), WD_TAB_ALIGNMENT.RIGHT)
        izq = pie.add_run(limpiar(EVALUACION + "  |  " + EQUIPO + "  |  " + SECCION)
                          + "	" + limpiar("Página "))
        izq.font.size = Pt(7.6)
        izq.font.color.rgb = GRIS_PIE
        num = campo(pie, " PAGE ")
        num.font.size = Pt(7.6)
        num.font.color.rgb = GRIS_PIE

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
    # 1.04 y no 1.06: con 1.06 el documento se iba tres centimetros a una
    # decimotercera hoja que quedaba practicamente vacia. Calibri 10 con
    # interlineado sencillo ya deja 12.2 pt de linea, de modo que el texto
    # sigue teniendo el mismo aire que el del PDF.
    normal.paragraph_format.line_spacing = 1.04
    # El equivalente en Word de allowWidows/allowOrphans del PDF: ninguna linea
    # suelta de un parrafo se queda al pie ni abre la pagina siguiente.
    normal.paragraph_format.widow_control = True

    def justificable(texto):
        """Texto de prosa listo para justificar sin abrir rios.

        Word no parte por si solo scripts/generar_archivos_entrada.py ni
        estacion_CODIGO_AAAAMMDD.jsonl, asi que reparte todo el sobrante de la
        linea entre los espacios y abre el rio. El espacio de ancho cero
        U+200B es el caracter con el que se le dice "aqui puedes partir": no
        ocupa nada, no se imprime y no cambia ni una letra del texto.

        Va DESPUES de limpiar(), que sustituye por "?" todo lo que pase de
        U+00FF y se llevaria por delante la marca.
        """
        return con_puntos_de_corte(limpiar(texto))

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

    def pegar(par):
        """Ata un parrafo al siguiente: el equivalente de keepWithNext.

        Es lo que impide que un titulo -o un parrafo que anuncia con dos
        puntos lo que viene detras- se quede solo al pie de una hoja, y lo que
        mantiene una figura con su pie y un comando con su salida.
        """
        par.paragraph_format.keep_with_next = True
        return par

    def respirar(puntos=5):
        """Separacion despues de un bloque o de una tabla.

        Un parrafo vacio del estilo normal mide una linea entera -unos 14 pt-
        aunque no lleve nada escrito, y el documento gastaba asi mas de diez
        centimetros repartidos entre los 19 bloques de transcripcion y las 8
        tablas. Con una fuente diminuta el mismo parrafo mide lo que se le pida
        y la separacion queda igual a la del PDF.
        """
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
        par.paragraph_format.line_spacing = 1.0
        par.add_run("").font.size = Pt(puntos)
        return par

    def entera(tab):
        """Cabecera repetida en cada pagina y filas que no se parten.

        Una tabla que cruza una pagina se sigue leyendo si la cabecera vuelve a
        salir arriba; lo que no se lee es una fila cortada por la mitad. Y una
        tabla corta no debe cruzar nada: se ata fila con fila para que viaje
        entera, que es lo que hace el PDF con las que caben en 4 cm.
        """
        filas = tab.rows
        corta = len(filas) <= FILAS_TABLA_JUNTAS
        for i, fila in enumerate(filas):
            pr = fila._tr.get_or_add_trPr()
            pr.append(OxmlElement("w:cantSplit"))
            if i == 0:
                pr.append(OxmlElement("w:tblHeader"))
            if corta and i < len(filas) - 1:
                for celda in fila.cells:
                    pegar(celda.paragraphs[-1])
        return tab

    for blk in bloques:
        t = blk["t"]

        if t == "portada":
            # El bloque se centra en la hoja igual que en el PDF. Los tres
            # parrafos vacios de antes lo dejaban terminando cerca de la mitad
            # de la pagina, con media hoja en blanco debajo, y ademas no se
            # parecia a la portada del PDF.
            #
            # Word no deja medir un parrafo, asi que el alto del bloque se
            # calcula: cada linea mide su cuerpo por 1.2 -la proporcion con la
            # que Word compone una linea sencilla- por el interlineado 1.04 de
            # este documento, mas la separacion que lleve debajo.
            hueco = doc.add_paragraph()
            alto = [0.0]

            def centrado(texto, size, negrita=False, color=None, espacio=6):
                par = doc.add_paragraph()
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                par.paragraph_format.space_after = Pt(espacio)
                run = par.add_run(limpiar(texto))
                run.font.size = Pt(size)
                run.bold = negrita
                if color:
                    run.font.color.rgb = RGBColor(*color)
                alto[0] += size * 1.2 * 1.04 + espacio
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
            separador = doc.add_paragraph()
            separador.paragraph_format.space_after = Pt(0)
            separador.add_run("").font.size = Pt(10)
            alto[0] += 10 * 1.2 * 1.04
            centrado(FECHA, 11, espacio=0)
            # El parrafo de arriba se estira hasta la parte del hueco que le
            # toca. line_spacing en puntos fija el alto exacto de la linea.
            hueco.paragraph_format.space_before = Pt(0)
            hueco.paragraph_format.space_after = Pt(0)
            hueco.add_run("").font.size = Pt(1)
            hueco.paragraph_format.line_spacing = Pt(
                max(1.0, (ALTO_DE_TEXTO_PT - alto[0]) * ALTO_PORTADA))

        elif t == "salto":
            par = doc.add_paragraph()
            par.add_run().add_break(WD_BREAK.PAGE)

        elif t == "h1":
            # Mismo criterio que el PDF: el titulo pertenece a lo que viene
            # DEBAJO, asi que lleva mucho aire encima y casi ninguno debajo.
            # Word SUMA los dos espacios de parrafos contiguos (a diferencia
            # de reportlab, que se queda con el mayor), de modo que encima de
            # un h1 quedan estos 12 pt mas los 5 del parrafo anterior.
            par = pegar(doc.add_paragraph())
            par.paragraph_format.space_before = Pt(12)
            par.paragraph_format.space_after = Pt(1.5)
            run = par.add_run(limpiar(blk["x"]))
            run.bold = True
            run.font.size = Pt(15)
            run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

        elif t == "h2":
            par = pegar(doc.add_paragraph())
            par.paragraph_format.space_before = Pt(9)
            par.paragraph_format.space_after = Pt(1)
            run = par.add_run(limpiar(blk["x"]))
            run.bold = True
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x2E, 0x5A, 0x88)

        elif t == "p":
            par = doc.add_paragraph(justificable(blk["x"]))
            par.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if blk["x"].rstrip().endswith(":"):
                pegar(par)

        elif t in ("ul", "ol"):
            estilo = "List Bullet" if t == "ul" else "List Number"
            for item in blk["x"]:
                par = doc.add_paragraph(justificable(item), style=estilo)
                par.alignment = WD_ALIGN_PARAGRAPH.LEFT
                par.paragraph_format.space_after = Pt(3)

        elif t == "code":
            # Una linea del comando por parrafo, igual que en el PDF: metidas
            # todas en un solo run, Word convierte los saltos en espacios y un
            # comando de cinco lineas salia escrito de corrido.
            cmd = limpiar(blk["cmd"]).splitlines() if blk["cmd"] else []
            salida = limpiar(blk["x"]).splitlines() or [""]
            # Mismo criterio que el PDF: hasta LINEAS_CODIGO_JUNTAS el bloque
            # no se parte, y si es mas largo se garantiza que el comando se
            # lleve consigo al menos LINEAS_CODIGO_MINIMAS lineas de salida.
            total = len(cmd) + len(salida)
            if total <= LINEAS_CODIGO_JUNTAS:
                juntas = total
            else:
                juntas = len(cmd) + LINEAS_CODIGO_MINIMAS
            escritas = 0
            for linea in cmd:
                par = doc.add_paragraph()
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.space_before = Pt(4) if not escritas else Pt(0)
                par.paragraph_format.line_spacing = 1.0
                sombrear(par, "E8EAED")
                run = par.add_run(linea)
                run.font.name = "Consolas"
                run.font.size = Pt(7.5)
                run.bold = True
                escritas += 1
                if escritas < juntas:
                    pegar(par)
            for linea in salida:
                par = doc.add_paragraph()
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.line_spacing = 1.0
                sombrear(par, "F4F5F7")
                run = par.add_run(linea)
                run.font.name = "Consolas"
                run.font.size = Pt(7.5)
                escritas += 1
                if escritas < juntas:
                    pegar(par)
            respirar()

        elif t == "nota":
            par = doc.add_paragraph()
            par.alignment = WD_ALIGN_PARAGRAPH.LEFT
            par.paragraph_format.left_indent = Cm(0.3)
            par.paragraph_format.space_before = Pt(6)
            par.paragraph_format.space_after = Pt(8)
            sombrear(par, "FBF3E2")
            borde_izquierdo(par, "B7791F")
            run = par.add_run(justificable(blk["x"]))
            run.font.size = Pt(9.5)

        elif t == "tabla":
            tab = doc.add_table(rows=1, cols=len(blk["head"]))
            tab.style = "Table Grid"
            hdr = tab.rows[0].cells
            for i, texto in enumerate(blk["head"]):
                hdr[i].text = ""
                par = hdr[i].paragraphs[0]
                par.paragraph_format.space_after = Pt(1)
                run = par.add_run(limpiar(texto))
                run.bold = True
                run.font.size = Pt(8)
            for fila in blk["rows"]:
                celdas = tab.add_row().cells
                for i, texto in enumerate(fila):
                    celdas[i].text = ""
                    par = celdas[i].paragraphs[0]
                    par.paragraph_format.space_after = Pt(1)
                    run = par.add_run(limpiar(str(texto)))
                    run.font.size = Pt(7.5)
            if blk.get("w"):
                total = sum(blk["w"])
                disponible = 17.0
                for i, ancho in enumerate(blk["w"]):
                    cm = Cm(ancho / total * disponible)
                    for fila in tab.rows:
                        fila.cells[i].width = cm
            entera(tab)
            respirar()

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
            # El pie cierra su figura: pegado a la imagen por arriba (0 pt) y
            # separado del bloque siguiente por abajo (10 pt). Sumando los 5 pt
            # del parrafo anterior, encima de la figura quedan 9 pt: menos que
            # los 10 que quedan debajo del pie, que es lo que agrupa el
            # conjunto figura+pie y lo separa de lo que venga despues.
            par = pegar(doc.add_paragraph())
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.space_before = Pt(4)
            par.paragraph_format.space_after = Pt(0)
            par.add_run().add_picture(ruta, width=Cm(ancho))
            pie = doc.add_paragraph()
            pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pie.paragraph_format.space_before = Pt(0)
            pie.paragraph_format.space_after = Pt(10)
            run = pie.add_run(limpiar(blk["pie"]))
            run.italic = True
            run.font.size = Pt(7.5)

        else:
            morir("bloque desconocido en el modelo: %r" % t)

    # Igual que los Spacer colgantes del PDF: un parrafo vacio detras del
    # ultimo bloque basta para que Word abra una hoja mas y la deje con el pie
    # de pagina y nada mas. Se quitan todos los que queden al final.
    while doc.paragraphs:
        ultimo = doc.paragraphs[-1]
        if ultimo.text.strip() or ultimo._p.findall(qn("w:r") + "/" + qn("w:drawing")):
            break
        ultimo._p.getparent().remove(ultimo._p)

    # Word EXIGE un parrafo detras de una tabla final y lo pone el solo si no
    # esta: con el estilo normal mide 17 pt y se llevaba una hoja entera para
    # el solo. Se pone aqui uno propio, del tamano minimo, que cabe en lo que
    # queda de la ultima pagina.
    if len(doc.element.body) and doc.element.body[-2].tag == qn("w:tbl"):
        cierre = respirar(1)
        cierre.paragraph_format.line_spacing = Pt(2)

    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))


# ---------------------------------------------------------------------------
# Render PDF
# ---------------------------------------------------------------------------

def render_pdf(bloques, destino):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (BaseDocTemplate, CondPageBreak, Frame, Image,
                                    KeepTogether, PageBreak, PageTemplate,
                                    Paragraph, Spacer, Table, TableStyle)
    from xml.sax.saxutils import escape

    AZUL = colors.HexColor("#1F3A5F")
    AZUL2 = colors.HexColor("#2E5A88")
    GRIS = colors.HexColor("#F4F5F7")
    GRIS2 = colors.HexColor("#E8EAED")
    CREMA = colors.HexColor("#FBF3E2")
    AMBAR = colors.HexColor("#B7791F")

    base = getSampleStyleSheet()
    S = {}
    # allowWidows=0 impide que la ultima linea de un parrafo se quede sola al
    # principio de la pagina siguiente; allowOrphans=0 (el valor por omision de
    # reportlab) impide lo simetrico, una primera linea sola al pie.
    #
    # uriWasteReduce es el interruptor con el que reportlab permite partir una
    # palabra "de tipo direccion" cuando no cabe entera al final de la linea.
    # El numero es cuanta linea esta dispuesto a desperdiciar ANTES de partir:
    # 0.05 significa "si queda mas de un 5 % de linea libre, parte en vez de
    # estirar los espacios", que es justo lo que mata los rios. Va solo en los
    # estilos de prosa justificada; los bloques de transcripcion y las tablas
    # no lo llevan y siguen exactamente igual que antes.
    rutas_partibles_en_pdf()
    # ALINEACION: bandera a la izquierda, no justificado a ambos lados.
    # Este informe cita rutas largas e indivisibles (gestion_ambiental/,
    # salida/resumen_ambiental.txt, alertas_por_indicador/). Justificando, el
    # motor solo puede repartir el sobrante en los espacios entre palabras y
    # aparecen "rios": lineas con blancos de dos a diez veces lo normal, que se
    # ven antes de leerlas. Se probaron puntos de corte invisibles dentro de las
    # rutas y bajaron el efecto, pero no lo eliminaron. Con bandera a la
    # izquierda el espacio entre palabras es constante y el sobrante se va al
    # margen derecho, que es donde no molesta. Es ademas lo habitual en
    # documentacion tecnica con rutas y codigo.
    S["p"] = ParagraphStyle("p", parent=base["BodyText"], fontName="Helvetica",
                            fontSize=9.0, leading=10.9, alignment=TA_LEFT,
                            spaceAfter=3.5, allowWidows=0, allowOrphans=0,
                            uriWasteReduce=0.05)
    # keepWithNext deja constancia de la intencion en el propio estilo, pero no
    # basta: reportlab solo lo encadena con UN flowable y ademas se rinde si el
    # siguiente es un Spacer, que es justo lo que abre un bloque de codigo, una
    # tabla o una nota. El agrupado real lo hace agrupar_unidades() mas abajo.
    # Un titulo pertenece a lo que viene DESPUES, y el espacio en blanco es lo
    # unico que se lo dice al lector. Antes el reparto estaba casi al reves:
    # el h1 tenia 6 pt encima y 5 debajo, y el h2 6 y 2.5, de modo que "5.2
    # Inventario y cuadratura de las alertas" y "6.5 Memoria del sistema..."
    # -los dos vienen justo detras del pie de una figura- se agrupaban con la
    # figura de arriba en vez de con el texto que anuncian. Ahora el
    # spaceBefore es varias veces el spaceAfter en los dos niveles.
    #
    # reportlab NO suma los dos espacios: el hueco real entre dos flowables es
    # max(spaceAfter del anterior, spaceBefore del siguiente). Por eso subir el
    # spaceBefore de un titulo solo se nota frente a lo que tenga menos
    # spaceAfter que el; frente al pie de una figura, que deja 10, el hueco de
    # encima lo pone el pie. Con 10 y 7 el hueco de encima es de 10 pt en todos
    # los casos, contra 1.5 y 0.5 pt por debajo.
    S["h1"] = ParagraphStyle("h1", parent=S["p"], fontName="Helvetica-Bold",
                             fontSize=13, leading=15.5, textColor=AZUL,
                             spaceBefore=10, spaceAfter=1.5, alignment=0,
                             keepWithNext=1)
    S["h2"] = ParagraphStyle("h2", parent=S["p"], fontName="Helvetica-Bold",
                             fontSize=10.2, leading=12.4, textColor=AZUL2,
                             spaceBefore=7, spaceAfter=0.5, alignment=0,
                             keepWithNext=1)
    S["li"] = ParagraphStyle("li", parent=S["p"], leftIndent=0.55 * cm,
                             bulletIndent=0.15 * cm, spaceAfter=2.5)
    # uriWasteReduce=0: una transcripcion se parte donde se partia antes. El
    # informe promete que estos bloques son la salida literal de la maquina, y
    # cambiar por donde se doblan las lineas largas no es cosa del acabado.
    # Lo mismo vale para las celdas de tabla y para los pies de figura, que no
    # van justificados y por tanto no tienen rios que arreglar.
    S["code"] = ParagraphStyle("code", parent=S["p"], fontName="Courier",
                               fontSize=6.6, leading=7.3, alignment=0,
                               spaceAfter=0, spaceBefore=0, uriWasteReduce=0)
    S["cmd"] = ParagraphStyle("cmd", parent=S["code"], fontName="Courier-Bold")
    S["nota"] = ParagraphStyle("nota", parent=S["p"], fontSize=8.8, leading=10.9,
                               leftIndent=0.25 * cm, rightIndent=0.15 * cm)
    # El pie CIERRA su figura, no encabeza la siguiente. Con 2 pt encima y 5
    # debajo, mas los 2 del espaciador que abre la figura siguiente, quedaban
    # 7 pt entre un pie y la figura de abajo contra 2 pt entre la figura y su
    # propio pie: en la pagina de las tres capturas de la instalacion cada pie
    # se leia como el titulo de la figura siguiente. Ahora el reparto es 0.5 pt
    # hacia arriba -el pie queda pegado a su imagen- contra 10 pt hacia abajo,
    # que con el espaciador de la figura siguiente suman 10.5.
    S["pie"] = ParagraphStyle("pie", parent=S["p"], fontName="Helvetica-Oblique",
                              fontSize=7.4, leading=8.7, alignment=TA_CENTER,
                              spaceBefore=0.5, spaceAfter=10, uriWasteReduce=0)
    S["td"] = ParagraphStyle("td", parent=S["p"], fontSize=7.5, leading=8.6,
                             spaceAfter=0, alignment=0, uriWasteReduce=0)
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
    alto_util = A4[1] - 1.5 * cm - 1.6 * cm

    # -----------------------------------------------------------------------
    # Reglas de maquetacion
    #
    # ARRASTRE_TITULO: cuanto contenido tiene que caber DEBAJO de un titulo
    # para que el titulo se quede en esa pagina. 4.6 cm son unas diez lineas de
    # cuerpo: suficiente para que el lector vea de que trata la seccion antes
    # de pasar de hoja. Es tambien el hueco maximo que puede quedar al pie
    # cuando el titulo se empuja a la pagina siguiente.
    #
    # Las cifras de los bloques de transcripcion y de las tablas viven a nivel
    # de modulo (LINEAS_CODIGO_JUNTAS y companeras), porque el DOCX aplica
    # exactamente las mismas.
    # -----------------------------------------------------------------------
    ARRASTRE_TITULO = 4.0 * cm

    def esc(texto):
        return escape(limpiar(texto))

    historia = []
    # Cada bloque del modelo produce una "unidad": la lista de flowables que le
    # corresponde mas la informacion que hace falta para colocarla sin romperla
    # (si es indivisible y cuanto mide). Trabajar por unidades y no flowable a
    # flowable es lo que permite decidir que un titulo viaja con lo que anuncia.
    unidades = []

    def alto_de(f):
        """Alto que ocupa un flowable, espaciado propio incluido."""
        try:
            _, alto = f.wrap(ancho_util, alto_util)
            return alto + f.getSpaceBefore() + f.getSpaceAfter()
        except Exception:
            return 0.0


    def pie_pagina(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.6)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        if doc_.page > 1:
            canvas.drawString(2.0 * cm, 0.95 * cm,
                              limpiar(EVALUACION + "  |  " + EQUIPO + "  |  " + SECCION))
            canvas.drawRightString(A4[0] - 2.0 * cm, 0.95 * cm,
                                   "Página %d" % doc_.page)
            canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
            canvas.line(2.0 * cm, 1.3 * cm, A4[0] - 2.0 * cm, 1.3 * cm)
        canvas.restoreState()

    for blk in bloques:
        t = blk["t"]
        marca = len(historia)

        if t == "portada":
            # El bloque de la portada se centra en la hoja en vez de colgar de
            # un espaciador fijo. Con los 3.6 cm de antes empezaba a un quinto
            # de la altura y terminaba a dos tercios, dejando todo el tercio
            # inferior vacio; la pagina se leia caida hacia arriba.
            portada = [
                Paragraph(esc(UNIVERSIDAD), S["portada_u"]),
                Paragraph(esc(ASIGNATURA), S["portada_a"]),
                Paragraph(esc(EVALUACION), S["portada_t"]),
                Paragraph(esc(SUBTITULO), S["portada_s"]),
                Paragraph(esc(EQUIPO + "  -  " + SECCION), S["portada_e"]),
                Paragraph(esc(LENGUAJE), S["portada_n"]),
                Spacer(1, 0.8 * cm),
                Paragraph("<b>" + esc("Integrantes") + "</b>", S["portada_n"]),
            ]
            for nombre in INTEGRANTES:
                portada.append(Paragraph(esc(nombre), S["portada_n"]))
            portada.append(Spacer(1, 1.0 * cm))
            portada.append(Paragraph(esc(FECHA), S["portada_n"]))
            # ALTO_PORTADA reparte el hueco sobrante entre arriba y abajo. No
            # es 0.5 sino algo menos, porque el centro optico de una hoja queda
            # por encima del geometrico: un bloque centrado al milimetro se ve
            # bajo. Es la misma proporcion que usa el DOCX.
            historia.append(Spacer(1, max(0.0, (alto_util - sum(alto_de(f) for f in portada))
                                          * ALTO_PORTADA)))
            historia.extend(portada)

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
            # Cuantas filas del bloque son el comando: la segunda pasada las
            # necesita para no dejar nunca un comando sin nada de su salida.
            tab._lineas_cmd = n_cmd
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
                ("TOPPADDING", (0, 0), (-1, -1), 2.4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
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
            historia.append(Spacer(1, 0.5))
            historia.append(img)
            historia.append(Paragraph(esc(blk["pie"]), S["pie"]))

        else:
            morir("bloque desconocido en el modelo: %r" % t)

        flows = historia[marca:]
        del historia[marca:]
        if flows:
            u = {"t": t, "f": flows, "alto": sum(alto_de(f) for f in flows),
                 "lineas": 0,
                 # Un parrafo que termina en dos puntos anuncia lo que viene
                 # detras igual que lo hace un titulo -"la salida es la del
                 # anfitrion y esta archivada en:"- y se maqueta igual, para
                 # que no quede solo al pie con su bloque en la hoja siguiente.
                 "anuncia": t == "p" and blk["x"].rstrip().endswith(":")}
            tab = next((f for f in flows if isinstance(f, Table)), None)
            if tab is not None:
                # alto_de() ya llamo a wrap(), asi que las alturas de fila
                # reales estan calculadas y se puede medir "cabecera mas tres
                # filas" o "comando mas seis lineas" en puntos de verdad.
                alturas = list(getattr(tab, "_rowHeights", None) or [])
                u["lineas"] = len(tab._cellvalues)
                if t == "code":
                    n_min = getattr(tab, "_lineas_cmd", 0) + LINEAS_CODIGO_MINIMAS
                else:
                    n_min = 1 + FILAS_TABLA_MINIMAS
                u["minimo"] = sum(alturas[:n_min]) + 12
            # "parte" tiene que decir exactamente lo que la colocacion va a
            # hacer con la unidad, no lo que en abstracto podria hacerse con
            # ella: si aqui se diera por partible un bloque que luego se coloca
            # entero, un titulo se quedaria al pie confiando en que el bloque
            # empezaria debajo, y el bloque se iria completo a la hoja
            # siguiente dejando el titulo solo, que es el defecto que se
            # queria evitar.
            if t == "code":
                u["parte"] = u["lineas"] > LINEAS_CODIGO_JUNTAS
            elif t == "tabla":
                u["parte"] = (u["alto"] > ARRASTRE_TITULO
                              and u["lineas"] > FILAS_TABLA_MINIMAS)
            else:
                u["parte"] = False
            unidades.append(u)

    # -----------------------------------------------------------------------
    # Segunda pasada: colocacion
    # -----------------------------------------------------------------------
    # Aqui se decide que puede partirse entre paginas y que no. Sin esto un
    # titulo de seccion podia quedar solo al pie de una hoja con su contenido
    # en la siguiente, y un comando podia quedar separado de su salida.
    def agrupar_unidades(unidades):
        salida = []
        i = 0
        n = len(unidades)
        while i < n:
            u = unidades[i]

            if u["t"] not in ("h1", "h2") and not u.get("anuncia"):
                salida.append(u)
                i += 1
                continue

            # Un titulo -o un parrafo que termina en dos puntos- arrastra
            # consigo lo que anuncia: primero los titulos que le sigan (un h1
            # seguido de su h2) y despues contenido de verdad, hasta llenar
            # ARRASTRE_TITULO. Se toman unidades enteras mientras quepan; una
            # unidad que no quepa y que sepa partirse -una tabla larga, una
            # transcripcion larga- se deja fuera, porque empujarla entera
            # dejaria mas hueco del que se quiere evitar.
            grupo = list(u["f"])
            usado = u["alto"]
            j = i + 1
            while j < n and unidades[j]["t"] in ("h1", "h2"):
                grupo += unidades[j]["f"]
                usado += unidades[j]["alto"]
                j += 1
            arrastradas = 0
            while j < n:
                sig = unidades[j]
                if sig["t"] in ("salto", "portada"):
                    break
                cabe = usado + sig["alto"] <= ARRASTRE_TITULO
                # Aunque no quepa, una unidad indivisible (una figura con su
                # pie, una nota) se arrastra igualmente si es la primera: dejar
                # el titulo solo al pie es peor que el hueco que pueda quedar.
                if not cabe and (arrastradas or sig["parte"]):
                    break
                grupo += sig["f"]
                usado += sig["alto"]
                arrastradas += 1
                j += 1
                if not cabe:
                    break
            if arrastradas == 0:
                # Lo que sigue es una tabla o una transcripcion larga, que sabe
                # partirse: basta con exigir que el titulo tenga sitio para
                # empezarla debajo.
                salida.append({"t": u["t"], "cond": ARRASTRE_TITULO,
                               "f": grupo, "junto": False})
            else:
                salida.append({"t": u["t"], "f": grupo, "junto": True})
            i = j
        return salida

    for u in agrupar_unidades(unidades):
        if u.get("cond"):
            historia.append(CondPageBreak(u["cond"]))
        if u.get("junto") or (u["t"] == "fig"):
            historia.append(KeepTogether(u["f"]))
        elif u["t"] == "code":
            # Comando y salida viajan juntos mientras el bloque quepa en un
            # cuarto de pagina; los mas largos se parten, pero nunca dejando
            # menos de LINEAS_CODIGO_MINIMAS lineas con el comando.
            if u["lineas"] <= LINEAS_CODIGO_JUNTAS:
                historia.append(KeepTogether(u["f"]))
            else:
                historia.append(CondPageBreak(u["minimo"]))
                historia.extend(u["f"])
        elif u["t"] == "tabla":
            if u["alto"] <= ARRASTRE_TITULO:
                historia.append(KeepTogether(u["f"]))
            else:
                historia.append(CondPageBreak(u["minimo"]))
                historia.extend(u["f"])
        else:
            historia.extend(u["f"])

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
        print("  Fig %-2d  %-42s %5dx%-5d px -> %5.1f x %-5.1f pt   %3d dpi%s"
              % (f["n"], f["ruta"], f["nativo"][0], f["nativo"][1],
                 f["colocada"][0], f["colocada"][1],
                 round(72.0 / f["escala"]), marca))
    if bajas:
        print("  ATENCION: figuras con texto por debajo de %.1f pt: %s"
              % (TEXTO_MINIMO_PT, ", ".join(str(n) for n in bajas)))
    else:
        print("  Ninguna figura quedo con su texto por debajo de %.1f pt."
              % TEXTO_MINIMO_PT)


if __name__ == "__main__":
    main()
