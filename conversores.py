# -*- coding: utf-8 -*-
"""Motor del Convertidor Markdown (sin interfaz).

Dos sentidos:
  * AMarkdown:      PDF, Word, PowerPoint, Excel, HTML... -> Markdown (MarkItDown)
  * DesdeMarkdown:  Markdown -> Word (.docx), PDF, HTML o texto sin formato

Cada sentido es un "modo" con la misma forma: `preparar()` carga lo pesado,
`leer(origen)` interpreta el archivo una sola vez y `escribir(...)` genera
cada formato de salida pedido.
"""
import base64
import mimetypes
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from urllib.parse import unquote, urlparse


class ErrorAlGuardar(Exception):
    """El archivo se convirtió, pero Windows no dejó escribir el resultado."""


class SinTexto(Exception):
    """El archivo no contiene texto extraíble (p. ej. un PDF escaneado)."""


class ErrorPdf(Exception):
    """No se pudo generar el PDF."""


# --------------------------------------------------------------------------
# Utilidades comunes
# --------------------------------------------------------------------------

def ruta_de_salida(origen, carpeta, extension, sobrescribir, usadas):
    """Devuelve una ruta de salida que no pisa nada indebido.

    Nunca devuelve el propio archivo de origen ni una ruta ya usada en
    esta tanda. Si `sobrescribir` es falso, tampoco una que ya exista.
    """
    def libre(ruta):
        clave = str(ruta).lower()
        if clave in usadas or clave == str(origen).lower():
            return False
        return sobrescribir or not ruta.exists()

    candidata = carpeta / (origen.stem + extension)
    numero = 2
    while not libre(candidata):
        candidata = carpeta / f"{origen.stem} ({numero}){extension}"
        numero += 1
    usadas.add(str(candidata).lower())
    return candidata


def leer_texto(ruta):
    """Lee un archivo de texto tolerando las codificaciones habituales."""
    datos = Path(ruta).read_bytes()
    if datos[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return datos.decode("utf-16")
    try:
        return datos.decode("utf-8-sig")
    except UnicodeDecodeError:
        return datos.decode("cp1252", errors="replace")


def explicar_error(error):
    """Mensaje corto y entendible para la columna de estado."""
    nombre = type(error).__name__
    texto = str(error).strip().splitlines()[0] if str(error).strip() else ""
    if isinstance(error, ErrorAlGuardar):
        return "no se pudo guardar (archivo abierto o carpeta protegida)"
    if isinstance(error, SinTexto):
        return "sin texto (¿documento escaneado?)"
    if isinstance(error, ErrorPdf):
        return texto or "no se pudo crear el PDF"
    if isinstance(error, FileNotFoundError):
        return "el archivo original ya no existe"
    if "UnsupportedFormat" in nombre:
        return "formato no compatible"
    if "FileConversion" in nombre:
        return "no se pudo leer el archivo"
    if "MissingDependency" in nombre:
        return "falta un componente para este formato"
    return (texto[:90] or nombre)


# --------------------------------------------------------------------------
# Sentido 1: cualquier documento -> Markdown
# --------------------------------------------------------------------------

class AMarkdown:
    clave = "a_md"
    pestana = "Archivos  →  Markdown"
    descripcion = ("Carga PDF, Word, PowerPoint, Excel, HTML y más. "
                   "Se crea un documento Markdown por cada archivo.")
    filtro = ("Documentos compatibles", (
        ".pdf", ".docx", ".pptx", ".xlsx", ".xls", ".html", ".htm", ".csv",
        ".json", ".xml", ".txt", ".epub", ".msg", ".ipynb", ".zip",
        ".jpg", ".jpeg", ".png", ".mp3", ".wav",
    ))
    formatos = ((".md", "Markdown (.md)"), (".txt", "Texto (.txt)"))
    por_defecto = (".md",)
    boton = "Convertir a Markdown"

    def __init__(self):
        self._conversor = None

    def preparar(self):
        if self._conversor is None:
            from markitdown import MarkItDown
            self._conversor = MarkItDown(enable_plugins=False)

    def leer(self, origen):
        resultado = self._conversor.convert(str(origen))
        texto = getattr(resultado, "markdown", None) or getattr(resultado, "text_content", "") or ""
        if not texto.strip():
            raise SinTexto()
        return texto

    def escribir(self, texto, origen, destino, extension):
        destino.write_text(texto, encoding="utf-8")


# --------------------------------------------------------------------------
# Sentido 2: Markdown -> Word / PDF / HTML / texto
# --------------------------------------------------------------------------

CSS = """
body{margin:0;background:#fff;color:#1f2328;font-family:"Segoe UI",system-ui,-apple-system,Helvetica,Arial,sans-serif;font-size:16px;line-height:1.6}
main{max-width:820px;margin:0 auto;padding:40px 28px 64px}
h1,h2,h3,h4,h5,h6{line-height:1.25;margin:1.6em 0 .6em;font-weight:600}
h1{font-size:2em;border-bottom:1px solid #d8dee4;padding-bottom:.3em}
h2{font-size:1.5em;border-bottom:1px solid #d8dee4;padding-bottom:.3em}
h3{font-size:1.25em}
main>:first-child{margin-top:0}
p,ul,ol,table,pre,blockquote{margin:0 0 1em}
li>ul,li>ol,li>p{margin-bottom:0}
a{color:#0969da}
code{font-family:Consolas,"Cascadia Mono",monospace;font-size:.9em;background:#f3f4f6;padding:.15em .35em;border-radius:4px}
pre{background:#f6f8fa;padding:14px 16px;border-radius:6px;white-space:pre-wrap;word-wrap:break-word}
pre code{background:none;padding:0}
blockquote{border-left:4px solid #d0d7de;color:#59636e;padding:0 1em;margin-left:0}
table{border-collapse:collapse;max-width:100%}
th,td{border:1px solid #d0d7de;padding:6px 12px;vertical-align:top}
th{background:#f6f8fa;font-weight:600}
img{max-width:100%}
hr{border:0;border-top:1px solid #d8dee4;margin:2em 0}
@page{size:A4;margin:2cm}
@media print{
  body{font-size:11pt}
  main{max-width:none;padding:0}
  h1,h2,h3,h4{break-after:avoid}
  pre,tr,img,blockquote{break-inside:avoid}
}
"""

PLANTILLA = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>@TITULO@</title>
<style>@CSS@</style>
</head>
<body>
<main>
@CUERPO@
</main>
</body>
</html>
"""


class _Documento:
    """Un Markdown ya interpretado: árbol HTML, título e imágenes locales."""

    def __init__(self, origen):
        from bs4 import BeautifulSoup
        from markdown_it import MarkdownIt

        self.origen = Path(origen)
        texto = leer_texto(origen)
        if not texto.strip():
            raise SinTexto()
        # CommonMark + tablas y tachado (lo habitual en Markdown de GitHub).
        cuerpo = MarkdownIt("commonmark").enable(["table", "strikethrough"]).render(texto)
        self.arbol = BeautifulSoup(cuerpo, "html.parser")
        primero = self.arbol.find(["h1", "h2"])
        self.titulo = (primero.get_text(" ", strip=True) if primero else "") or self.origen.stem
        self._marcar_imagenes()
        self._html = None

    def _marcar_imagenes(self):
        """Anota en cada <img> local la ruta real del archivo, si existe."""
        for img in self.arbol.find_all("img"):
            fuente = (img.get("src") or "").strip()
            if not fuente or fuente.startswith(("http://", "https://", "data:")):
                continue
            if fuente.lower().startswith("file:"):
                ruta = Path(unquote(urlparse(fuente).path).lstrip("/"))
            else:
                ruta = Path(unquote(fuente))
            if not ruta.is_absolute():
                ruta = self.origen.parent / ruta
            if ruta.is_file():
                img["data-ruta"] = str(ruta)

    def html(self):
        """Página HTML completa y autónoma (imágenes locales incrustadas)."""
        if self._html is None:
            from bs4 import BeautifulSoup
            from html import escape

            copia = BeautifulSoup(str(self.arbol), "html.parser")
            for img in copia.find_all("img"):
                ruta = img.attrs.pop("data-ruta", None)
                if not ruta:
                    continue
                tipo = mimetypes.guess_type(ruta)[0] or "application/octet-stream"
                try:
                    datos = base64.b64encode(Path(ruta).read_bytes()).decode("ascii")
                except OSError:
                    continue
                img["src"] = f"data:{tipo};base64,{datos}"
            self._html = (PLANTILLA
                          .replace("@TITULO@", escape(self.titulo))
                          .replace("@CSS@", CSS)
                          .replace("@CUERPO@", str(copia)))
        return self._html


class DesdeMarkdown:
    clave = "desde_md"
    pestana = "Markdown  →  Word, PDF, HTML"
    descripcion = ("Carga archivos Markdown (.md) y conviértelos a Word, PDF, "
                   "página web o texto sin formato. Puedes marcar varios a la vez.")
    filtro = ("Markdown y texto", (".md", ".markdown", ".txt"))
    formatos = (
        (".docx", "Word (.docx)"),
        (".pdf", "PDF (.pdf)"),
        (".html", "Página web (.html)"),
        (".txt", "Texto sin formato (.txt)"),
    )
    por_defecto = (".docx",)
    boton = "Convertir"

    def preparar(self):
        import bs4  # noqa: F401  (comprobar que están antes de empezar)
        import docx  # noqa: F401
        import markdown_it  # noqa: F401

    def leer(self, origen):
        return _Documento(origen)

    def escribir(self, documento, origen, destino, extension):
        if extension == ".docx":
            _EscritorDocx(documento).guardar(destino)
        elif extension == ".pdf":
            html_a_pdf(documento.html(), destino)
        elif extension == ".html":
            destino.write_text(documento.html(), encoding="utf-8")
        elif extension == ".txt":
            destino.write_text(html_a_texto(documento.arbol), encoding="utf-8")
        else:
            raise ValueError(f"Formato desconocido: {extension}")


# ---- PDF (con Edge o Chrome, que ya vienen en Windows) --------------------

def buscar_navegador():
    """Edge o Chrome (cualquier navegador Chromium sirve para imprimir a PDF)."""
    carpetas = [os.environ.get(v) for v in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA")]
    relativas = (r"Microsoft\Edge\Application\msedge.exe", r"Google\Chrome\Application\chrome.exe")
    for relativa in relativas:
        for carpeta in carpetas:
            if carpeta and (Path(carpeta) / relativa).is_file():
                return Path(carpeta) / relativa
    fijas = (
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    )
    for ruta in fijas:
        if Path(ruta).is_file():
            return Path(ruta)
    for nombre in ("msedge", "microsoft-edge", "google-chrome", "chromium", "chromium-browser"):
        hallado = shutil.which(nombre)
        if hallado:
            return Path(hallado)
    return None


def html_a_pdf(html, destino):
    navegador = buscar_navegador()
    if navegador is None:
        raise ErrorPdf("PDF: hace falta Microsoft Edge o Chrome")
    temporal = Path(tempfile.mkdtemp(prefix="convmd_"))
    try:
        fuente = temporal / "documento.html"
        fuente.write_text(html, encoding="utf-8")
        pdf = temporal / "documento.pdf"
        try:
            subprocess.run(
                [str(navegador), "--headless=new", "--disable-gpu", "--no-first-run",
                 "--disable-extensions", "--no-pdf-header-footer", "--print-to-pdf-no-header",
                 f"--user-data-dir={temporal / 'perfil'}", f"--print-to-pdf={pdf}",
                 fuente.as_uri()],
                capture_output=True, timeout=180,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired:
            raise ErrorPdf("PDF: el navegador tardó demasiado") from None
        if not pdf.is_file() or pdf.stat().st_size == 0:
            raise ErrorPdf("PDF: el navegador no generó el archivo")
        shutil.copyfile(pdf, destino)
    finally:
        shutil.rmtree(temporal, ignore_errors=True)


# ---- Texto sin formato ----------------------------------------------------

_BLOQUES = {
    "p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "pre", "blockquote",
    "table", "hr", "div", "section", "article", "dl", "dt", "dd", "figure",
}
_TITULOS = {"h1": 1, "h2": 2, "h3": 3, "h4": 4, "h5": 5, "h6": 6}


def _es_bloque(nodo):
    return getattr(nodo, "name", None) in _BLOQUES


def _plano(nodos):
    """Texto de una secuencia de nodos en línea, con espacios normalizados."""
    partes = []
    for nodo in nodos:
        if getattr(nodo, "name", None) == "br":
            partes.append("\n")
        elif getattr(nodo, "name", None) == "img":
            partes.append(f"[imagen: {nodo.get('alt', '')}]" if nodo.get("alt") else "[imagen]")
        elif hasattr(nodo, "get_text"):
            partes.append(nodo.get_text())
        else:
            partes.append(str(nodo))
    lineas = [re.sub(r"[ \t\r\n]+", " ", l).strip() for l in "".join(partes).split("\n")]
    return "\n".join(lineas).strip()


def html_a_texto(arbol):
    salida = []

    def parrafo(texto, sangria=""):
        if texto:
            salida.append("\n".join(sangria + l for l in texto.split("\n")))
            salida.append("")

    def lista(nodo, sangria):
        numero = _entero(nodo.get("start"), 1)
        for li in nodo.find_all("li", recursive=False):
            marca = f"{numero}. " if nodo.name == "ol" else "- "
            propios = [h for h in li.children if getattr(h, "name", None) not in ("ul", "ol")]
            texto = _plano(propios).replace("\n", " ")
            salida.append(sangria + marca + texto)
            for sub in li.find_all(["ul", "ol"], recursive=False):
                lista(sub, sangria + "   ")
            numero += 1

    def recorrer(nodos, sangria=""):
        sueltos = []
        for nodo in list(nodos) + [None]:
            if nodo is not None and not _es_bloque(nodo):
                sueltos.append(nodo)
                continue
            parrafo(_plano(sueltos), sangria)
            sueltos = []
            if nodo is None:
                break
            nombre = nodo.name
            if nombre in _TITULOS:
                titulo = _plano(nodo.children)
                if titulo:
                    salida.append(titulo.upper() if nombre == "h1" else titulo)
                    salida.append(("=" if nombre == "h1" else "-") * min(len(titulo), 70) if nombre in ("h1", "h2") else "")
                    if nombre in ("h1", "h2"):
                        salida.append("")
            elif nombre in ("ul", "ol"):
                lista(nodo, sangria)
                salida.append("")
            elif nombre == "pre":
                parrafo(nodo.get_text().rstrip("\n"), sangria + "    ")
            elif nombre == "blockquote":
                recorrer(nodo.children, sangria + "    ")
            elif nombre == "table":
                for fila in nodo.find_all("tr"):
                    celdas = [_plano(c.children).replace("\n", " ") for c in fila.find_all(["th", "td"])]
                    salida.append(sangria + " | ".join(celdas))
                salida.append("")
            elif nombre == "hr":
                salida.append(sangria + "-" * 40)
                salida.append("")
            elif nombre in ("p", "dt", "dd"):
                parrafo(_plano(nodo.children), sangria + ("    " if nombre == "dd" else ""))
            else:
                recorrer(nodo.children, sangria)

    recorrer(arbol.children)
    texto = "\n".join(salida)
    return re.sub(r"\n{3,}", "\n\n", texto).strip() + "\n"


def _entero(valor, por_defecto):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return por_defecto


# ---- Word (.docx) ---------------------------------------------------------

# Orden que exige el formato de Word para los hijos de <w:pPr>.
_TRAS_SOMBRA = (
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
)
_TRAS_BORDE = ("w:shd",) + _TRAS_SOMBRA
_PASO = 0.63  # cm de sangría por nivel de lista


class _EscritorDocx:
    """Recorre el árbol HTML del Markdown y construye un documento de Word."""

    def __init__(self, documento):
        from docx import Document
        from docx.shared import Pt

        self.documento = documento
        self.doc = Document()
        normal = self.doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(11)
        self.doc.core_properties.title = documento.titulo
        self._bloques(list(documento.arbol.children), 0.0, False, 0)

    def guardar(self, destino):
        self.doc.save(str(destino))

    # -- bloques ------------------------------------------------------------

    def _parrafo(self, sangria, estilo=None):
        from docx.shared import Cm
        p = self.doc.add_paragraph(style=estilo)
        if sangria:
            p.paragraph_format.left_indent = Cm(sangria)
        return p

    # `sangria` es el margen base en cm (crece dentro de una cita) y `nivel`
    # la profundidad de lista: el contenido de un elemento de lista de
    # profundidad n se alinea a sangria + _PASO * n.

    def _bloques(self, nodos, sangria, cita, nivel):
        margen = sangria + _PASO * nivel
        sueltos = []
        for nodo in nodos + [None]:
            if nodo is not None and not _es_bloque(nodo):
                sueltos.append(nodo)
                continue
            if _tiene_contenido(sueltos):
                self._en_linea(sueltos, self._parrafo(margen), _formato(cita))
            sueltos = []
            if nodo is not None:
                self._bloque(nodo, sangria, cita, nivel)

    def _bloque(self, nodo, sangria, cita, nivel):
        nombre = nodo.name
        margen = sangria + _PASO * nivel
        if nombre in _TITULOS:
            p = self._parrafo(margen, f"Heading {_TITULOS[nombre]}")
            self._en_linea(nodo.children, p, {})
        elif nombre in ("p", "dt", "dd"):
            if _tiene_contenido(nodo.children):
                extra = _PASO if nombre == "dd" else 0
                formato = _formato(cita)
                if nombre == "dt":
                    formato["negrita"] = True
                self._en_linea(nodo.children, self._parrafo(margen + extra), formato)
        elif nombre in ("ul", "ol"):
            self._lista(nodo, sangria, cita, nivel)
        elif nombre == "pre":
            self._codigo(nodo, margen)
        elif nombre == "blockquote":
            self._bloques(list(nodo.children), sangria + 1.0, True, nivel)
        elif nombre == "table":
            self._tabla(nodo)
        elif nombre == "hr":
            self._linea(margen)
        else:
            self._bloques(list(nodo.children), sangria, cita, nivel)

    def _lista(self, nodo, sangria, cita, nivel):
        from docx.shared import Cm, Pt

        ordenada = nodo.name == "ol"
        numero = _entero(nodo.get("start"), 1)
        for li in nodo.find_all("li", recursive=False):
            hijos = list(li.children)
            cabeza, resto = [], []
            for i, hijo in enumerate(hijos):
                if _es_bloque(hijo):
                    if hijo.name == "p" and not _tiene_contenido(cabeza):
                        cabeza, resto = list(hijo.children), hijos[i + 1:]
                    else:
                        resto = hijos[i:]
                    break
                cabeza.append(hijo)

            if ordenada:
                p = self.doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(sangria + _PASO * (nivel + 1))
                p.paragraph_format.first_line_indent = Cm(-_PASO)
                p.paragraph_format.space_after = Pt(3)
                p.add_run(f"{numero}.\t")
            else:
                estilo = "List Bullet" if nivel == 0 else f"List Bullet {min(nivel + 1, 3)}"
                p = self.doc.add_paragraph(style=estilo)
                if sangria:
                    p.paragraph_format.left_indent = Cm(sangria + _PASO * (nivel + 1))
            self._en_linea(cabeza, p, _formato(cita))
            if resto:
                self._bloques(resto, sangria, cita, nivel + 1)
            numero += 1

    def _codigo(self, nodo, sangria):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt

        p = self._parrafo(sangria)
        lineas = nodo.get_text().rstrip("\n").split("\n")
        for i, linea in enumerate(lineas):
            run = p.add_run(linea)
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            if i < len(lineas) - 1:
                run.add_break()
        sombra = OxmlElement("w:shd")
        sombra.set(qn("w:val"), "clear")
        sombra.set(qn("w:color"), "auto")
        sombra.set(qn("w:fill"), "F3F4F6")
        p._p.get_or_add_pPr().insert_element_before(sombra, *_TRAS_SOMBRA)

    def _linea(self, sangria):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        p = self._parrafo(sangria)
        bordes = OxmlElement("w:pBdr")
        abajo = OxmlElement("w:bottom")
        for clave, valor in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", "BFBFBF")):
            abajo.set(qn("w:" + clave), valor)
        bordes.append(abajo)
        p._p.get_or_add_pPr().insert_element_before(bordes, *_TRAS_BORDE)

    def _tabla(self, nodo):
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        filas = [f.find_all(["th", "td"], recursive=False) for f in nodo.find_all("tr")]
        filas = [f for f in filas if f]
        if not filas:
            return
        columnas = max(len(f) for f in filas)
        tabla = self.doc.add_table(rows=len(filas), cols=columnas)
        tabla.style = "Table Grid"
        alineaciones = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}
        for i, fila in enumerate(filas):
            for j, celda in enumerate(fila):
                p = tabla.cell(i, j).paragraphs[0]
                self._en_linea(celda.children, p, {"negrita": True} if celda.name == "th" else {})
                estilo = (celda.get("style") or "") + " " + (celda.get("align") or "")
                for clave, valor in alineaciones.items():
                    if clave in estilo:
                        p.alignment = valor
        self.doc.add_paragraph()  # separa la tabla de lo que siga

    # -- contenido en línea ---------------------------------------------------

    def _en_linea(self, nodos, p, formato):
        """Añade a `p` el texto de `nodos` respetando negrita, cursiva, enlaces..."""
        # "inicio": estamos al principio o tras un espacio; ahí sobran más espacios.
        self._recorrer(nodos, p, formato, {"inicio": True})

    def _recorrer(self, nodos, p, formato, estado):
        for nodo in nodos:
            nombre = getattr(nodo, "name", None)
            if nombre is None:
                texto = str(nodo)
                if not formato.get("codigo"):
                    texto = re.sub(r"[ \t\r\n]+", " ", texto)
                    if estado["inicio"]:
                        texto = texto.lstrip()
                if texto:
                    self._run(p, texto, formato)
                    estado["inicio"] = texto.endswith(" ")
            elif nombre == "br":
                p.add_run().add_break()
                estado["inicio"] = True
            elif nombre == "img":
                self._imagen(p, nodo, formato)
                estado["inicio"] = False
            else:
                nuevo = dict(formato)
                if nombre in ("strong", "b"):
                    nuevo["negrita"] = True
                elif nombre in ("em", "i"):
                    nuevo["cursiva"] = True
                elif nombre in ("code", "kbd", "samp"):
                    nuevo["codigo"] = True
                elif nombre in ("del", "s", "strike"):
                    nuevo["tachado"] = True
                elif nombre == "sup":
                    nuevo["superindice"] = True
                elif nombre == "sub":
                    nuevo["subindice"] = True
                elif nombre == "a":
                    destino = (nodo.get("href") or "").strip()
                    if destino and not destino.startswith("#"):
                        nuevo["enlace"] = destino
                self._recorrer(nodo.children, p, nuevo, estado)

    def _run(self, p, texto, formato):
        from docx.shared import Pt, RGBColor

        run = p.add_run(texto)
        if formato.get("negrita"):
            run.bold = True
        if formato.get("cursiva"):
            run.italic = True
        if formato.get("tachado"):
            run.font.strike = True
        if formato.get("superindice"):
            run.font.superscript = True
        if formato.get("subindice"):
            run.font.subscript = True
        if formato.get("codigo"):
            run.font.name = "Consolas"
            run.font.size = Pt(10)
        if formato.get("cita"):
            run.font.color.rgb = RGBColor(0x59, 0x63, 0x6E)
        if formato.get("enlace"):
            self._enlazar(p, run, formato["enlace"])
        return run

    def _enlazar(self, p, run, url):
        from docx.opc.constants import RELATIONSHIP_TYPE
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import RGBColor

        run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        run.font.underline = True
        try:
            relacion = p.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        except Exception:
            return  # se queda como texto con aspecto de enlace
        enlace = OxmlElement("w:hyperlink")
        enlace.set(qn("r:id"), relacion)
        enlace.append(run._r)
        p._p.append(enlace)

    def _imagen(self, p, nodo, formato):
        from docx.shared import Cm

        ruta = nodo.get("data-ruta")
        if ruta:
            try:
                imagen = p.add_run().add_picture(ruta)
                ancho_maximo = Cm(15.5)
                if imagen.width > ancho_maximo:
                    imagen.height = int(imagen.height * ancho_maximo / imagen.width)
                    imagen.width = ancho_maximo
                return
            except Exception:
                pass
        alt = nodo.get("alt") or nodo.get("src") or ""
        self._run(p, f"[imagen: {alt}]" if alt else "[imagen]", formato)


def _formato(cita):
    return {"cita": True, "cursiva": True} if cita else {}


def _tiene_contenido(nodos):
    for nodo in nodos:
        if getattr(nodo, "name", None) is not None:
            if nodo.name == "img" or nodo.get_text(strip=True):
                return True
        elif str(nodo).strip():
            return True
    return False
