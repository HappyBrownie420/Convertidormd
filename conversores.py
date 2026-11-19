# -*- coding: utf-8 -*-
"""Motor del Convertidor Markdown (sin interfaz).

Dos sentidos:
  * AMarkdown:      PDF, Word, PowerPoint, Excel, HTML... -> Markdown (MarkItDown)
  * DesdeMarkdown:  Markdown -> Word (.docx), PDF, Excel (.xlsx), HTML o texto

Cada sentido es un "modo" con la misma forma: `preparar()` carga lo pesado,
`leer(origen)` interpreta el archivo una sola vez y `escribir(...)` genera
cada formato de salida pedido.
"""
import base64
import html as html_lib
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
        if Path(origen).suffix.lower() in (".html", ".htm", ".xhtml"):
            texto = html_a_markdown(origen)
            if not texto.strip():
                raise SinTexto()
            return texto
        resultado = self._conversor.convert(str(origen))
        texto = getattr(resultado, "markdown", None) or getattr(resultado, "text_content", "") or ""
        if not texto.strip():
            raise SinTexto()
        return texto

    def escribir(self, texto, origen, destino, extension):
        destino.write_text(texto, encoding="utf-8")


# ---- HTML -> Markdown (también páginas hechas con JavaScript) -------------
# Muchas páginas (p. ej. las clases interactivas por bloques) no traen el
# texto escrito en el HTML: lo genera JavaScript al abrirlas. Para verlas
# como las ve una persona se abren con Edge/Chrome sin ventana, con
# recorrer_pagina.js inyectado, que pasa por todos los bloques y revela las
# respuestas de los quizzes. Las fórmulas LaTeX (\( \), \[ \], $$) se
# conservan tal cual para que el Markdown las muestre como fórmulas.

SCRIPT_RECORRIDO = Path(__file__).with_name("recorrer_pagina.js")
# Scripts que convierten el LaTeX en dibujos: se quitan para conservar el LaTeX.
_SCRIPTS_DE_FORMULAS = re.compile(
    r"<script\b[^>]*\bsrc=[\"'][^\"']*(mathjax|katex|auto-render)[^\"']*[\"'][^>]*>\s*</script>",
    re.I)
_CLASE_AVISO = re.compile(
    r"analog|warn|alert|note|nota|tip|hint|callout|info|important|highlight|ojo|"
    r"danger|caution|remember|recuerda|key|clave|ok-box|success|error-box|box|card-note", re.I)
_CLASE_EXPLICACION = re.compile(r"feedback|explica|solution|solucion|answer|respuesta|retro", re.I)
_CLASE_CORRECTA = re.compile(r"(^|[\s_-])(correct|correcta|right|ok)($|[\s_-])", re.I)
_NAVEGACION = re.compile(
    r"^(siguiente|anterior|next|prev(ious)?|volver|atrás|continuar|inicio|fin)\b|[←→‹›«»◀▶✓]", re.I)
_FORMULA = re.compile(r"\\\[(.+?)\\\]|\$\$(.+?)\$\$|\\\((.+?)\\\)", re.S)


def renderizar_html(origen, html):
    """HTML tal como queda tras ejecutar su JavaScript, o None si no se puede."""
    navegador = buscar_navegador()
    if navegador is None or not SCRIPT_RECORRIDO.is_file():
        return None
    script = SCRIPT_RECORRIDO.read_text(encoding="utf-8")
    pagina = _SCRIPTS_DE_FORMULAS.sub("", html)
    # Las rutas relativas (imágenes, scripts propios) se resuelven desde la carpeta original.
    base = f'<base href="{Path(origen).resolve().parent.as_uri()}/">'
    if re.search(r"<head[^>]*>", pagina, re.I):
        pagina = re.sub(r"(<head[^>]*>)", lambda m: m.group(1) + base, pagina, count=1, flags=re.I)
    else:
        pagina = base + pagina
    inyectado = f"<script>{script}</script>"
    if re.search(r"</body\s*>", pagina, re.I):
        pagina = re.sub(r"</body\s*>", lambda m: inyectado + m.group(0), pagina, count=1, flags=re.I)
    else:
        pagina += inyectado
    temporal = Path(tempfile.mkdtemp(prefix="convmd_"))
    try:
        fuente = temporal / "pagina.html"
        fuente.write_text(pagina, encoding="utf-8")
        try:
            resultado = subprocess.run(
                [str(navegador), "--headless=new", "--disable-gpu", "--no-first-run",
                 "--disable-extensions", "--mute-audio", f"--user-data-dir={temporal / 'perfil'}",
                 "--virtual-time-budget=60000", "--dump-dom", fuente.as_uri()],
                capture_output=True, timeout=180, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except (OSError, subprocess.TimeoutExpired):
            return None
        salida = resultado.stdout.decode("utf-8", "replace")
        return salida if "__convmd__" in salida else None
    finally:
        shutil.rmtree(temporal, ignore_errors=True)


def _preparar_html(html):
    """Ordena el HTML para que salga un Markdown legible. Devuelve (html, fórmulas)."""
    from bs4 import BeautifulSoup

    sopa = BeautifulSoup(html, "html.parser")
    titulo = sopa.title.get_text(" ", strip=True) if sopa.title else ""
    principal = sopa.find(id="__convmd__")
    recorrida = principal is not None and principal.has_attr("data-pasos")
    raiz = principal or sopa.body or sopa
    for basura in raiz.find_all(["script", "style", "noscript", "template", "link", "meta"]):
        basura.decompose()

    # KaTeX o MathML ya dibujados: recuperar el LaTeX original si viene anotado.
    for nota in raiz.find_all("annotation", attrs={"encoding": "application/x-tex"}):
        contenedor = nota.find_parent(class_="katex-display") or nota.find_parent(class_="katex") \
            or nota.find_parent("math")
        if contenedor is not None:
            bloque = contenedor.name == "math" and contenedor.get("display") == "block" \
                or "katex-display" in (contenedor.get("class") or [])
            tex = nota.get_text()
            contenedor.replace_with(f"\\[{tex}\\]" if bloque else f"\\({tex}\\)")

    # Opciones de quiz (varios botones juntos) -> lista, marcando la correcta.
    for grupo in list(raiz.find_all(True)):
        hijos = [h for h in grupo.find_all(["button", "input"], recursive=False)
                 if h.name == "button" or h.get("type") == "button"]
        if len(hijos) < 2 or grupo.find_parent(attrs={"data-convmd": "lista"}):
            continue
        lista = sopa.new_tag("ul")
        lista["data-convmd"] = "lista"
        for boton in hijos:
            item = sopa.new_tag("li")
            for parte in list(boton.contents):
                item.append(parte.extract())
            if boton.name == "input":
                item.append(boton.get("value", ""))
            clases = " ".join(boton.get("class") or [])
            if _CLASE_CORRECTA.search(clases) and "incorrect" not in clases.lower():
                item.append(" ✅")
            lista.append(item)
        hijos[0].insert_before(lista)
        for boton in hijos:
            boton.decompose()
    # Botones sueltos de navegación (Siguiente, Anterior…) no aportan nada.
    for boton in raiz.find_all("button"):
        if _NAVEGACION.search(boton.get_text(" ", strip=True)) or not boton.get_text(strip=True):
            boton.decompose()

    # Recuadros de explicación, avisos y analogías -> citas (>).
    for caja in list(raiz.find_all(["div", "aside", "section", "p"])):
        if caja.parent is None or caja.find_parent("blockquote"):
            continue
        clases = " ".join(caja.get("class") or [])
        if not clases or "formula" in clases.lower():
            continue
        es_explicacion = bool(_CLASE_EXPLICACION.search(clases))
        if not es_explicacion and not _CLASE_AVISO.search(clases):
            continue
        if not caja.get_text(strip=True):
            caja.decompose()
            continue
        if caja.find(["h1", "h2", "table"]) or len(caja.get_text()) > 3000:
            continue            # es una sección entera, no un recuadro
        primero = next((h for h in caja.children if getattr(h, "name", None) or str(h).strip()), None)
        if getattr(primero, "name", None) in ("b", "strong"):
            primero.insert_after(sopa.new_tag("br"))
        elif es_explicacion:
            etiqueta = sopa.new_tag("strong")
            etiqueta.string = "Explicación:"
            caja.insert(0, " ")
            caja.insert(0, etiqueta)
        caja.name = "blockquote"
        caja.attrs = {}

    # Etiquetas pequeñas sobre los títulos (p. ej. «Definición») -> cursiva.
    for etiqueta in raiz.find_all(class_=re.compile(r"eyebrow|kicker|overline|subtitle|tag", re.I)):
        if etiqueta.name in ("div", "span", "p") and etiqueta.get_text(strip=True) \
                and not etiqueta.find(["div", "p", "ul", "ol", "table"]):
            texto = etiqueta.get_text(" ", strip=True)
            etiqueta.clear()
            cursiva = sopa.new_tag("em")
            cursiva.string = texto
            etiqueta.append(cursiva)
            etiqueta.name = "p"

    cuerpo = raiz.decode_contents() if raiz is not sopa else str(sopa)

    # Fórmulas: se apartan con una marca para que el conversor no las toque.
    formulas = []

    def apartar(m):
        bloque = m.group(1) is not None or m.group(2) is not None
        tex = next(g for g in m.groups() if g is not None)
        formulas.append((bloque, html_lib.unescape(tex)))
        return f"CONVMDFORMULA{len(formulas) - 1:05d}X"

    cuerpo = _FORMULA.sub(apartar, cuerpo)
    primer_titulo = raiz.find(["h1", "h2"])
    repetido = primer_titulo is not None and primer_titulo.get_text(" ", strip=True) == titulo
    if titulo and recorrida and not repetido:
        cuerpo = f"<h1>{html_lib.escape(titulo)}</h1>" + cuerpo
    return f"<html><head><title>{html_lib.escape(titulo)}</title></head><body>{cuerpo}</body></html>", formulas


def html_a_markdown(origen):
    """Convierte un archivo HTML a Markdown, ejecutando antes su JavaScript."""
    from markitdown.converters import HtmlConverter

    html = leer_texto(origen)
    renderizado = None
    if re.search(r"<script\b", html, re.I):
        renderizado = renderizar_html(origen, html)
    limpio, formulas = _preparar_html(renderizado or _SCRIPTS_DE_FORMULAS.sub("", html))
    texto = HtmlConverter().convert_string(limpio).markdown or ""

    def devolver(m):
        bloque, tex = formulas[int(m.group(1))]
        tex = tex.strip()
        if bloque:
            return "$$ " + re.sub(r"\s*\n\s*", " ", tex) + " $$"
        return "$" + tex + "$"

    texto = re.sub(r"CONVMDFORMULA(\d{5})X", devolver, texto)
    texto = re.sub(r"(?m)^>[ \t]{2,}", "> ", texto)
    # Título en negrita de un recuadro: en su propia línea dentro de la cita.
    texto = re.sub(r"(?m)^(> \*\*[^\n]*\*\*)[ \t]*\n> ", r"\1\n>\n> ", texto)
    texto = re.sub(r"\$\$\n\$\$", "$$\n\n$$", texto)     # cada fórmula grande en su párrafo
    texto = re.sub(r"[ \t]+\n", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip() + "\n"


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
    pestana = "Markdown  →  Word, PDF, Excel…"
    descripcion = ("Carga archivos Markdown (.md) y conviértelos a Word, PDF, Excel, página web "
                   "o texto. Puedes marcar varios. En Excel, cada tabla va a su propia hoja.")
    filtro = ("Markdown y texto", (".md", ".markdown", ".txt"))
    formatos = (
        (".docx", "Word (.docx)"),
        (".pdf", "PDF (.pdf)"),
        (".xlsx", "Excel (.xlsx)"),
        (".html", "Página web (.html)"),
        (".txt", "Texto (.txt)"),
    )
    por_defecto = (".docx",)
    boton = "Convertir"

    def preparar(self):
        import bs4  # noqa: F401  (comprobar que están antes de empezar)
        import docx  # noqa: F401
        import markdown_it  # noqa: F401
        import openpyxl  # noqa: F401

    def leer(self, origen):
        return _Documento(origen)

    def escribir(self, documento, origen, destino, extension):
        if extension == ".docx":
            _EscritorDocx(documento).guardar(destino)
        elif extension == ".xlsx":
            html_a_excel(documento.arbol, destino)
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


# ---- Excel (.xlsx) --------------------------------------------------------
# Cada tabla del Markdown va a una hoja, con el nombre del título que la
# precede. El resto del documento (si lo hay) va a una hoja «Texto».

_PROHIBIDOS_EN_HOJA = re.compile(r"[\[\]:*?/\\]")


def _nombre_hoja(propuesta, usados):
    """Nombre de hoja válido para Excel (31 caracteres, sin []:*?/\\) y sin repetir."""
    nombre = _PROHIBIDOS_EN_HOJA.sub(" ", propuesta or "")
    nombre = re.sub(r"\s+", " ", nombre).strip().strip("'")[:31].strip() or "Hoja"
    base, numero = nombre, 2
    while nombre.lower() in usados:
        sufijo = f" ({numero})"
        nombre = base[:31 - len(sufijo)].rstrip() + sufijo
        numero += 1
    usados.add(nombre.lower())
    return nombre


def _valor_celda(texto):
    """Convierte a número lo que claramente lo es, para poder calcular en Excel."""
    limpio = texto.strip()
    if limpio in ("", "NaN"):                       # NaN: celda vacía según MarkItDown
        return None
    sin_signo = limpio.lstrip("-")
    if re.fullmatch(r"-?\d{1,15}", limpio) and not (len(sin_signo) > 1 and sin_signo[0] == "0"):
        return int(limpio)                          # "007" y números larguísimos son códigos
    if re.fullmatch(r"-?\d+\.\d+", limpio):
        return float(limpio)
    # Coma decimal ("12,5"). "1,234" es ambiguo (¿miles?) y se deja como texto.
    if re.fullmatch(r"-?\d+,\d+", limpio) and not re.fullmatch(r"-?\d{1,3},\d{3}", limpio):
        return float(limpio.replace(",", "."))
    return texto


def _escribir_celda(hoja, fila, columna, valor):
    celda = hoja.cell(row=fila, column=columna, value=valor)
    if isinstance(valor, str) and valor.startswith("="):
        celda.data_type = "s"      # texto literal: nunca ejecutar fórmulas del Markdown
    return celda


def _filas_de_lista(nodo, nombres, sangria):
    numero = _entero(nodo.get("start"), 1)
    for li in nodo.find_all("li", recursive=False):
        marca = f"{numero}. " if nodo.name == "ol" else "• "
        propios = [h for h in li.children if getattr(h, "name", None) not in ("ul", "ol", "table", "pre")]
        yield sangria + marca + _plano(propios).replace("\n", " "), "p"
        for sub in li.find_all(["ul", "ol"], recursive=False):
            yield from _filas_de_lista(sub, nombres, sangria + "    ")
        numero += 1


def _filas_de_texto(nodos, nombres, sangria=""):
    """Recorre los bloques y da (texto, tipo) por cada fila de la hoja «Texto»."""
    sueltos = []
    for nodo in list(nodos) + [None]:
        if nodo is not None and not _es_bloque(nodo):
            sueltos.append(nodo)
            continue
        texto = _plano(sueltos).replace("\n", " ")
        sueltos = []
        if texto:
            yield sangria + texto, "p"
        if nodo is None:
            break
        nombre = nodo.name
        if nombre in _TITULOS:
            texto = _plano(nodo.children).replace("\n", " ")
            if texto:
                yield texto, nombre
        elif nombre in ("ul", "ol"):
            yield from _filas_de_lista(nodo, nombres, sangria)
        elif nombre == "pre":
            for linea in nodo.get_text().rstrip("\n").split("\n"):
                yield sangria + linea, "codigo"
        elif nombre == "blockquote":
            yield from _filas_de_texto(nodo.children, nombres, sangria + "    ")
        elif nombre == "table":
            yield f"[Tabla en la hoja «{nombres.get(id(nodo), '')}»]", "nota"
        elif nombre in ("p", "dt", "dd"):
            texto = _plano(nodo.children).replace("\n", " ")
            if texto:
                yield sangria + texto, "p"
        elif nombre != "hr":
            yield from _filas_de_texto(nodo.children, nombres, sangria)


def html_a_excel(arbol, destino):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    libro = Workbook()
    libro.remove(libro.active)
    usados, nombres = set(), {}
    negrita = Font(bold=True)
    relleno = PatternFill("solid", start_color="E8EEF7", end_color="E8EEF7")

    # Una hoja por tabla
    titulo, numero = "", 0
    for nodo in arbol.find_all([*_TITULOS, "table"]):
        if nodo.name != "table":
            titulo = _plano(nodo.children).replace("\n", " ")
            continue
        numero += 1
        nombre = _nombre_hoja(titulo or f"Tabla {numero}", usados)
        nombres[id(nodo)] = nombre
        hoja = libro.create_sheet(nombre)
        anchos, con_cabecera = {}, False
        for i, fila in enumerate(nodo.find_all("tr"), start=1):
            for j, celda in enumerate(fila.find_all(["th", "td"], recursive=False), start=1):
                texto = _plano(celda.children)
                es_cabecera = celda.name == "th"
                if es_cabecera:
                    valor = None if re.fullmatch(r"Unnamed: \d+", texto) else (texto or None)
                else:
                    valor = _valor_celda(texto)
                c = _escribir_celda(hoja, i, j, valor)
                estilo = (celda.get("style") or "") + " " + (celda.get("align") or "")
                horizontal = "center" if "center" in estilo else "right" if "right" in estilo else None
                c.alignment = Alignment(horizontal=horizontal, vertical="top", wrap_text="\n" in texto)
                if es_cabecera:
                    c.font, c.fill = negrita, relleno
                    con_cabecera = con_cabecera or i == 1
                enlaces = [a["href"] for a in celda.find_all("a", href=True) if not a["href"].startswith("#")]
                if len(enlaces) == 1 and isinstance(valor, str):
                    c.hyperlink = enlaces[0]
                    c.font = Font(color="0563C1", underline="single", bold=es_cabecera)
                largo = max((len(linea) for linea in texto.split("\n")), default=0)
                anchos[j] = max(anchos.get(j, 0), largo)
        for j, largo in anchos.items():
            hoja.column_dimensions[get_column_letter(j)].width = min(max(largo + 3, 9), 60)
        if con_cabecera:
            hoja.freeze_panes = "A2"

    # El resto del documento, una línea por fila
    filas = list(_filas_de_texto(list(arbol.children), nombres))
    if not nombres or any(tipo in ("p", "codigo") for _, tipo in filas):
        hoja = libro.create_sheet(_nombre_hoja("Texto", usados))
        fuentes = {
            "h1": Font(bold=True, size=16), "h2": Font(bold=True, size=13),
            "h3": Font(bold=True, size=12), "h4": negrita, "h5": negrita, "h6": negrita,
            "codigo": Font(name="Consolas", size=10),
            "nota": Font(italic=True, color="6B7280"),
        }
        i = 0
        for texto, tipo in filas:
            if tipo in _TITULOS and i:
                i += 1                              # fila en blanco antes de cada título
            i += 1
            c = _escribir_celda(hoja, i, 1, texto)
            c.alignment = Alignment(vertical="top", wrap_text=tipo != "codigo")
            if tipo in fuentes:
                c.font = fuentes[tipo]
        hoja.column_dimensions["A"].width = 110
    libro.save(str(destino))


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
