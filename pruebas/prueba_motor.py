# -*- coding: utf-8 -*-
"""Prueba del motor sin interfaz.

Convierte muestra.md a Word, PDF, HTML y texto en una carpeta temporal y
comprueba que cada resultado se puede volver a leer y conserva el contenido.

    py pruebas/prueba_motor.py
"""
import sys
import tempfile
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI.parent))
import conversores as c  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CLAVES = [
    "Guía de rescate", "Subpunto anidado", "Describir los síntomas", "Crítico",
    "paciencia", "def saludar", "ñ, á, é, ü", "enlace a un sitio",
]
fallos = []


def comprobar(condicion, mensaje):
    print(("  ok    " if condicion else "  FALLO ") + mensaje)
    if not condicion:
        fallos.append(mensaje)


with tempfile.TemporaryDirectory(prefix="convmd_prueba_") as temporal:
    salida = Path(temporal)
    origen = AQUI / "muestra.md"

    print("Markdown -> otros formatos")
    modo = c.DesdeMarkdown()
    modo.preparar()
    documento = modo.leer(origen)
    comprobar(documento.titulo == "Guía de rescate: primeros pasos", "título tomado del primer encabezado")

    usadas = set()
    creados = {}
    for extension, _ in modo.formatos:
        destino = c.ruta_de_salida(origen, salida, extension, False, usadas)
        if extension == ".pdf" and c.buscar_navegador() is None:
            print("  (sin Edge ni Chrome: se omite el PDF)")
            continue
        try:
            modo.escribir(documento, origen, destino, extension)
        except Exception as error:
            comprobar(False, f"{extension}: {type(error).__name__}: {error}")
            continue
        creados[extension] = destino
        comprobar(destino.stat().st_size > 0, f"{extension} creado ({destino.stat().st_size} bytes)")

    if ".txt" in creados:
        texto = creados[".txt"].read_text(encoding="utf-8")
        comprobar("~~" not in texto and "**" not in texto, ".txt sin marcas de Markdown")
        comprobar("   - Subpunto anidado" in texto, ".txt conserva la lista anidada")

    if ".docx" in creados:
        from docx import Document
        word = Document(str(creados[".docx"]))
        estilos = [p.style.name for p in word.paragraphs]
        comprobar("Heading 1" in estilos and "Heading 2" in estilos, ".docx con títulos")
        comprobar("List Bullet 2" in estilos, ".docx con lista anidada")
        comprobar(len(word.tables) == 1, ".docx con la tabla")
        comprobar(len(word.inline_shapes) == 1, ".docx con la imagen")

    print("Vuelta a Markdown (MarkItDown)")
    ida = c.AMarkdown()
    ida.preparar()
    for extension in (".docx", ".pdf", ".html"):
        if extension not in creados:
            continue
        try:
            texto = ida.leer(creados[extension])
        except Exception as error:
            comprobar(False, f"leer {extension}: {type(error).__name__}: {error}")
            continue
        faltan = [clave for clave in CLAVES if clave not in texto]
        comprobar(not faltan, f"{extension} conserva el contenido" + (f" (faltan: {faltan})" if faltan else ""))

    print("Nombres de salida")
    usadas = set()
    primera = c.ruta_de_salida(origen, AQUI, ".md", True, usadas)
    comprobar(primera != origen, "nunca se escribe sobre el archivo de origen")
    segunda = c.ruta_de_salida(origen, AQUI, ".md", True, usadas)
    comprobar(segunda != primera, "dos salidas de la misma tanda no chocan")

print()
print("TODO BIEN" if not fallos else f"{len(fallos)} comprobación(es) fallida(s)")
sys.exit(1 if fallos else 0)
