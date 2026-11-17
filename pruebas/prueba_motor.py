# -*- coding: utf-8 -*-
"""Prueba del motor sin interfaz.

Convierte muestra.md a Word, PDF, Excel, HTML y texto en una carpeta temporal y
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

    if ".xlsx" in creados:
        from openpyxl import load_workbook
        libro = load_workbook(creados[".xlsx"])
        comprobar(libro.sheetnames == ["Tabla", "Texto"],
                  f".xlsx: una hoja por tabla más la de texto ({libro.sheetnames})")
        hoja = libro["Tabla"]
        comprobar([celda.value for celda in hoja[1]] == ["Especie", "Peso (kg)", "Estado"], ".xlsx con cabecera")
        comprobar(hoja["B2"].value == 12.5 and hoja["B3"].value == 3.8, ".xlsx guarda los números como números")
        comprobar(hoja["C3"].value == "Crítico" and hoja.freeze_panes == "A2", ".xlsx con texto y cabecera fija")
        comprobar(libro["Texto"]["A1"].value == "Guía de rescate: primeros pasos", ".xlsx con el resto del documento")

    print("Vuelta a Markdown (MarkItDown)")
    ida = c.AMarkdown()
    ida.preparar()

    # Excel -> Markdown -> Excel: mismas hojas y mismos valores
    from openpyxl import Workbook, load_workbook
    original = Workbook()
    ventas = original.active
    ventas.title = "Ventas"
    for fila in (["Producto", "Unidades", "Precio"], ["Pienso", 12, 19.95], ["Correa", None, 7.5], ["Collar", "A-007", 3]):
        ventas.append(fila)
    gastos = original.create_sheet("Gastos 2026")
    for fila in (["Concepto", "Importe"], ["Veterinario", 250], ["Transporte", 40.25]):
        gastos.append(fila)
    original.save(salida / "libro.xlsx")
    (salida / "libro.md").write_text(ida.leer(salida / "libro.xlsx"), encoding="utf-8")
    modo.escribir(modo.leer(salida / "libro.md"), salida / "libro.md", salida / "libro_vuelta.xlsx", ".xlsx")
    vuelta = load_workbook(salida / "libro_vuelta.xlsx")
    comprobar(vuelta.sheetnames == ["Ventas", "Gastos 2026"], f"ida y vuelta de Excel: mismas hojas ({vuelta.sheetnames})")
    iguales = all(
        [[c_.value for c_ in fila] for fila in vuelta[nombre].iter_rows()]
        == [[c_.value for c_ in fila] for fila in original[nombre].iter_rows()]
        for nombre in original.sheetnames if nombre in vuelta.sheetnames
    )
    comprobar(iguales, "ida y vuelta de Excel: mismos valores (números, textos y celdas vacías)")
    casos = {"12": 12, "-3": -3, "3.8": 3.8, "12,5": 12.5, "007": "007", "1,234": "1,234",
             "NaN": None, "": None, "Perro": "Perro", "12 kg": "12 kg"}
    mal = {texto: c._valor_celda(texto) for texto, esperado in casos.items() if c._valor_celda(texto) != esperado}
    comprobar(not mal, "solo se convierte a número lo inequívoco" + (f" (mal: {mal})" if mal else ""))

    # Un texto que parece fórmula debe quedar como texto
    (salida / "formula.md").write_text("| a |\n|---|\n| =1+1 |\n", encoding="utf-8")
    modo.escribir(modo.leer(salida / "formula.md"), salida / "formula.md", salida / "formula.xlsx", ".xlsx")
    celda = load_workbook(salida / "formula.xlsx").active["A2"]
    comprobar(celda.data_type == "s" and celda.value == "=1+1", "las fórmulas del Markdown no se ejecutan en Excel")

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
