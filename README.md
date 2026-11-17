# Convertidor Markdown

Aplicación de escritorio sencilla para convertir documentos **a Markdown** y
Markdown **a otros formatos**, por lotes y sin usar la terminal.

| Pestaña | Entrada | Salida |
|---|---|---|
| **Archivos → Markdown** | PDF, Word, PowerPoint, Excel, HTML, CSV, JSON, XML, EPUB, ZIP… | `.md` o `.txt` |
| **Markdown → Word, PDF, Excel…** | `.md`, `.markdown`, `.txt` | `.docx`, `.pdf`, `.xlsx`, `.html`, `.txt` |

La conversión a Markdown la hace [MarkItDown](https://github.com/microsoft/markitdown)
de Microsoft; esta app le pone una ventana encima y añade el camino de vuelta.

## Qué hace

- Carga uno o varios archivos (o una carpeta entera) y crea **un documento por cada uno**.
- Tú decides dónde se guarda: junto al original o en la carpeta que elijas.
- En el sentido inverso puedes marcar varios formatos a la vez (por ejemplo Word y PDF).
- **A Excel**: cada tabla del Markdown va a su propia hoja, con el nombre del título
  que la precede. Los números se guardan como números (también con coma decimal,
  `12,5`) para poder sumar; el resto del documento va a una hoja «Texto».
- No pisa nada sin permiso: si el archivo ya existe crea `nombre (2).md`, salvo que
  marques «Reemplazar».
- Si algo falla, la columna **Estado** dice qué archivo y por qué; al pulsar
  Convertir otra vez solo se repite lo que faltó.
- Doble clic en una fila abre el resultado.
- Recuerda tus opciones entre sesiones (`config.json`, que no se sube al repositorio).

## Requisitos

- Windows 10/11 (es donde está probada).
- [Python 3.10 o superior](https://www.python.org/downloads/) con Tkinter (viene incluido).
- Para crear PDF: Microsoft Edge o Google Chrome (Edge ya viene con Windows).

## Instalación

```bash
git clone <URL-de-tu-repositorio>
cd markitdown-app
py -m pip install -r requirements.txt
```

El paso de `pip` es opcional: si faltan componentes, la app lo detecta al abrirse
y ofrece instalarlos.

## Uso

Doble clic en **`Convertidor Markdown.bat`**, o desde una terminal:

```bash
py convertidor_markdown.pyw
```

Para tener la app en el Escritorio y en el menú Inicio, con su icono:

```bash
powershell -ExecutionPolicy Bypass -File herramientas\crear_accesos_directos.ps1
```

Esos accesos directos apuntan al lanzador de Python (`pyw.exe`), así que se pueden
**anclar a Inicio o a la barra de tareas** con clic derecho (Windows no lo permite
con un `.bat`) y siguen sirviendo aunque actualices Python.

También puedes arrastrar archivos sobre el acceso directo o sobre el `.bat` para
abrir la app con ellos ya cargados: los `.md` van a la pestaña inversa y el resto
a «Archivos → Markdown».

## Si actualizas Python

La app no está atada a una versión concreta de Python. Al abrirse:

1. Si el Python con el que arranca tiene todos los componentes, lo usa.
2. Si no, busca otro Python instalado que sí los tenga y se reabre con él.
3. Si ninguno los tiene (lo normal justo después de instalar una versión nueva),
   pregunta y los instala con `pip`.

La versión de Python y de MarkItDown en uso se ve arriba a la derecha de la ventana.

## Limitaciones

- **PDF escaneados** (solo imágenes): no tienen texto que extraer y aparecen como
  «Sin texto». Harían falta OCR o un modelo de visión, que esta app no configura.
- **Markdown → Word**: se conservan títulos, negrita, cursiva, tachado, enlaces,
  listas, tablas, citas, bloques de código e imágenes locales. No hay notas al pie
  ni listas de tareas con casillas.
- **Markdown → PDF** se genera imprimiendo el HTML con el navegador, así que el
  aspecto es el de la página web.
- **Markdown → Excel**: `1,234` se deja como texto porque no se puede saber si es
  «mil doscientos» o «uno coma dos»; tampoco se convierten fechas ni porcentajes.
  Un texto que empiece por `=` se guarda como texto, nunca como fórmula.
- **Excel → Markdown → Excel**: MarkItDown lee un código como `007` como el número 7,
  así que los ceros a la izquierda se pierden en ese viaje.
- En macOS y Linux el código está preparado (`python convertidor_markdown.pyw`),
  pero no se ha probado.

## Estructura

| Archivo | Para qué |
|---|---|
| `convertidor_markdown.pyw` | La ventana (Tkinter) y el arranque |
| `conversores.py` | El motor de conversión en ambos sentidos, sin interfaz |
| `Convertidor Markdown.bat` | Lanzador de doble clic para Windows |
| `requirements.txt` | Componentes de Python necesarios |
| `icono.ico`, `icono.png` | Icono de la ventana y de los accesos directos |
| `herramientas/` | Crear los accesos directos y regenerar el icono |
| `pruebas/` | Prueba del motor con un Markdown de muestra |

## Pruebas

```bash
py pruebas/prueba_motor.py
```

Convierte `pruebas/muestra.md` a todos los formatos en una carpeta temporal y
comprueba que cada uno se puede volver a leer y conserva el contenido, incluida
la ida y vuelta Excel → Markdown → Excel.

## Créditos

- [MarkItDown](https://github.com/microsoft/markitdown) (Microsoft) — conversión a Markdown.
- [markdown-it-py](https://github.com/executablebooks/markdown-it-py) — lectura de Markdown.
- [python-docx](https://github.com/python-openxml/python-docx) — creación de documentos de Word.
- [openpyxl](https://openpyxl.readthedocs.io/) — creación de libros de Excel.

## Licencia

[MIT](LICENSE): puedes usar, modificar y redistribuir el código libremente
manteniendo el aviso de autoría. Se entrega «tal cual», sin garantía.
