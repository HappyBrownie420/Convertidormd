# -*- coding: utf-8 -*-
"""Convertidor Markdown: ventana principal.

Dos pestañas, una por sentido:
  * Archivos -> Markdown (PDF, Word, PowerPoint, Excel, HTML...)
  * Markdown -> Word, PDF, HTML o texto

El trabajo de conversión está en conversores.py; aquí solo vive la interfaz
y el arranque (encontrar un Python que tenga los componentes necesarios).
"""
import importlib
import importlib.util
import json
import os
import queue
import re
import subprocess
import sys
import threading
from pathlib import Path

# Con pythonw no hay consola: algunas bibliotecas escriben avisos en
# stdout/stderr y fallarían si valen None.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from conversores import AMarkdown, DesdeMarkdown, ErrorAlGuardar, explicar_error, ruta_de_salida

TITULO = "Convertidor Markdown"
ARCHIVO_CONFIG = Path(__file__).with_name("config.json")
PENDIENTE = "Pendiente"
SIN_VENTANA = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def abrir(ruta):
    """Abre un archivo o carpeta con el programa predeterminado del sistema."""
    if hasattr(os, "startfile"):
        os.startfile(ruta)
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(ruta)])


# --------------------------------------------------------------------------
# Una pestaña: lista de archivos + opciones + conversión en segundo plano
# --------------------------------------------------------------------------

class Panel(ttk.Frame):
    def __init__(self, padre, modo, config, escala):
        m = int(14 * escala)
        super().__init__(padre, padding=m)
        self.modo = modo
        self.escala = escala
        self.cola = queue.Queue()
        self.trabajando = False
        self.origenes = {}        # id de fila -> Path del archivo original
        self.hechos = {}          # id de fila -> {extensión: Path creado}
        self.ultima_carpeta = None
        self._creados = self._con_problemas = 0
        self._fallo_al_guardar = False

        elegidos = config.get("formatos") or list(modo.por_defecto)
        self.destino = tk.StringVar(value=config.get("destino", "junto"))
        self.carpeta = tk.StringVar(value=config.get("carpeta", ""))
        self.formatos = {ext: tk.BooleanVar(value=ext in elegidos) for ext, _ in modo.formatos}
        self.sobrescribir = tk.BooleanVar(value=bool(config.get("sobrescribir", False)))
        self.estado = tk.StringVar(value="Añade archivos para empezar.")

        self._construir(m)
        self._actualizar_controles()
        self.after(80, self._revisar_cola)

    # ---- construcción -----------------------------------------------------

    def _construir(self, m):
        e = self.escala
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ttk.Label(self, text=self.modo.descripcion, style="Sub.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, m))

        zona = ttk.Frame(self)
        zona.grid(row=1, column=0, sticky="nsew")
        zona.columnconfigure(0, weight=1)
        zona.rowconfigure(0, weight=1)

        self.tabla = ttk.Treeview(
            zona, columns=("archivo", "carpeta", "estado"), show="headings", selectmode="extended")
        for columna, titulo, ancho in (("archivo", "Archivo", 230), ("carpeta", "Carpeta", 170),
                                       ("estado", "Estado", 380)):
            self.tabla.heading(columna, text=titulo, anchor="w")
            self.tabla.column(columna, width=int(ancho * e), anchor="w")
        self.tabla.tag_configure("ok", foreground="#1a7f37")
        self.tabla.tag_configure("error", foreground="#c62828")
        self.tabla.tag_configure("aviso", foreground="#9a6700")
        self.tabla.tag_configure("proceso", foreground="#0b57d0")
        self.tabla.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(zona, orient="vertical", command=self.tabla.yview)
        barra.grid(row=0, column=1, sticky="ns")
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.bind("<Double-1>", self._abrir_resultado)
        self.tabla.bind("<Delete>", lambda _e: self._quitar())

        botones = ttk.Frame(zona)
        botones.grid(row=0, column=2, sticky="n", padx=(m // 2, 0))
        self.b_archivos = ttk.Button(botones, text="Añadir archivos…", command=self._elegir_archivos)
        self.b_carpeta = ttk.Button(botones, text="Añadir carpeta…", command=self._elegir_carpeta_origen)
        self.b_quitar = ttk.Button(botones, text="Quitar selección", command=self._quitar)
        self.b_vaciar = ttk.Button(botones, text="Vaciar lista", command=self._vaciar)
        for boton in (self.b_archivos, self.b_carpeta, self.b_quitar, self.b_vaciar):
            boton.pack(fill="x", pady=(0, 6))

        # Qué crear
        tipos = ttk.LabelFrame(self, text="Qué crear", padding=(m, m // 2))
        tipos.grid(row=2, column=0, sticky="ew", pady=(m, 0))
        self.casillas = []
        for ext, etiqueta in self.modo.formatos:
            casilla = ttk.Checkbutton(tipos, text=etiqueta, variable=self.formatos[ext])
            casilla.pack(side="left", padx=(0, 18))
            self.casillas.append(casilla)

        # Dónde guardar
        opciones = ttk.LabelFrame(self, text="Dónde guardar", padding=(m, m // 2))
        opciones.grid(row=3, column=0, sticky="ew", pady=(m // 2, 0))
        opciones.columnconfigure(1, weight=1)
        ttk.Radiobutton(
            opciones, text="Junto a cada archivo original", value="junto",
            variable=self.destino, command=self._actualizar_controles,
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(
            opciones, text="En esta carpeta:", value="carpeta",
            variable=self.destino, command=self._actualizar_controles,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.e_carpeta = ttk.Entry(opciones, textvariable=self.carpeta)
        self.e_carpeta.grid(row=1, column=1, sticky="ew", padx=6, pady=(4, 0))
        self.b_destino = ttk.Button(opciones, text="Elegir…", command=self._elegir_destino)
        self.b_destino.grid(row=1, column=2, pady=(4, 0))
        ttk.Checkbutton(
            opciones, text="Reemplazar si ya existe un archivo con el mismo nombre",
            variable=self.sobrescribir,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # Pie: progreso y acciones
        pie = ttk.Frame(self)
        pie.grid(row=4, column=0, sticky="ew", pady=(m, 0))
        pie.columnconfigure(0, weight=1)
        self.progreso = ttk.Progressbar(pie, mode="determinate")
        self.progreso.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        ttk.Label(pie, textvariable=self.estado, style="Sub.TLabel").grid(row=1, column=0, sticky="w")
        self.b_abrir = ttk.Button(pie, text="Abrir carpeta de resultados", command=self._abrir_carpeta)
        self.b_abrir.grid(row=1, column=1, padx=(0, 8))
        self.b_convertir = ttk.Button(
            pie, text=self.modo.boton, style="Principal.TButton", command=self._convertir)
        self.b_convertir.grid(row=1, column=2)

    # ---- lista de archivos ------------------------------------------------

    def _elegir_archivos(self):
        nombre, extensiones = self.modo.filtro
        rutas = filedialog.askopenfilenames(
            parent=self, title="Elige uno o varios archivos",
            filetypes=[(nombre, " ".join("*" + x for x in extensiones)), ("Todos los archivos", "*.*")],
        )
        self.agregar(Path(r) for r in rutas)

    def _elegir_carpeta_origen(self):
        carpeta = filedialog.askdirectory(parent=self, title="Carpeta con los archivos a convertir")
        if not carpeta:
            return
        extensiones = self.modo.filtro[1]
        encontrados = sorted(
            p for p in Path(carpeta).iterdir() if p.is_file() and p.suffix.lower() in extensiones)
        if not encontrados:
            messagebox.showinfo(TITULO, "En esa carpeta no hay archivos compatibles.", parent=self)
            return
        self.agregar(encontrados)

    def agregar(self, rutas):
        ya = {str(p).lower() for p in self.origenes.values()}
        for ruta in rutas:
            ruta = Path(ruta)
            if not ruta.is_file() or str(ruta).lower() in ya:
                continue
            fila = self.tabla.insert("", "end", values=(ruta.name, str(ruta.parent), PENDIENTE))
            self.origenes[fila] = ruta
            ya.add(str(ruta).lower())
        if self.origenes:
            self.estado.set(f"{len(self.origenes)} archivo(s) en la lista.")
        self._actualizar_controles()

    def _quitar(self):
        if self.trabajando:
            return
        for fila in self.tabla.selection():
            self.tabla.delete(fila)
            self.origenes.pop(fila, None)
            self.hechos.pop(fila, None)
        self.estado.set(f"{len(self.origenes)} archivo(s) en la lista." if self.origenes
                        else "Añade archivos para empezar.")
        self._actualizar_controles()

    def _vaciar(self):
        if self.trabajando:
            return
        self.tabla.delete(*self.tabla.get_children())
        self.origenes.clear()
        self.hechos.clear()
        self.progreso.configure(value=0)
        self.estado.set("Añade archivos para empezar.")
        self._actualizar_controles()

    # ---- destino ----------------------------------------------------------

    def _elegir_destino(self):
        carpeta = filedialog.askdirectory(
            parent=self, title="Carpeta donde guardar los resultados",
            initialdir=self.carpeta.get() or None,
        )
        if carpeta:
            self.carpeta.set(str(Path(carpeta)))
            self.destino.set("carpeta")
            self._actualizar_controles()

    def _actualizar_controles(self):
        libre = "disabled" if self.trabajando else "normal"
        for control in (self.b_archivos, self.b_carpeta, self.b_quitar, self.b_vaciar,
                        self.b_destino, *self.casillas):
            control.configure(state=libre)
        en_carpeta = self.destino.get() == "carpeta" and not self.trabajando
        self.e_carpeta.configure(state="normal" if en_carpeta else "disabled")
        self.b_convertir.configure(
            state="normal" if (self.origenes and not self.trabajando) else "disabled")
        self.b_abrir.configure(state="normal" if self.ultima_carpeta else "disabled")

    # ---- conversión -------------------------------------------------------

    def _convertir(self):
        if self.trabajando:
            return
        elegidos = [ext for ext, _ in self.modo.formatos if self.formatos[ext].get()]
        if not elegidos:
            messagebox.showwarning(TITULO, "Marca al menos un tipo de documento en «Qué crear».", parent=self)
            return
        tareas = []
        for fila in self.tabla.get_children():
            faltan = [ext for ext in elegidos if ext not in self.hechos.get(fila, {})]
            if faltan:
                tareas.append((fila, self.origenes[fila], faltan))
        if not tareas:
            messagebox.showinfo(
                TITULO, "Todos los archivos de la lista ya están convertidos.\n"
                        "Añade más archivos, marca otro tipo de documento o vacía la lista.",
                parent=self)
            return

        carpeta = None
        if self.destino.get() == "carpeta":
            texto = self.carpeta.get().strip().strip('"')
            if not texto:
                messagebox.showwarning(TITULO, "Elige la carpeta donde guardar los documentos.", parent=self)
                return
            carpeta = Path(texto)
            if not carpeta.exists() and not carpeta.parent.exists():
                messagebox.showwarning(TITULO, f"La carpeta no existe:\n{carpeta}", parent=self)
                return

        self.winfo_toplevel().guardar_config()
        self.trabajando = True
        self._creados = self._con_problemas = 0
        self._fallo_al_guardar = False
        self.progreso.configure(maximum=len(tareas), value=0)
        self.estado.set("Preparando el conversor…")
        self._actualizar_controles()
        threading.Thread(
            target=self._trabajo, args=(tareas, carpeta, self.sobrescribir.get()), daemon=True,
        ).start()

    def _trabajo(self, tareas, carpeta, sobrescribir):
        """Se ejecuta en un hilo aparte: solo habla con la interfaz por la cola."""
        try:
            self.modo.preparar()
        except Exception as error:
            self.cola.put(("fatal", f"{type(error).__name__}: {error}"))
            return

        usadas = set()
        for numero, (fila, origen, extensiones) in enumerate(tareas, start=1):
            self.cola.put(("proceso", fila, f"Convirtiendo {numero} de {len(tareas)}: {origen.name}"))
            creados, fallos, sin_permiso = {}, [], False
            try:
                contenido = self.modo.leer(origen)
            except Exception as error:
                fallos.append(explicar_error(error))
            else:
                destino_dir = carpeta or origen.parent
                for ext in extensiones:
                    try:
                        try:
                            if not destino_dir.exists():
                                destino_dir.mkdir()
                            destino = ruta_de_salida(origen, destino_dir, ext, sobrescribir, usadas)
                            self.modo.escribir(contenido, origen, destino, ext)
                        except OSError as error:
                            raise ErrorAlGuardar(str(error)) from error
                    except Exception as error:
                        sin_permiso = sin_permiso or isinstance(error, ErrorAlGuardar)
                        fallos.append(f"{ext}: {explicar_error(error)}")
                    else:
                        creados[ext] = destino
            self.cola.put(("resultado", fila, creados, fallos, sin_permiso))
            self.cola.put(("avance", numero))
        self.cola.put(("fin",))

    def _revisar_cola(self):
        try:
            while True:
                mensaje = self.cola.get_nowait()
                tipo = mensaje[0]
                if tipo == "proceso":
                    self._pintar(mensaje[1], "Convirtiendo…", "proceso")
                    self.estado.set(mensaje[2])
                elif tipo == "resultado":
                    self._anotar_resultado(*mensaje[1:])
                elif tipo == "avance":
                    self.progreso.configure(value=mensaje[1])
                elif tipo == "fin":
                    self._terminar()
                elif tipo == "fatal":
                    self.trabajando = False
                    self.estado.set("No se pudo cargar el conversor.")
                    self._actualizar_controles()
                    messagebox.showerror(
                        TITULO,
                        "No se pudo cargar el conversor en este Python.\n\n"
                        f"{mensaje[1]}\n\nPython usado:\n{sys.executable}",
                        parent=self,
                    )
        except queue.Empty:
            pass
        self.after(80, self._revisar_cola)

    def _pintar(self, fila, texto, etiqueta):
        if self.tabla.exists(fila):
            self.tabla.set(fila, "estado", texto)
            self.tabla.item(fila, tags=(etiqueta,))
            self.tabla.see(fila)

    def _anotar_resultado(self, fila, creados, fallos, sin_permiso):
        if fila not in self.origenes:      # la fila ya no existe
            return
        hechos = self.hechos.setdefault(fila, {})
        hechos.update(creados)
        self._creados += len(creados)
        self._fallo_al_guardar = self._fallo_al_guardar or sin_permiso
        if creados:
            self.ultima_carpeta = next(iter(creados.values())).parent
        nombres = ", ".join(ruta.name for ruta in hechos.values())
        if fallos:
            self._con_problemas += 1
            motivo = "; ".join(fallos)
            if hechos:
                self._pintar(fila, f"Creado: {nombres} · Falló {motivo}", "aviso")
            else:
                self.hechos.pop(fila, None)
                etiqueta = "aviso" if "sin texto" in motivo else "error"
                self._pintar(fila, motivo[:1].upper() + motivo[1:], etiqueta)
        else:
            self._pintar(fila, f"Listo → {nombres}", "ok")

    def _terminar(self):
        self.trabajando = False
        partes = [f"{self._creados} documento(s) creado(s)"]
        if self._con_problemas:
            partes.append(f"{self._con_problemas} archivo(s) con problemas (mira la columna Estado)")
        self.estado.set("Terminado: " + ", ".join(partes) + ".")
        self._actualizar_controles()
        if self._fallo_al_guardar:
            messagebox.showwarning(
                TITULO,
                "Windows no dejó guardar algún documento. Suele ser porque el archivo de "
                "destino está abierto en otro programa, o porque la carpeta está protegida "
                "(pasa en Documentos con la protección contra ransomware).\n\n"
                "Cierra el archivo o elige otra carpeta en «En esta carpeta» y pulsa "
                "Convertir de nuevo: solo se repetirá lo que falló.",
                parent=self,
            )

    # ---- abrir resultados -------------------------------------------------

    def _abrir_resultado(self, _evento=None):
        for destino in self.hechos.get(self.tabla.focus(), {}).values():
            if destino.exists():
                abrir(destino)
                return

    def _abrir_carpeta(self):
        if self.ultima_carpeta and self.ultima_carpeta.exists():
            abrir(self.ultima_carpeta)

    def config(self):
        return {
            "destino": self.destino.get(),
            "carpeta": self.carpeta.get(),
            "formatos": [ext for ext, var in self.formatos.items() if var.get()],
            "sobrescribir": self.sobrescribir.get(),
        }


# --------------------------------------------------------------------------
# Ventana principal
# --------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self, archivos_iniciales=()):
        super().__init__()
        self.title(TITULO)
        self.escala = self.winfo_fpixels("1i") / 96.0
        self.geometry(f"{int(920 * self.escala)}x{int(680 * self.escala)}")
        self.minsize(int(760 * self.escala), int(560 * self.escala))
        self.report_callback_exception = self._error_inesperado
        self._estilos()

        config = self._leer_config()
        m = int(14 * self.escala)
        cabecera = ttk.Frame(self, padding=(m, m, m, m // 2))
        cabecera.pack(fill="x")
        ttk.Label(cabecera, text=TITULO, style="Titulo.TLabel").pack(side="left")
        ttk.Label(cabecera, text=versiones(), style="Sub.TLabel").pack(side="right")

        self.pestanas = ttk.Notebook(self)
        self.pestanas.pack(fill="both", expand=True, padx=m // 2, pady=(0, m // 2))
        self.paneles = []
        for modo in (AMarkdown(), DesdeMarkdown()):
            panel = Panel(self.pestanas, modo, config.get(modo.clave, {}), self.escala)
            self.pestanas.add(panel, text=f"  {modo.pestana}  ")
            self.paneles.append(panel)
        try:
            self.pestanas.select(int(config.get("pestana", 0)))
        except (tk.TclError, ValueError, TypeError):
            pass

        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        if archivos_iniciales:
            self._repartir(Path(a) for a in archivos_iniciales)

    def _estilos(self):
        estilo = ttk.Style(self)
        if "vista" in estilo.theme_names():
            estilo.theme_use("vista")
        estilo.configure("Treeview", rowheight=int(26 * self.escala))
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 15, "bold"))
        estilo.configure("Sub.TLabel", foreground="#5b6470")
        estilo.configure("Principal.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 8))
        estilo.configure("TNotebook.Tab", padding=(10, 6))

    def _repartir(self, rutas):
        """Archivos recibidos al abrir (p. ej. soltados sobre el acceso directo):
        los .md van a la pestaña inversa; el resto, a la de convertir a Markdown."""
        a_md, desde_md = [], []
        for ruta in rutas:
            (desde_md if ruta.suffix.lower() in (".md", ".markdown") else a_md).append(ruta)
        self.paneles[0].agregar(a_md)
        self.paneles[1].agregar(desde_md)
        if desde_md and not a_md:
            self.pestanas.select(1)
        elif a_md:
            self.pestanas.select(0)

    # ---- configuración y cierre -------------------------------------------

    def _leer_config(self):
        try:
            datos = json.loads(ARCHIVO_CONFIG.read_text(encoding="utf-8"))
            return datos if isinstance(datos, dict) else {}
        except (OSError, ValueError):
            return {}

    def guardar_config(self):
        datos = {panel.modo.clave: panel.config() for panel in self.paneles}
        try:
            datos["pestana"] = self.pestanas.index(self.pestanas.select())
            ARCHIVO_CONFIG.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
        except (OSError, tk.TclError):
            pass  # recordar las opciones es opcional

    def _cerrar(self):
        if any(p.trabajando for p in self.paneles) and not messagebox.askyesno(
            TITULO, "Hay una conversión en curso. ¿Cerrar de todos modos?", parent=self
        ):
            return
        self.guardar_config()
        self.destroy()

    def _error_inesperado(self, tipo, valor, _traza):
        messagebox.showerror(TITULO, f"Error inesperado:\n{tipo.__name__}: {valor}", parent=self)


# --------------------------------------------------------------------------
# Arranque: funcionar con cualquier Python instalado
# --------------------------------------------------------------------------
# La app no depende de una versión concreta de Python. Al abrirse:
#   1. Si este Python tiene todos los componentes, se usa.
#   2. Si no, se busca otro Python instalado que sí los tenga y se reabre con él.
#   3. Si ninguno los tiene (p. ej. tras actualizar Python), ofrece instalarlos.

MARCA_RELANZADO = "CONVERTIDOR_MD_RELANZADO"
# (módulo que se importa, paquete que se instala con pip)
REQUISITOS = (
    ("markitdown", "markitdown[all]"),
    ("markdown_it", "markdown-it-py"),
    ("docx", "python-docx"),
    ("bs4", "beautifulsoup4"),
)


def faltantes():
    """Paquetes de pip que le faltan a este Python."""
    importlib.invalidate_caches()
    return [paquete for modulo, paquete in REQUISITOS if importlib.util.find_spec(modulo) is None]


def otros_python():
    """Los demás Python instalados, según el lanzador `py` de Windows."""
    try:
        salida = subprocess.run(
            ["py", "-0p"], capture_output=True, text=True, timeout=20, creationflags=SIN_VENTANA,
        ).stdout
    except (OSError, subprocess.SubprocessError):
        return []
    actual = Path(sys.executable).resolve().parent
    rutas = []
    for linea in salida.splitlines():
        hallado = re.search(r"[A-Za-z]:\\.*\.exe\s*$", linea)
        if not hallado:
            continue
        ruta = Path(hallado.group(0).strip())
        if ruta.exists() and ruta.resolve().parent != actual and ruta not in rutas:
            rutas.append(ruta)
    return rutas


def python_completo():
    """Otro Python instalado que ya tenga todos los componentes, o None."""
    modulos = [modulo for modulo, _ in REQUISITOS]
    prueba = ("import importlib.util as u, sys; "
              f"sys.exit(0 if all(u.find_spec(m) for m in {modulos!r}) else 1)")
    for ruta in otros_python():
        try:
            listo = subprocess.run(
                [str(ruta), "-c", prueba], timeout=60, creationflags=SIN_VENTANA,
            ).returncode == 0
        except (OSError, subprocess.SubprocessError):
            continue
        if listo:
            return ruta
    return None


def instalar(paquetes):
    """Instala paquetes en este Python mostrando el avance en una consola."""
    python = Path(sys.executable).with_name("python.exe")
    if not python.exists():
        python = Path(sys.executable)
    try:
        subprocess.run(
            [str(python), "-m", "pip", "install", *paquetes],
            creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
        )
    except OSError:
        return False
    return not faltantes()


def preparar_python():
    """Devuelve True si se puede seguir en este Python; False si hay que salir."""
    faltan = faltantes()
    if not faltan:
        return True

    if not os.environ.get(MARCA_RELANZADO):
        otro = python_completo()
        if otro:
            sin_consola = otro.with_name("pythonw.exe")
            ejecutable = sin_consola if sin_consola.exists() else otro
            subprocess.Popen(
                [str(ejecutable), str(Path(__file__).resolve()), *sys.argv[1:]],
                env={**os.environ, MARCA_RELANZADO: "1"},
            )
            return False

    version = f"{sys.version_info.major}.{sys.version_info.minor}"
    raiz = tk.Tk()
    raiz.withdraw()
    try:
        aceptar = messagebox.askyesno(
            TITULO,
            f"A Python {version} le faltan componentes para convertir "
            "(suele pasar después de actualizar Python):\n\n  "
            + "\n  ".join(faltan)
            + "\n\n¿Instalarlos ahora? Hace falta internet y tarda un par de minutos.",
            parent=raiz,
        )
        if not aceptar:
            return False
        if instalar(faltan):
            return True
        messagebox.showerror(
            TITULO,
            "No se pudieron instalar los componentes.\n\n"
            "Puedes intentarlo a mano en una terminal:\n"
            "py -m pip install " + " ".join(f'"{p}"' for p in faltan),
            parent=raiz,
        )
        return False
    finally:
        raiz.destroy()


def versiones():
    """Texto con las versiones en uso, para mostrarlo en la ventana."""
    try:
        from importlib.metadata import version
        motor = version("markitdown")
    except Exception:
        motor = "?"
    return f"Python {sys.version_info.major}.{sys.version_info.minor} · MarkItDown {motor}"


def main():
    try:  # texto nítido en pantallas con escala >100 %
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    if not preparar_python():
        return
    App(sys.argv[1:]).mainloop()


if __name__ == "__main__":
    main()
