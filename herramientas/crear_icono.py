# -*- coding: utf-8 -*-
"""Genera el icono de la app (icono.ico e icono.png): «md» con una flecha.

    py herramientas/crear_icono.py

Solo hace falta volver a ejecutarlo si se quiere cambiar el diseño.
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

RAIZ = Path(__file__).resolve().parent.parent
LADO = 1024                      # se dibuja grande y se reduce para que quede nítido
FONDO_ARRIBA = (38, 118, 245)
FONDO_ABAJO = (10, 62, 170)
BLANCO = (255, 255, 255)
ACENTO = (255, 200, 61)


def fuente(tamano):
    for nombre in ("segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf"):
        try:
            return ImageFont.truetype(nombre, tamano)
        except OSError:
            continue
    return ImageFont.load_default(tamano)


def dibujar():
    # Fondo: cuadrado redondeado con degradado vertical
    degradado = Image.new("RGB", (LADO, LADO))
    trazo = ImageDraw.Draw(degradado)
    for y in range(LADO):
        t = y / (LADO - 1)
        color = tuple(round(a + (b - a) * t) for a, b in zip(FONDO_ARRIBA, FONDO_ABAJO))
        trazo.line([(0, y), (LADO, y)], fill=color)
    mascara = Image.new("L", (LADO, LADO), 0)
    ImageDraw.Draw(mascara).rounded_rectangle([24, 24, LADO - 24, LADO - 24], radius=210, fill=255)
    icono = Image.new("RGBA", (LADO, LADO), (0, 0, 0, 0))
    icono.paste(degradado, (0, 0), mascara)
    trazo = ImageDraw.Draw(icono)

    # «md»
    letra = fuente(520)
    caja = trazo.textbbox((0, 0), "md", font=letra)
    ancho, alto = caja[2] - caja[0], caja[3] - caja[1]
    x = (LADO - ancho) / 2 - caja[0]
    y = 395 - alto / 2 - caja[1]
    trazo.text((x, y), "md", font=letra, fill=BLANCO)

    # Flecha hacia la derecha
    centro_y, grosor = 790, 78
    izquierda, derecha, punta = 250, 640, 800
    trazo.rounded_rectangle(
        [izquierda, centro_y - grosor / 2, derecha + 20, centro_y + grosor / 2],
        radius=grosor / 2, fill=ACENTO)
    trazo.polygon(
        [(derecha - 10, centro_y - 150), (punta, centro_y), (derecha - 10, centro_y + 150)],
        fill=ACENTO)
    return icono


if __name__ == "__main__":
    grande = dibujar()
    grande.resize((256, 256), Image.LANCZOS).save(RAIZ / "icono.png")
    grande.resize((256, 256), Image.LANCZOS).save(
        RAIZ / "icono.ico",
        sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    print("Creados:", RAIZ / "icono.ico", "y", RAIZ / "icono.png")
