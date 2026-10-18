#!/usr/bin/env python3
"""Erzeugt assets/docubuilder.ico und assets/docubuilder.png aus den SVG-Vorlagen.

Nur nötig, wenn sich die SVGs ändern; die erzeugten Dateien liegen im Repo.
Benötigt: python -m pip install resvg-py pillow

Kleine Größen (bis 32 px) nutzen docubuilder-small.svg (gröbere Linien, besser lesbar),
größere die detaillierte docubuilder.svg.
"""
import io
import os

import resvg_py
from PIL import Image

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
SMALL = os.path.join(ASSETS, "docubuilder-small.svg")
LARGE = os.path.join(ASSETS, "docubuilder.svg")
ICO_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)


def render(svg, size):
    png = resvg_py.svg_to_bytes(svg_path=svg, width=size, height=size)
    return Image.open(io.BytesIO(bytes(png))).convert("RGBA")


def main():
    images = [render(SMALL if s <= 32 else LARGE, s) for s in ICO_SIZES]
    largest = images[-1]
    largest.save(os.path.join(ASSETS, "docubuilder.ico"), format="ICO",
                 sizes=[(s, s) for s in ICO_SIZES], append_images=images[:-1])
    largest.save(os.path.join(ASSETS, "docubuilder.png"))   # Fenstersymbol (Linux) und Vorschau
    render(SMALL, 32).save(os.path.join(ASSETS, "docubuilder-32.png"))
    print("Erzeugt:", ", ".join(f"{s}px" for s in ICO_SIZES))


if __name__ == "__main__":
    main()
