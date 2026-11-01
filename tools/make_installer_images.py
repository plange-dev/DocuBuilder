#!/usr/bin/env python3
"""Erzeugt die Bilder des MSI-Installationsprogramms im Design von DocuBuilder:

  assets/installer-banner.bmp   493x58   Kopfzeile der Dialoge
  assets/installer-dialog.bmp   493x312  Willkommens- und Abschlussseite (links die farbige Leiste)

Nur nötig, wenn sich Logo oder Farben ändern; die erzeugten Dateien liegen im Repo.
Benötigt: python -m pip install pillow
Farben wie das Icon: Blau #3654E8, Dunkelblau #1F2A44, Orange #FFB020.
"""
import os

from PIL import Image, ImageDraw

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
BLUE, NAVY, ORANGE = (0x36, 0x54, 0xE8), (0x1F, 0x2A, 0x44), (0xFF, 0xB0, 0x20)
LIGHT, WHITE = (0xF4, 0xF6, 0xFB), (0xFF, 0xFF, 0xFF)
SS = 4   # Überabtastung für glatte Kanten


def logo(size):
    img = Image.open(os.path.join(ASSETS, "docubuilder.png")).convert("RGBA")
    return img.resize((size, size), Image.LANCZOS)


def paste(base, img, xy):
    base.paste(img, xy, img)


def gradient(w, h, top, bottom):
    img = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / max(1, h - 1)
        d.line([(0, y), (w, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)))
    return img


def banner():
    w, h = 493, 58
    img = Image.new("RGB", (w, h), WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle([0, h - 3, w, h], fill=BLUE)          # Farbstreifen unten
    d.rectangle([0, h - 3, 120, h], fill=ORANGE)      # Akzent wie im Icon
    paste(img, logo(44), (w - 44 - 14, 5))            # Logo rechts, Text steht links
    return img


def dialog():
    w, h, side = 493, 312, 164
    img = Image.new("RGB", (w, h), WHITE)
    img.paste(gradient(side, h, NAVY, BLUE), (0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle([side, 0, side + 3, h], fill=ORANGE)  # Akzentkante zwischen Leiste und Text
    # Logo mit hellem Kreis dahinter, damit es sich vom Verlauf abhebt
    big = Image.new("RGBA", (side * SS, 130 * SS), (0, 0, 0, 0))
    bd = ImageDraw.Draw(big)
    c = side * SS // 2
    bd.ellipse([c - 62 * SS, 65 * SS - 62 * SS, c + 62 * SS, 65 * SS + 62 * SS],
               fill=(255, 255, 255, 38))
    circle = big.resize((side, 130), Image.LANCZOS)
    paste(img, circle, (0, 40))
    paste(img, logo(96), ((side - 96) // 2, 40 + 65 - 48))
    # drei Dokumentzeilen als dezentes Muster am unteren Rand der Leiste
    for i, ln in enumerate((96, 72, 48)):
        y = 236 + i * 20
        d.rounded_rectangle([34, y, 34 + ln, y + 8], radius=4, fill=(255, 255, 255))
        if i == 0:
            d.rounded_rectangle([34, y, 34 + ln, y + 8], radius=4, fill=ORANGE)
    return img


def main():
    banner().save(os.path.join(ASSETS, "installer-banner.bmp"), format="BMP")
    dialog().save(os.path.join(ASSETS, "installer-dialog.bmp"), format="BMP")
    print("assets/installer-banner.bmp, assets/installer-dialog.bmp")


if __name__ == "__main__":
    main()
