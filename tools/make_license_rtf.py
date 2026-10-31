#!/usr/bin/env python3
"""Wandelt LICENSE in build/License.rtf um (Lizenzseite des MSI-Installationsprogramms)."""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def rtf_escape(text):
    out = []
    for ch in text:
        if ch in "\\{}":
            out.append("\\" + ch)
        elif ch == "\n":
            out.append("\\par\n")
        elif ord(ch) > 127:
            code = ord(ch)
            out.append("\\u%d?" % (code if code < 32768 else code - 65536))
        else:
            out.append(ch)
    return "".join(out)


def main():
    with open(os.path.join(ROOT, "LICENSE"), encoding="utf-8") as f:
        text = f.read().replace("\r\n", "\n")
    os.makedirs(os.path.join(ROOT, "build"), exist_ok=True)
    target = os.path.join(ROOT, "build", "License.rtf")
    with open(target, "w", encoding="ascii", newline="") as f:
        f.write("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Segoe UI;}}\\fs18\n" + rtf_escape(text) + "}")
    print(target)


if __name__ == "__main__":
    main()
