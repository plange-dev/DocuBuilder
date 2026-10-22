#!/usr/bin/env bash
# Installiert DocuBuilder unter Linux für alle Benutzer (mit sudo ausführen).
#
#   sudo ./install.sh                                   Einzelplatz (Einstellungen in /etc/docubuilder)
#   sudo ./install.sh --server /srv/docubuilder         Server: gemeinsamer Datenordner auf diesem Rechner
#   sudo ./install.sh --arbeitsplatz /mnt/docubuilder   Arbeitsplatz: Datenordner des Servers (eingehängt)
#
# Weitere Optionen:
#   --map '\\SERVER\Vorlagen=/mnt/vorlagen'   Pfade der Einträge übersetzen (mehrfach möglich), wenn die
#                                             Einträge mit Windows-Pfaden angelegt wurden
#   --desktop                                 Verknüpfung auf den Desktop aller vorhandenen Benutzer legen
#   --deinstallieren                          DocuBuilder entfernen (gemeinsame Daten bleiben erhalten)
set -euo pipefail

PREFIX=/opt/docubuilder
SRC="$(cd "$(dirname "$0")" && pwd)"
MODE=einzel
FOLDER=""
MAPS=()
DESKTOP=0

usage() { sed -n '2,15p' "$0" | sed 's/^# \{0,1\}//'; exit 1; }

while [ $# -gt 0 ]; do
  case "$1" in
    --server) MODE=server; FOLDER="${2:-}"; shift 2 || usage ;;
    --arbeitsplatz) MODE=arbeitsplatz; FOLDER="${2:-}"; shift 2 || usage ;;
    --map) MAPS+=("${2:-}"); shift 2 || usage ;;
    --desktop) DESKTOP=1; shift ;;
    --deinstallieren) MODE=deinstallieren; shift ;;
    -h|--help) usage ;;
    *) echo "Unbekannte Option: $1"; usage ;;
  esac
done

if [ "$(id -u)" -ne 0 ]; then
  echo "Bitte mit sudo ausführen: sudo $0 $*"
  exit 1
fi

desktop_dirs() {
  # Desktop-Ordner aller Benutzer mit Home unter /home (Desktop oder Schreibtisch)
  for home in /home/*; do
    [ -d "$home" ] || continue
    for d in "$home/Desktop" "$home/Schreibtisch"; do
      [ -d "$d" ] && echo "$d"
    done
  done
}

if [ "$MODE" = deinstallieren ]; then
  rm -rf "$PREFIX"
  rm -f /usr/local/bin/docubuilder /usr/share/applications/docubuilder.desktop
  while read -r d; do rm -f "$d/docubuilder.desktop"; done < <(desktop_dirs)
  echo "DocuBuilder wurde entfernt. Gemeinsame Daten (z. B. /etc/docubuilder) bleiben erhalten."
  exit 0
fi

if [ ! -f "$SRC/DocuBuilder" ]; then
  echo "Die Programmdatei DocuBuilder fehlt neben install.sh."
  exit 1
fi
if [ "$MODE" != einzel ] && [ -z "$FOLDER" ]; then
  echo "Bitte den gemeinsamen Ordner angeben."
  usage
fi

install -d -m 755 "$PREFIX"
install -m 755 "$SRC/DocuBuilder" "$PREFIX/DocuBuilder"
[ -f "$SRC/docubuilder.png" ] && install -m 644 "$SRC/docubuilder.png" "$PREFIX/docubuilder.png"
for f in README.md LICENSE; do
  [ -f "$SRC/$f" ] && install -m 644 "$SRC/$f" "$PREFIX/$f"
done
ln -sf "$PREFIX/DocuBuilder" /usr/local/bin/docubuilder

# server.txt: gemeinsamer Ordner und Pfadübersetzungen (liest DocuBuilder beim Start)
rm -f "$PREFIX/server.txt"
case "$MODE" in
  einzel)
    # Gemeinsamer Ordner für alle Benutzer dieses Rechners; Sticky-Bit: nur der Eigentümer darf löschen
    install -d -m 1777 /etc/docubuilder
    ;;
  server|arbeitsplatz)
    {
      echo "# Gemeinsamer Ordner von DocuBuilder (von install.sh angelegt)"
      echo "$FOLDER"
      for m in "${MAPS[@]+"${MAPS[@]}"}"; do echo "$m"; done
    } > "$PREFIX/server.txt"
    chmod 644 "$PREFIX/server.txt"
    if [ "$MODE" = server ]; then
      install -d -m 1777 "$FOLDER"
    elif [ ! -d "$FOLDER" ]; then
      echo "Hinweis: $FOLDER ist im Moment nicht erreichbar. Bitte die Freigabe des Servers dort"
      echo "einhängen (z. B. per /etc/fstab, siehe README)."
    fi
    ;;
esac

cat > /usr/share/applications/docubuilder.desktop <<EOF
[Desktop Entry]
Type=Application
Name=DocuBuilder
Comment=Buttons, die hinterlegte Excel-Listen öffnen
Exec=$PREFIX/DocuBuilder
Icon=$PREFIX/docubuilder.png
Terminal=false
Categories=Office;
EOF
chmod 644 /usr/share/applications/docubuilder.desktop

if [ "$DESKTOP" = 1 ]; then
  while read -r d; do
    install -m 755 /usr/share/applications/docubuilder.desktop "$d/docubuilder.desktop"
    chown "$(stat -c %U "$(dirname "$d")")": "$d/docubuilder.desktop" 2>/dev/null || true
  done < <(desktop_dirs)
fi

echo "DocuBuilder ist installiert ($MODE): $PREFIX"
[ "$MODE" != einzel ] && echo "Gemeinsamer Ordner: $FOLDER"
echo "Starten über das Anwendungsmenü oder mit: docubuilder"
