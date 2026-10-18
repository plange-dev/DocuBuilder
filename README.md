# DocuBuilder

Kleines Desktop-Tool für Windows und Linux: Eine Auswahl an Buttons, jeder Button öffnet eine hinterlegte Excel-Liste. Einträge lassen sich im Einstellungsfenster beliebig hinzufügen, ändern, löschen und sortieren.

![DocuBuilder in 40 Sekunden: Buttons, Einstellungen, Hell/Dunkel, Default und Schutz der Einstellungen](docs/DocuBuilder-Erklaerung.gif)

*Kurzvorstellung (40 Sekunden). Die Animation ist eine nachgebaute Darstellung der Oberfläche mit Beispieldaten, keine Bildschirmaufnahme.*

## Funktionen
- Buttons öffnen Excel-Dateien mit dem Standardprogramm (Windows: Excel, Linux: `xdg-open`)
- Unterstützt lokale Pfade, Netzpfade (`\\Server\Freigabe\datei.xlsx`) und Web-Links
- Einstellungen mit drei Reitern: **Einträge**, **Allgemein** (Ansicht hell/dunkel/System) und **Info**
- **Default-Fenster**: beim ersten Start legst du die Anzahl der Buttons und den Ordner mit deinen Dateien fest (Excel-Dateien aus dem Ordner können direkt als Buttons übernommen werden). Danach öffnet es sich nur noch mit **ALT+F11** (alternativ **STRG+UMSCHALT+F11**)
- Einstellungen liegen **verschlüsselt, schreibgeschützt und versteckt** im Benutzerprofil (siehe Abschnitt „Schutz der Einstellungen“)
- Keine Abhängigkeiten außer Python 3 (tkinter)

## Erste Einrichtung (Default)
Beim ersten Start öffnet sich das Fenster **Default**:
1. Anzahl der Buttons festlegen (1 bis 200)
2. Ordner mit den Dateien wählen
3. Optional Excel-Dateien aus dem Ordner als Buttons übernehmen. Freie Plätze heißen „Auswahl 1, 2, …“ und lassen sich unter Einstellungen → Einträge belegen.

Optional schützt du das Fenster mit einem **Passwort**. Später erreichst du es nur mit **ALT+F11** oder alternativ **STRG+UMSCHALT+F11** (falls dein System ALT+F11 selbst belegt). Das ist ein Komfortschutz gegen versehentliches Ändern, keine Sicherheitsfunktion: Wer die `config.json` bearbeiten kann, kann die Werte ebenfalls ändern. Relative Pfade in den Einträgen werden ab dem gewählten Ordner gesucht.

## Starten (ohne Build)
```bash
python docubuilder.py        # Windows
python3 docubuilder.py       # Linux (benötigt python3-tk)
```

## Fertige Programme herunterladen
Unter **Releases** liegen `DocuBuilder.exe` (Windows) und `DocuBuilder` (Linux).

## Selbst bauen
- Windows: `build_windows.bat` doppelklicken, Ergebnis in `dist\`
- Windows signieren: `sign_windows.bat <Fingerprint>` (siehe CODE_SIGNING.md)
- Linux: `./build_linux.sh`, Ergebnis in `dist/`

## Automatischer Build über GitHub
Der Workflow `.github/workflows/build.yml` baut bei jedem Push eines Tags `v*` (z. B. `v1.0.0`) beide Versionen und hängt sie an ein Release.
Zum manuellen Test: Reiter **Actions** → **Build** → **Run workflow**.

## Code Signing
Die Windows-Version wird mit einem eigenen Code-Signing-Zertifikat signiert (`sign_windows.bat`). Ablauf und Prüfung siehe [CODE_SIGNING.md](CODE_SIGNING.md).

## Schutz der Einstellungen
Die Einstellungen werden nicht mehr als lesbare `config.json` gespeichert, sondern in der Datei `.docubuilder.dat`:
- **Ort:** Windows `%APPDATA%\DocuBuilder\` (versteckter Ordner), Linux `~/.config/docubuilder/` (Rechte nur für den Benutzer)
- **Datei:** verschlüsselt, mit Manipulationserkennung, schreibgeschützt und unter Windows versteckt. In einem Texteditor ist nichts lesbar; veränderte Dateien werden erkannt und zurückgesetzt.
- **Passwort (optional):** im Default-Fenster setzbar. Es wird nur als Hash gespeichert und schützt das Öffnen des Default-Fensters.
- **Portabel:** Liegt eine leere Datei `portable.txt` neben dem Programm, wird die Datei stattdessen im Programmordner abgelegt.
- **Alte Versionen:** Eine vorhandene `config.json` wird beim ersten Start übernommen und danach gelöscht.

**Grenzen (bitte ehrlich einordnen):** Das ist ein Schutz gegen versehentliches oder neugieriges Ansehen und Ändern durch normale Benutzer, kein Schutz gegen Administratoren oder Fachleute. Der Schlüsselanteil steckt im Programm, und wer das Benutzerprofil löscht oder die Datei entfernt, setzt die Einstellungen (inklusive Passwort) zurück. Für echten Zugriffsschutz braucht es Betriebssystem-Rechte, z. B. ein getrenntes Benutzerkonto ohne Schreibrechte auf den Programmordner.

## Windows-Hinweis: SmartScreen
Beim ersten Start kann Windows „Der Computer wurde durch Windows geschützt" anzeigen. Das ist bei neuen Programmen normal, auch bei signierten: SmartScreen vertraut Dateien erst nach einer gewissen Zahl an Downloads und Starts.

So startest du DocuBuilder trotzdem:
1. Auf **Weitere Informationen** klicken
2. Auf **Trotzdem ausführen** klicken

Unter Rechtsklick → Eigenschaften → Digitale Signaturen siehst du den geprüften Herausgeber. Die Prüfsumme (`.sha256`) im Release kannst du mit `certutil -hashfile DocuBuilder.exe SHA256` vergleichen.

Auf Rechnern, die du selbst verwaltest, lässt sich die Meldung vorab entfernen: Rechtsklick auf die .exe → Eigenschaften → Haken bei **Zulassen** → OK (oder in PowerShell: `Unblock-File .\DocuBuilder.exe`).

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst.

## Lizenz
MIT
