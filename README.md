# DocuBuilder

Kleines Desktop-Tool für Windows und Linux: Eine Auswahl an Buttons, jeder Button öffnet eine hinterlegte Excel-Liste. Einträge lassen sich im Einstellungsfenster beliebig hinzufügen, ändern, löschen und sortieren.

## Funktionen
- Buttons öffnen Excel-Dateien mit dem Standardprogramm (Windows: Excel, Linux: `xdg-open`)
- Unterstützt lokale Pfade, Netzpfade (`\\Server\Freigabe\datei.xlsx`) und Web-Links
- Einstellungen mit zwei Reitern: **Einträge** und **Allgemein** (Ansicht hell/dunkel/System, Updates)
- **Menü** oben rechts: Nach Updates suchen, Info zu DocuBuilder, Beenden
- **Anmeldung** beim Start mit Benutzername und Passwort; das Konto **Admin** besteht immer
- **Default-Fenster**: beim ersten Start legst du das Admin-Passwort, die Benutzer, die Anzahl der Buttons und den Ordner mit deinen Dateien fest (Excel-Dateien aus dem Ordner können direkt als Buttons übernommen werden). Danach öffnet es sich nur noch mit **ALT+F11** (alternativ **STRG+UMSCHALT+F11**)
- Einstellungen liegen **verschlüsselt und versteckt**; ändern dürfen sie nur der **Admin und Benutzer mit Recht** (siehe Abschnitt „Benutzer und Rechte“)
- Keine Abhängigkeiten außer Python 3 (tkinter)

## Erste Einrichtung (Default)
Beim ersten Start öffnet sich das Fenster **Default**:
1. Anzahl der Buttons festlegen (1 bis 200)
2. Ordner mit den Dateien wählen
3. Optional Excel-Dateien aus dem Ordner als Buttons übernehmen. Freie Plätze heißen „Auswahl 1, 2, …“ und lassen sich unter Einstellungen → Einträge belegen.
4. **Passwort für das Konto „Admin“** festlegen (Pflicht, mindestens 4 Zeichen)
5. Optional **Benutzer anlegen** (siehe unten)

Danach bist du als Admin angemeldet. Später erreichst du das Fenster mit **ALT+F11** oder alternativ **STRG+UMSCHALT+F11** (falls dein System ALT+F11 selbst belegt). Relative Pfade in den Einträgen werden ab dem gewählten Ordner gesucht.

## Benutzer und Rechte
Beim Start meldet man sich mit Benutzername und Passwort an (Groß-/Kleinschreibung beim Namen egal). Oben rechts steht, wer angemeldet ist; **Abmelden** führt zurück zur Anmeldung.

| Rolle | Buttons benutzen | Einträge ändern | Default-Fenster (Anzahl, Ordner) | Benutzer verwalten |
|---|---|---|---|---|
| **Admin** (fest, nicht löschbar) | ja | ja | ja | ja |
| Benutzer mit **„Darf ändern“** | ja | ja | ja | nein |
| Benutzer ohne Recht | ja | nein | nein | nein |

**Benutzer verwalten** (nur Admin): anmelden als Admin → **ALT+F11** → Bereich **Benutzer**:
- **Hinzufügen …**: Name, Passwort, Häkchen „Darf ändern“
- **Passwort …**: neues Passwort für den markierten Benutzer setzen (z. B. wenn es vergessen wurde)
- **Recht an/aus**: „Darf ändern“ umschalten
- **Löschen**: Benutzer entfernen
- **Admin-Passwort ändern**: in den beiden Feldern oben eintragen (leer = unverändert)

Mit **Übernehmen** wird gespeichert. Das Konto „Admin“ kann weder gelöscht noch eingeschränkt werden.

## DocuBuilder entfernen
Wer DocuBuilder nicht mehr verwenden will: als **Admin** anmelden → **Einstellungen → Allgemein → „DocuBuilder-Daten entfernen …“**. Nach einer Rückfrage und erneuter Eingabe des Admin-Passworts werden gelöscht:
- die gemeinsame Konfiguration `C:\ProgramData\DocuBuilder\` (Einträge, Benutzer, Passwörter). Windows fragt dafür nach Administratorrechten (UAC).
- die eigene Ansicht `%APPDATA%\DocuBuilder\`

Danach beendet sich DocuBuilder, und die Programmdatei (`DocuBuilder.exe`) kann gelöscht werden. Andere Windows-Benutzer können ihre Ansicht unter `%APPDATA%\DocuBuilder` selbst löschen. Im portablen Modus wird nur die Einstellungsdatei gelöscht, nicht der Programmordner.

**Admin-Passwort vergessen?** Dann gibt es im Programm keinen Weg zurück. Ein Windows-Administrator muss die Datei `C:\ProgramData\DocuBuilder\.docubuilder.dat` löschen (versteckte Dateien einblenden). Danach startet DocuBuilder mit der Ersteinrichtung, alle Einträge und Benutzer sind dann weg.

## Starten (ohne Build)
```bash
python docubuilder.py        # Windows
python3 docubuilder.py       # Linux (benötigt python3-tk)
```

## Fertige Programme herunterladen
Unter **Releases** liegen `DocuBuilder.exe` (Windows) und `DocuBuilder` (Linux).

## Installieren (Windows)
`DocuBuilder-Setup-<Version>.exe` aus den **Releases** starten. Das Installationsprogramm
- installiert nach `C:\Program Files\DocuBuilder` (Administratorrechte nötig),
- legt einen Eintrag im Startmenü an (Desktop-Verknüpfung optional),
- richtet den geschützten Einstellungsordner `C:\ProgramData\DocuBuilder` ein; ein Start „als Administrator“ ist danach nicht mehr nötig,
- installiert neuere Versionen einfach über die alte; Einträge und Benutzer bleiben erhalten.

**Deinstallieren:** Windows-Einstellungen → Apps → Installierte Apps → DocuBuilder → Deinstallieren. Dabei wird gefragt, ob auch die Einstellungen (Einträge, Benutzer, Passwörter) gelöscht werden sollen.

Alternativ gibt es weiterhin die einzelne `DocuBuilder.exe` ohne Installation.

## Updates
> **Hinweis:** Die Update-Funktion ist vorerst abgeschaltet (*Menü → Nach Updates suchen* ist ausgegraut). Sie wird freigeschaltet, sobald DocuBuilder signiert ausgeliefert wird, weil Smart App Control unsignierte Updates blockiert. Bis dahin neue Versionen über das Setup aus den [Releases](https://github.com/plange-dev/DocuBuilder/releases) installieren.

**Menü → Nach Updates suchen …** (oder *Einstellungen → Allgemein → Jetzt nach Updates suchen*) fragt die Releases auf GitHub ab:
- Eine **Stable**-Version sucht nur Stable-Releases, eine **Nightly**-Version zusätzlich die Vorabversionen.
- Gibt es eine neuere Version, zeigt das Fenster die Änderungen. Das Update läuft in zwei Schritten (nur für den **Admin**):
  1. **Herunterladen:** lädt das Update und prüft die SHA-256-Prüfsumme. DocuBuilder kann dabei weiter benutzt werden.
  2. **DocuBuilder schließen und installieren:** Nach einer Rückfrage wird DocuBuilder geschlossen; erst dann wird installiert. Mit **Später** bleibt das Update bereit – beim Beenden von DocuBuilder wird noch einmal gefragt, ob es jetzt installiert werden soll.
  - **Installierte Version:** Das Setup wartet, bis DocuBuilder beendet ist, und läuft dann ohne Rückfragen; Windows fragt nach Administratorrechten. Danach startet DocuBuilder neu. Einträge und Benutzer bleiben erhalten.
  - **Einzelne `DocuBuilder.exe`:** Die Datei wird ausgetauscht und neu gestartet. Dafür braucht DocuBuilder Schreibrechte im eigenen Ordner.
- Optional: *Einstellungen → Allgemein → „Beim Anmelden nach Updates suchen“* (standardmäßig aus) meldet sich nur, wenn es ein Update gibt.

## Selbst bauen
- Windows-Installationsprogramm: `build_installer.bat` doppelklicken (benötigt Inno Setup 6: `winget install JRSoftware.InnoSetup`), Ergebnis `dist\DocuBuilder-Setup-<Version>.exe` und die portable `dist\DocuBuilder.exe`, jeweils mit Prüfsumme. Mit gesetztem `DOCUBUILDER_CERT_SHA1` wird alles signiert (siehe [CODE_SIGNING.md](CODE_SIGNING.md)).
- Windows: `build_windows.bat` doppelklicken, Ergebnis in `dist\`
- Icon: Vorlagen sind `assets/docubuilder.svg` (groß) und `assets/docubuilder-small.svg` (bis 32 px). Nach Änderungen `python tools/make_icons.py` ausführen (benötigt `pip install resvg-py pillow`); das erzeugt `assets/docubuilder.ico` und `assets/docubuilder.png`, die im Repo liegen.
- Windows signieren: `sign_windows.bat <Fingerprint>` (siehe CODE_SIGNING.md)
- Linux: `./build_linux.sh`, Ergebnis in `dist/`

## Versionen: Stable und Nightly
Versionsnummern haben vier Stellen: `Major.Minor.Patch.Build`.

| Variante | Branch | Version | Wann |
|---|---|---|---|
| **Stable** | `main` | `1.0.0.0` → `1.1.0.0` → `1.2.0.0` … | jeder Merge (Push) auf `main` |
| **Nightly** | `nightly` | `1.0.0.1` → `1.0.0.2` → `1.0.0.3` … | jede Änderung (Push) auf `nightly` |

**Ablauf:**
1. Neue Änderungen kommen in den Branch **`nightly`**. Jeder Push baut automatisch eine Vorabversion (Release „DocuBuilder Nightly 1.0.0.x“, als *Pre-release* markiert). Die letzte Stelle zählt die Commits auf `nightly` seit dem letzten Stable-Release.
2. Ist alles getestet: Pull Request **`nightly` → `main`** mergen. Der Workflow erhöht die Version von `1.0.0.0` auf **`1.1.0.0`**, committet sie, setzt den Tag `v1.1.0.0` und veröffentlicht das Release.
3. Danach führt der Workflow `main` automatisch in `nightly` zurück. Die nächste Nightly heißt dann `1.1.0.1`.

Beim allerersten Lauf auf `main` (noch kein Tag `v1.0.0.0`) wird **1.0.0.0** veröffentlicht, ohne zu erhöhen.

**Versionsnummer nie von Hand ändern**, das macht der Workflow. Zum Nachsehen oder für Sonderfälle gibt es `tools/version.py`:
```bash
python tools/version.py show              # aktuelle Version
python tools/version.py set 2.0.0.0       # z. B. für einen großen Versionssprung (auf main committen)
```

**Workflows** (`.github/workflows/`): `stable.yml` (main), `nightly.yml` (nightly), `_build.yml` (gemeinsamer Bau für Windows und Linux, inkl. SignPath). Manuell starten: Reiter **Actions** → Workflow wählen → **Run workflow**.

**Einmalig einrichten auf GitHub:**
- Branch `nightly` anlegen (wird sonst beim ersten Stable-Release automatisch aus `main` erstellt).
- Falls `main` geschützt ist (Branch protection): dem Workflow das Pushen erlauben (*Allow specified actors to bypass* → `github-actions`), sonst kann er die Versionsnummer nicht committen.

## Code Signing
Die Windows-Version wird mit einem eigenen Code-Signing-Zertifikat signiert (`sign_windows.bat`). Ablauf und Prüfung siehe [CODE_SIGNING.md](CODE_SIGNING.md).

## Schutz der Einstellungen
Einträge, Dateiordner und Benutzerkonten gelten für **alle Windows-Benutzer** des Rechners. Die Ansicht (hell/dunkel) speichert jeder Windows-Benutzer für sich.

| | Gemeinsame Einstellungen + Benutzerkonten | Persönliche Ansicht |
|---|---|---|
| Windows | `C:\ProgramData\DocuBuilder\.docubuilder.dat` | `%APPDATA%\DocuBuilder\.docubuilder-user.dat` |
| Linux | `/etc/docubuilder/.docubuilder.dat` | `~/.config/docubuilder/.docubuilder-user.dat` |

- **Verschlüsselt und versteckt:** In einem Texteditor ist nichts lesbar. Passwörter werden nur als Hash (PBKDF2) gespeichert.
- **Manipulationserkennung:** Wurde die Datei außerhalb des Programms verändert, startet DocuBuilder nicht mehr und meldet das. Ein Windows-Administrator muss die Datei dann löschen (danach Ersteinrichtung).
- **Nicht löschbar für normale Benutzer:** Der Ordner bekommt eigene Rechte: Windows-Benutzer dürfen die Datei lesen und (über DocuBuilder) überschreiben, aber nicht löschen oder umbenennen. So kann niemand durch Löschen eine neue Einrichtung mit eigenem Admin erzwingen.
  Das richtet DocuBuilder automatisch ein, wenn es den Ordner anlegt. Für den vollen Schutz DocuBuilder **einmal als Administrator starten** (Rechtsklick → *Als Administrator ausführen*); dann gehört der Ordner den Administratoren, und auch die Person, die ihn angelegt hat, kann ihn nicht mehr löschen. Unter Linux: einmal mit `sudo` starten, damit `/etc/docubuilder` angelegt wird.
- **Portabel:** Liegt eine leere Datei `portable.txt` neben dem Programm, liegen die gemeinsamen Einstellungen im Programmordner.
- **Alte Versionen:** Eine vorhandene `config.json` oder Einstellungsdatei aus Version 1.3 wird übernommen und danach gelöscht. Ein dort gesetztes Passwort für das Default-Fenster wird zum Passwort von „Admin“.

**Grenzen:** Die Anmeldung schützt vor normalen Benutzern. Windows-Administratoren können die Datei immer löschen (und damit alles zurücksetzen). Wer den Quelltext kennt, kann die Verschlüsselung nachbauen, weil der Schlüsselanteil im Programm steckt.

## Windows-Hinweis: Smart App Control
Ist unter Windows 11 **Smart App Control** eingeschaltet, blockiert Windows unsignierte Programme vollständig („Eine Anwendungssteuerungsrichtlinie hat diese Datei blockiert“, Fehler 4551), auch das Update-Setup. Abhilfe ist eine Signatur von DocuBuilder und Setup (siehe [CODE_SIGNING.md](CODE_SIGNING.md)). Bis dahin lässt sich DocuBuilder auf solchen PCs nur installieren oder aktualisieren, wenn Smart App Control ausgeschaltet ist (Windows-Sicherheit → App- & Browsersteuerung). Das ist eine Sicherheitsentscheidung für den PC; je nach Windows-Version lässt es sich danach nur durch Zurücksetzen von Windows wieder einschalten.

## Windows-Hinweis: SmartScreen
Beim ersten Start kann Windows „Der Computer wurde durch Windows geschützt" anzeigen. Das ist bei neuen Programmen normal, auch bei signierten: SmartScreen vertraut Dateien erst nach einer gewissen Zahl an Downloads und Starts.

So startest du DocuBuilder trotzdem:
1. Auf **Weitere Informationen** klicken
2. Auf **Trotzdem ausführen** klicken

Unter Rechtsklick → Eigenschaften → Digitale Signaturen siehst du den geprüften Herausgeber. Die Prüfsumme (`.sha256`) im Release kannst du mit `certutil -hashfile DocuBuilder.exe SHA256` vergleichen.

Auf Rechnern, die du selbst verwaltest, lässt sich die Meldung vorab entfernen: Rechtsklick auf die .exe → Eigenschaften → Haken bei **Zulassen** → OK (oder in PowerShell: `Unblock-File .\DocuBuilder.exe`).

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst.

Die Update-Suche ruft nur die öffentliche Release-Liste von GitHub ab (`api.github.com`) und lädt bei einem Update die Programmdatei von `github.com` herunter. Dabei werden keine Einträge, Benutzer oder sonstigen Daten übertragen. Die Suche läuft nur auf Knopfdruck oder, wenn unter *Einstellungen → Allgemein* eingeschaltet, einmal beim Anmelden.

## Lizenz
MIT
