# Code Signing

Ohne digitale Signatur blockiert **Smart App Control** (Windows 11) DocuBuilder und vor allem das **Update-Setup** („Eine Anwendungssteuerungsrichtlinie hat diese Datei blockiert“, Fehler 4551). SmartScreen warnt zusätzlich bei jedem neuen Download. Deshalb sollen `DocuBuilder.exe` **und** `DocuBuilder-Setup-<Version>.exe` mit einem Zertifikat einer anerkannten Zertifizierungsstelle signiert werden. Ein selbst erstelltes Zertifikat reicht dafür nicht.

## Für Privatpersonen: eigenes Zertifikat (z. B. Certum), lokal signieren
Ein Code-Signing-Zertifikat für Open-Source-Entwickler (z. B. von Certum, Zertifikat in der Cloud über die App **SimplySign**) wird auf deinen Namen ausgestellt; in Windows erscheint als Herausgeber z. B. „Open Source Developer, Patrick Lange“. Weil SimplySign eine Anmeldung auf deinem PC braucht, wird **lokal** signiert, nicht in GitHub Actions.

### Einmalig einrichten
1. Zertifikat bestellen und die Identitätsprüfung abschließen.
2. **SimplySign Desktop** installieren und anmelden. Danach ist das Zertifikat in Windows sichtbar.
3. **Windows SDK** installieren (enthält `signtool.exe`): `winget install Microsoft.WindowsSDK.10.0.26100`
   (oder über den Visual Studio Installer). Inno Setup: `winget install JRSoftware.InnoSetup`.
4. Fingerprint ablesen: `certmgr.msc` → Eigene Zertifikate → Zertifikat → Details → **Fingerabdruck**.

### Release signieren
GitHub Actions veröffentlicht jede Version zunächst unsigniert. So ersetzt du die Dateien durch signierte:

1. SimplySign Desktop starten und anmelden.
2. In einer Eingabeaufforderung im Projektordner (auf dem Stand des Releases, z. B. Branch `nightly`):
   ```
   set DOCUBUILDER_CERT_SHA1=<Fingerprint ohne Leerzeichen>
   build_installer.bat 1.0.1.8 nightly
   ```
   Version und Kanal genau wie im Release angeben (Stable: `build_installer.bat 1.1.0.0`). Das Skript signiert **Programm, portable .exe, Setup und Deinstallationsprogramm**, prüft die Signaturen und erzeugt die Prüfsummen. Die Version im Projekt wird danach wiederhergestellt.
3. Auf GitHub im Release (**Edit**) diese vier Dateien aus `dist\` ersetzen:
   `DocuBuilder.exe`, `DocuBuilder.exe.sha256`, `DocuBuilder-Setup-<Version>.exe`, `DocuBuilder-Setup-<Version>.exe.sha256`.

Wichtig: Immer **Datei und Prüfsumme zusammen** ersetzen – das Update-Tool in DocuBuilder prüft die Prüfsumme und lehnt sonst ab.

Zum Signieren einzelner Dateien gibt es außerdem `sign_windows.bat <Fingerprint> <Datei>`.

## Alternative: SignPath (automatisch in GitHub Actions)
Die Pipeline (`.github/workflows/_build.yml`) signiert automatisch, sobald SignPath eingerichtet ist. Solange die Variablen fehlen, baut sie unsigniert weiter.

Für Open-Source-Projekte gibt es kostenloses Signieren über die **SignPath Foundation** (Antrag auf signpath.org). Voraussetzungen erfüllt DocuBuilder bereits: öffentliches Repo, MIT-Lizenz, Build über GitHub Actions.

### Einrichten
1. Bei SignPath (bzw. SignPath Foundation) ein Projekt anlegen und das GitHub-Repo `plange-dev/DocuBuilder` als vertrauenswürdige Build-Quelle verbinden.
2. Zwei **Artifact Configurations** anlegen:
   - für das Programm: Inhalt von `signpath/artifact-configuration.xml` (signiert die portable `DocuBuilder.exe` und `DocuBuilder/DocuBuilder.exe` im Programmordner)
   - für das Setup: Inhalt von `signpath/artifact-configuration-setup.xml` (signiert `DocuBuilder-Setup.exe`)
3. Im GitHub-Repo unter **Settings → Secrets and variables → Actions** eintragen:

   | Art | Name | Inhalt |
   |---|---|---|
   | Secret | `SIGNPATH_API_TOKEN` | API-Token eines SignPath-CI-Benutzers |
   | Variable | `SIGNPATH_ORGANIZATION_ID` | Organisations-ID |
   | Variable | `SIGNPATH_PROJECT_SLUG` | Projekt-Slug |
   | Variable | `SIGNPATH_SIGNING_POLICY_SLUG` | Signing Policy (z. B. `release-signing`) |
   | Variable | `SIGNPATH_ARTIFACT_CONFIGURATION_SLUG` | Slug der Konfiguration für das Programm |
   | Variable | `SIGNPATH_SETUP_ARTIFACT_CONFIGURATION_SLUG` | Slug der Konfiguration für das Setup |

4. Einen Nightly-Build anstoßen (Push auf `nightly` oder **Actions → Nightly → Run workflow**). Im Log stehen „Signaturstatus: Valid“ und „Signaturstatus Setup: Valid“.

Wird die SignPath Foundation genutzt, muss im README und hier der Hinweis stehen:
„Free code signing provided by SignPath.io, certificate by SignPath Foundation.“

### Ablauf
1. PyInstaller baut `DocuBuilder.exe` → SignPath signiert sie.
2. Inno Setup baut das Setup mit der **signierten** `DocuBuilder.exe` → SignPath signiert das Setup.
3. Prüfsummen (`.sha256`) werden erst danach berechnet, passen also zu den signierten Dateien.

### Einschränkung: Deinstallationsprogramm
Über SignPath bleibt das Deinstallationsprogramm (`unins000.exe`), das Inno Setup selbst erzeugt, **unsigniert**; unter Smart App Control kann das Deinstallieren dann blockiert werden. Beim lokalen Signieren mit `build_installer.bat` (siehe oben) wird es dagegen mitsigniert.

## Weniger Fehlalarme bei Virenscannern
- Das Setup installiert einen **Programmordner** (`DocuBuilder.exe` + `_internal\`) statt einer sich selbst entpackenden Einzeldatei. Solche Einzeldateien werden von manchen Virenscannern fälschlich verdächtigt; die portable `DocuBuilder.exe` bleibt als Alternative.
- Nach jedem Release die Dateien z. B. bei VirusTotal prüfen. Meldet ein Scanner fälschlich etwas, die Datei beim Hersteller als Fehlalarm einreichen (Microsoft: „Submit a file for malware analysis“).

## Signatur prüfen
Rechtsklick auf die .exe → Eigenschaften → Digitale Signaturen, oder in PowerShell:
```
Get-AuthenticodeSignature .\DocuBuilder-Setup-1.0.1.4.exe
```
Zusätzlich liegt jedem Release eine `.sha256`-Prüfsumme bei (`certutil -hashfile <Datei> SHA256`).

## SmartScreen und Reputation
Auch eine gültige Signatur verhindert die SmartScreen-Meldung bei neuen Dateien nicht sofort; die Reputation baut sich über Downloads auf. Hilfreich ist, immer mit demselben Zertifikat zu signieren. Smart App Control dagegen lässt Programme mit gültiger Signatur einer anerkannten Stelle zu.

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst (z. B. durch das Öffnen eines hinterlegten Web-Links oder die Update-Suche).
