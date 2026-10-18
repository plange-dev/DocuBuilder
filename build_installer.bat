@echo off
REM Erzeugt im Ordner dist:
REM   DocuBuilder-Setup-<Version>.exe  Installationsprogramm (installiert den Programmordner)
REM   DocuBuilder.exe                  portable Einzeldatei
REM   *.sha256                         Pruefsummen (passend zu den ggf. signierten Dateien)
REM Benoetigt Python 3 und Inno Setup 6 (winget install JRSoftware.InnoSetup).
REM
REM Signieren (empfohlen, siehe CODE_SIGNING.md): vorher den Fingerprint setzen, z. B.
REM   set DOCUBUILDER_CERT_SHA1=0123456789ABCDEF0123456789ABCDEF01234567
REM Dann werden Programm, portable .exe, Setup und Deinstallationsprogramm signiert.
REM
REM Optional Version und Kanal eines bereits veroeffentlichten Releases (zum Nachsignieren):
REM   build_installer.bat 1.0.1.8 nightly
REM Die Version im Projekt wird danach wiederhergestellt.
setlocal
cd /d "%~dp0"
set "ROOT=%~dp0"
for /f %%v in ('py tools\version.py show') do set OLDVER=%%v
set "VER=%OLDVER%"
set "CHANNEL=stable"
if not "%~1"=="" set "VER=%~1"
if not "%~2"=="" set "CHANNEL=%~2"
if not "%~1"=="" py tools\version.py set %VER% --channel %CHANNEL% || goto :fehler

py -m pip install --upgrade pyinstaller || goto :fehler
REM Programmordner fuer das Setup (keine sich selbst entpackende Einzeldatei: weniger Fehlalarme)
py -m PyInstaller --noconfirm --onedir --windowed --name DocuBuilder --workpath build\onedir --version-file version_info.txt --icon assets/docubuilder.ico --add-data "assets/docubuilder.ico:assets" --add-data "assets/docubuilder.png:assets" docubuilder.py || goto :fehler
REM Portable Einzeldatei
py -m PyInstaller --noconfirm --onefile --windowed --name DocuBuilder --workpath build\onefile --version-file version_info.txt --icon assets/docubuilder.ico --add-data "assets/docubuilder.ico:assets" --add-data "assets/docubuilder.png:assets" docubuilder.py || goto :fehler

set "SIGNARGS="
if defined DOCUBUILDER_CERT_SHA1 (
  call tools\signfile.bat dist\DocuBuilder\DocuBuilder.exe || goto :fehler
  call tools\signfile.bat dist\DocuBuilder.exe || goto :fehler
  set SIGNARGS=/DSigned "/Sdocusign=cmd.exe /c call $q%ROOT%tools\signfile.bat$q $p"
) else (
  echo Hinweis: DOCUBUILDER_CERT_SHA1 nicht gesetzt - es wird NICHT signiert.
)

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" goto :kein_inno

"%ISCC%" /DAppVersion=%VER% /DAppChannel=%CHANNEL% %SIGNARGS% installer\DocuBuilder.iss || goto :fehler

for %%F in ("dist\DocuBuilder.exe" "dist\DocuBuilder-Setup-%VER%.exe") do (
  powershell -NoProfile -Command "$f='%%~F'; $h=(Get-FileHash $f -Algorithm SHA256).Hash.ToLower(); ($h + '  ' + (Split-Path $f -Leaf)) | Out-File -Encoding ascii ($f + '.sha256')" || goto :fehler
)
if not "%~1"=="" py tools\version.py set %OLDVER% >nul
echo.
echo Fertig: dist\DocuBuilder-Setup-%VER%.exe und dist\DocuBuilder.exe (mit .sha256)
pause
exit /b 0

:kein_inno
echo Inno Setup 6 nicht gefunden. Installieren mit: winget install JRSoftware.InnoSetup
:fehler
if not "%~1"=="" py tools\version.py set %OLDVER% >nul
pause
exit /b 1
