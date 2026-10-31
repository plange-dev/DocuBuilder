@echo off
REM Erzeugt im Ordner dist:
REM   DocuBuilder-Setup-<Version>.msi  Installationspaket (installiert den Programmordner)
REM   DocuBuilder.exe                  portable Einzeldatei
REM   *.sha256                         Pruefsummen (passend zu den ggf. signierten Dateien)
REM Benoetigt Python 3 und das WiX Toolset 5 (dotnet tool install --global wix --version 5.0.2).
REM
REM Signieren (empfohlen, siehe CODE_SIGNING.md): vorher den Fingerprint setzen, z. B.
REM   set DOCUBUILDER_CERT_SHA1=0123456789ABCDEF0123456789ABCDEF01234567
REM Dann werden Programm, portable .exe und das .msi-Paket signiert.
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

if defined DOCUBUILDER_CERT_SHA1 (
  call tools\signfile.bat dist\DocuBuilder\DocuBuilder.exe || goto :fehler
  call tools\signfile.bat dist\DocuBuilder.exe || goto :fehler
) else (
  echo Hinweis: DOCUBUILDER_CERT_SHA1 nicht gesetzt - es wird NICHT signiert.
)

where wix >nul 2>&1 || goto :kein_wix
wix extension list -g | findstr /i "WixToolset.UI.wixext" >nul || wix extension add -g WixToolset.UI.wixext/5.0.2 || goto :fehler
py tools\make_license_rtf.py || goto :fehler
wix build -arch x64 -ext WixToolset.UI.wixext -culture de-DE -d AppVersion=%VER% -d AppChannel=%CHANNEL% -d ProjectDir="%ROOT%." -d DistDir="%ROOT%dist" -d LicenseRtf="%ROOT%build\License.rtf" installer\DocuBuilder.wxs -o dist\DocuBuilder-Setup-%VER%.msi || goto :fehler
if defined DOCUBUILDER_CERT_SHA1 call tools\signfile.bat dist\DocuBuilder-Setup-%VER%.msi || goto :fehler

for %%F in ("dist\DocuBuilder.exe" "dist\DocuBuilder-Setup-%VER%.msi") do (
  powershell -NoProfile -Command "$f='%%~F'; $h=(Get-FileHash $f -Algorithm SHA256).Hash.ToLower(); ($h + '  ' + (Split-Path $f -Leaf)) | Out-File -Encoding ascii ($f + '.sha256')" || goto :fehler
)
if not "%~1"=="" py tools\version.py set %OLDVER% >nul
echo.
echo Fertig: dist\DocuBuilder-Setup-%VER%.msi und dist\DocuBuilder.exe (mit .sha256)
pause
exit /b 0

:kein_wix
echo WiX Toolset nicht gefunden. Installieren mit: dotnet tool install --global wix --version 5.0.2
:fehler
if not "%~1"=="" py tools\version.py set %OLDVER% >nul
pause
exit /b 1
