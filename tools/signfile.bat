@echo off
REM Signiert eine Datei mit dem Code-Signing-Zertifikat aus DOCUBUILDER_CERT_SHA1 (SHA1-Fingerprint).
REM Wird von build_installer.bat und von Inno Setup (SignTool, auch fuer das Deinstallationsprogramm)
REM aufgerufen. Zeitstempel-Server: DOCUBUILDER_TIMESTAMP_URL (Standard: Certum).
REM Voraussetzungen: Windows SDK (signtool.exe); Zertifikat sichtbar (bei Certum: SimplySign angemeldet).
setlocal
set "FILE=%~1"
if "%FILE%"=="" (
  echo signfile: keine Datei angegeben
  exit /b 1
)
if not defined DOCUBUILDER_CERT_SHA1 (
  echo signfile: DOCUBUILDER_CERT_SHA1 ist nicht gesetzt
  exit /b 1
)
set "THUMB=%DOCUBUILDER_CERT_SHA1: =%"
set "TSA=%DOCUBUILDER_TIMESTAMP_URL%"
if "%TSA%"=="" set "TSA=http://time.certum.pl"

set "SIGNTOOL="
where signtool >nul 2>&1 && set "SIGNTOOL=signtool"
if not defined SIGNTOOL for /f "delims=" %%I in ('dir /b /s "%ProgramFiles(x86)%\Windows Kits\10\bin\signtool.exe" 2^>nul ^| findstr /i "\x64\\"') do set "SIGNTOOL=%%I"
if not defined SIGNTOOL (
  echo signfile: signtool.exe nicht gefunden - bitte das Windows SDK installieren
  exit /b 1
)

"%SIGNTOOL%" sign /sha1 %THUMB% /fd sha256 /tr %TSA% /td sha256 /d "DocuBuilder" "%FILE%"
if errorlevel 1 (
  echo signfile: Signieren fehlgeschlagen: %FILE%
  exit /b 1
)
"%SIGNTOOL%" verify /pa "%FILE%" >nul
if errorlevel 1 (
  echo signfile: Signatur ungueltig: %FILE%
  exit /b 1
)
echo signiert: %FILE%
exit /b 0
