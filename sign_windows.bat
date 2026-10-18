@echo off
setlocal EnableExtensions
REM ============================================================
REM  Signiert DocuBuilder.exe mit einem Code-Signing-Zertifikat
REM  (z. B. Certum SimplySign) und erzeugt eine neue Pruefsumme.
REM
REM  Aufruf:  sign_windows.bat [Fingerprint] [Datei]
REM    Fingerprint = SHA1-Fingerprint des Zertifikats (Zertifikatsverwaltung
REM                  "certmgr.msc" - Eigene Zertifikate - Details - Fingerabdruck)
REM                  Alternativ als Umgebungsvariable DOCUBUILDER_CERT_SHA1 setzen.
REM    Datei       = Standard: dist\DocuBuilder.exe
REM
REM  Voraussetzungen:
REM   - Windows SDK (enthaelt signtool.exe) oder Visual Studio Build Tools
REM   - Zertifikat ist sichtbar (bei Certum: SimplySign Desktop gestartet und angemeldet)
REM ============================================================

set "FILE=%~2"
if "%FILE%"=="" set "FILE=dist\DocuBuilder.exe"

set "THUMB=%~1"
if "%THUMB%"=="" set "THUMB=%DOCUBUILDER_CERT_SHA1%"
if "%THUMB%"=="" set /p THUMB=SHA1-Fingerprint des Zertifikats: 
set "THUMB=%THUMB: =%"

set "TSA=%DOCUBUILDER_TIMESTAMP_URL%"
if "%TSA%"=="" set "TSA=http://time.certum.pl"

if not exist "%FILE%" (
  echo FEHLER: Datei nicht gefunden: %FILE%
  goto :fail
)
if "%THUMB%"=="" (
  echo FEHLER: Kein Fingerprint angegeben.
  goto :fail
)

REM --- Zertifikat pruefen (Fingerprint vorhanden, privater Schluessel nutzbar) ---
powershell -NoProfile -Command "$t='%THUMB%'.ToUpper(); $all=@(Get-ChildItem Cert:\CurrentUser\My,Cert:\LocalMachine\My -CodeSigningCert -EA 0); $c=$all | Where-Object { $_.Thumbprint -eq $t } | select -First 1; if ($c -and $c.HasPrivateKey) { 'Zertifikat gefunden: ' + $c.Subject + ' (gueltig bis ' + $c.NotAfter.ToString('dd.MM.yyyy') + ')'; exit 0 }; if ($c) { 'FEHLER: Zertifikat gefunden, aber ohne nutzbaren privaten Schluessel (SimplySign angemeldet? Token eingesteckt?)'; exit 1 }; 'FEHLER: Kein Code-Signing-Zertifikat mit dem Fingerprint ' + $t + ' gefunden.'; if ($all.Count) { 'Vorhandene Code-Signing-Zertifikate:'; $all | ForEach-Object { '  ' + $_.Thumbprint + '  ' + $_.Subject } } else { 'Auf diesem PC ist ueberhaupt kein Code-Signing-Zertifikat vorhanden.'; 'Bei Certum: SimplySign Desktop starten und anmelden. Sonst wird zuerst ein Zertifikat benoetigt (siehe CODE_SIGNING.md).' }; exit 1"
if errorlevel 1 goto :fail

REM --- signtool suchen ---
set "SIGNTOOL="
where signtool >nul 2>&1 && set "SIGNTOOL=signtool"
if not defined SIGNTOOL (
  set "WK=%ProgramFiles(x86)%\Windows Kits\10\bin"
  for /f "delims=" %%I in ('dir /b /s "%ProgramFiles(x86)%\Windows Kits\10\bin\signtool.exe" 2^>nul ^| findstr /i "\\x64\\"') do set "SIGNTOOL=%%I"
)
if not defined SIGNTOOL (
  echo FEHLER: signtool.exe nicht gefunden. Bitte das Windows SDK installieren:
  echo https://developer.microsoft.com/windows/downloads/windows-sdk/
  goto :fail
)

echo Signiere %FILE% ...
"%SIGNTOOL%" sign /sha1 %THUMB% /fd sha256 /tr %TSA% /td sha256 /d "DocuBuilder" "%FILE%"
if errorlevel 1 (
  echo FEHLER beim Signieren. Ist das Zertifikat sichtbar bzw. SimplySign angemeldet?
  goto :fail
)

echo Pruefe Signatur ...
"%SIGNTOOL%" verify /pa /v "%FILE%"
if errorlevel 1 (
  echo FEHLER: Die Signatur konnte nicht verifiziert werden.
  goto :fail
)

echo Erzeuge Pruefsumme ...
powershell -NoProfile -Command "$h=(Get-FileHash -Algorithm SHA256 '%FILE%').Hash.ToLower(); \"$h  DocuBuilder.exe\" | Out-File -Encoding ascii '%FILE%.sha256'"

echo.
echo Fertig. Signierte Datei: %FILE%
echo Pruefsumme:             %FILE%.sha256
echo.
echo Optional ans GitHub-Release anhaengen (benoetigt GitHub CLI):
echo   gh release upload v1.0.0 "%FILE%" "%FILE%.sha256" --clobber
pause
exit /b 0

:fail
pause
exit /b 1
