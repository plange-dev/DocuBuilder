; Installationsprogramm für DocuBuilder (Inno Setup 6)
; Bauen: build_installer.bat  (oder: ISCC.exe /DAppVersion=1.0.0.0 installer\DocuBuilder.iss)
; Erwartet den Programmordner dist\DocuBuilder\ (PyInstaller --onedir) und erzeugt
; dist\DocuBuilder-Setup-<Version>.exe.

#ifndef AppVersion
  #define AppVersion "1.0.0.0"
#endif
#ifndef AppChannel
  #define AppChannel "stable"
#endif
#if AppChannel == "nightly"
  #define VerName "DocuBuilder " + AppVersion + " (Nightly)"
#else
  #define VerName "DocuBuilder " + AppVersion
#endif

[Setup]
; AppId nie ändern: daran erkennt Windows Updates und Deinstallation
AppId={{3912EB95-FF31-4D11-AE0E-6D244A30E1CA}
AppName=DocuBuilder
AppVersion={#AppVersion}
AppVerName={#VerName}
AppPublisher=Patrick Lange
AppPublisherURL=https://github.com/plange-dev/DocuBuilder
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\DocuBuilder
DefaultGroupName=DocuBuilder
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
LicenseFile=..\LICENSE
OutputDir=..\dist
OutputBaseFilename=DocuBuilder-Setup-{#AppVersion}
SetupIconFile=..\assets\docubuilder.ico
UninstallDisplayIcon={app}\DocuBuilder.exe
UninstallDisplayName={#VerName}
CloseApplications=yes
; Explorer am Ende benachrichtigen, damit Verknüpfungen ihr Symbol neu laden
ChangesAssociations=yes
#ifdef Signed
; Setup und Deinstallationsprogramm signieren (build_installer.bat übergibt /Sdocusign=...)
SignTool=docusign $f
SignedUninstaller=yes
#endif
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
; Verknüpfung auf dem Desktop aller Benutzer (Netzwerk-Version: jeder meldet sich mit eigenem Konto an)
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[UninstallDelete]
; von der Seite „Betriebsart“ angelegt (siehe [Code])
Type: files; Name: "{app}\server.txt"

[Files]
; Programmordner (PyInstaller --onedir): DocuBuilder.exe und _internal\
Source: "..\dist\DocuBuilder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion
; Eigenes Symbol für die Verknüpfungen (siehe [Icons])
Source: "..\assets\docubuilder.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; AppUserModelID wie im Programm: angeheftetes Taskleisten-Symbol und Fenster gehören zusammen.
; IconFilename: Versionen bis 1.0.1.0 hatten das Python-Symbol in DocuBuilder.exe. Windows merkt
; sich Symbole pro Pfad, deshalb zeigte die Verknüpfung nach einem Update weiter das alte Symbol.
; Mit eigener .ico-Datei wird das richtige Symbol neu geladen (Issue #1).
Name: "{autoprograms}\DocuBuilder"; Filename: "{app}\DocuBuilder.exe"; IconFilename: "{app}\docubuilder.ico"; AppUserModelID: "PatrickLange.DocuBuilder"
Name: "{autodesktop}\DocuBuilder"; Filename: "{app}\DocuBuilder.exe"; IconFilename: "{app}\docubuilder.ico"; AppUserModelID: "PatrickLange.DocuBuilder"; Tasks: desktopicon

[Run]
; Gemeinsamen Einstellungsordner (C:\ProgramData\DocuBuilder) mit den passenden Rechten anlegen
Filename: "{app}\DocuBuilder.exe"; Parameters: "--setup-machine"; StatusMsg: "Einstellungsordner wird eingerichtet ..."; Flags: runhidden waituntilterminated
; Programm danach als normaler Benutzer starten (nicht mit Adminrechten)
Filename: "{app}\DocuBuilder.exe"; Description: "{cm:LaunchProgram,DocuBuilder}"; Flags: nowait postinstall skipifsilent runasoriginaluser
; Update aus DocuBuilder heraus (still, mit /UPDATE=1): Programm danach automatisch neu starten.
; Über den Explorer, damit es mit dessen sauberer Umgebung startet und nicht die Variablen des
; alten DocuBuilder erbt (sonst "Failed to load Python DLL ... _MEI...").
Filename: "{win}\explorer.exe"; Parameters: """{app}\DocuBuilder.exe"""; Flags: nowait runasoriginaluser; Check: IsUpdate

[Code]
function SetEnvironmentVariableW(lpName: String; lpValue: Cardinal): Integer;
  external 'SetEnvironmentVariableW@kernel32.dll stdcall';
function OpenProcess(dwDesiredAccess: Cardinal; bInheritHandle: Integer; dwProcessId: Cardinal): Cardinal;
  external 'OpenProcess@kernel32.dll stdcall';
function WaitForSingleObject(hHandle: Cardinal; dwMilliseconds: Cardinal): Cardinal;
  external 'WaitForSingleObject@kernel32.dll stdcall';
function CloseHandle(hObject: Cardinal): Integer;
  external 'CloseHandle@kernel32.dll stdcall';

const
  SYNCHRONIZE = $00100000;
  MODE_SINGLE = 0;
  MODE_SERVER = 1;
  MODE_CLIENT = 2;
  SHARE_NAME = 'DocuBuilder';

var
  ModePage: TInputOptionWizardPage;
  PathPage: TInputQueryWizardPage;

{ Betriebsart (Netzwerk-Version). Still installieren:
    /BETRIEB=einzel | server | arbeitsplatz   /ORDNER="\\Server\DocuBuilder"
  Einzelplatz: Einstellungen in C:\ProgramData\DocuBuilder (wie bisher).
  Server: gemeinsamer Datenordner auf diesem Rechner, wird als \\<Rechner>\DocuBuilder freigegeben.
  Arbeitsplatz: Daten liegen im freigegebenen Ordner des Servers. }
function SelectedMode: Integer;
begin
  Result := MODE_SINGLE;
  if ModePage.Values[MODE_SERVER] then
    Result := MODE_SERVER
  else if ModePage.Values[MODE_CLIENT] then
    Result := MODE_CLIENT;
end;

function IsServerMode: Boolean;
begin
  Result := SelectedMode = MODE_SERVER;
end;

function DefaultFolder(Mode: Integer): String;
begin
  if Mode = MODE_SERVER then
    Result := ExpandConstant('{sd}\DocuBuilder-Daten')
  else
    Result := '\\SERVER\' + SHARE_NAME;
end;

procedure InitializeWizard;
var
  Mode: Integer;
  Param, Folder: String;
begin
  ModePage := CreateInputOptionPage(wpSelectDir, 'Betriebsart',
    'Wie wird DocuBuilder genutzt?',
    'Im Netzwerk liegen Einträge, Benutzer und Rechte zentral auf dem Server. Jeder Arbeitsplatz ' +
    'greift darauf zu; alle können gleichzeitig arbeiten.', True, False);
  ModePage.Add('Einzelplatz – Einstellungen nur auf diesem Rechner');
  ModePage.Add('Server – gemeinsamer Datenordner liegt auf diesem Rechner und wird freigegeben');
  ModePage.Add('Arbeitsplatz im Netzwerk – Daten liegen auf dem Server');

  PathPage := CreateInputQueryPage(ModePage.ID, 'Gemeinsamer Ordner',
    'Wo liegen die gemeinsamen Daten?',
    'Server: lokaler Ordner auf diesem Rechner (z. B. D:\DocuBuilder-Daten).' + #13#10 +
    'Arbeitsplatz: freigegebener Ordner des Servers (z. B. \\SERVER\DocuBuilder).');
  PathPage.Add('Ordner:', False);

  Mode := StrToIntDef(GetPreviousData('Betriebsart', '0'), MODE_SINGLE);
  Folder := GetPreviousData('Ordner', '');
  Param := Lowercase(ExpandConstant('{param:BETRIEB|}'));
  if Param = 'einzel' then Mode := MODE_SINGLE;
  if Param = 'server' then Mode := MODE_SERVER;
  if Param = 'arbeitsplatz' then Mode := MODE_CLIENT;
  Param := ExpandConstant('{param:ORDNER|}');
  if Param <> '' then Folder := Param;
  if (Mode < MODE_SINGLE) or (Mode > MODE_CLIENT) then Mode := MODE_SINGLE;
  ModePage.Values[Mode] := True;
  if Folder = '' then Folder := DefaultFolder(Mode);
  PathPage.Values[0] := Folder;
end;

procedure RegisterPreviousData(PreviousDataKey: Integer);
begin
  SetPreviousData(PreviousDataKey, 'Betriebsart', IntToStr(SelectedMode));
  SetPreviousData(PreviousDataKey, 'Ordner', PathPage.Values[0]);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := (PageID = PathPage.ID) and (SelectedMode = MODE_SINGLE);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Folder: String;
begin
  Result := True;
  if CurPageID = ModePage.ID then
  begin
    { Vorschlag passend zur Betriebsart, solange noch der andere Vorschlag drinsteht }
    if (PathPage.Values[0] = DefaultFolder(MODE_SERVER)) or (PathPage.Values[0] = DefaultFolder(MODE_CLIENT)) then
      PathPage.Values[0] := DefaultFolder(SelectedMode);
  end
  else if CurPageID = PathPage.ID then
  begin
    Folder := Trim(PathPage.Values[0]);
    PathPage.Values[0] := Folder;
    if Folder = '' then
    begin
      MsgBox('Bitte einen Ordner angeben.', mbError, MB_OK);
      Result := False;
    end
    else if IsServerMode and (Copy(Folder, 1, 2) = '\\') then
    begin
      MsgBox('Auf dem Server bitte einen lokalen Ordner angeben (z. B. D:\DocuBuilder-Daten). ' +
             'Er wird anschließend als \\' + GetComputerNameString + '\' + SHARE_NAME + ' freigegeben.',
             mbError, MB_OK);
      Result := False;
    end
    else if (SelectedMode = MODE_CLIENT) and not DirExists(Folder) then
      Result := MsgBox('Der Ordner ist von hier aus nicht erreichbar:' + #13#10 + Folder + #13#10#13#10 +
                       'Trotzdem verwenden? (z. B. wenn der Server gerade aus ist)',
                       mbConfirmation, MB_YESNO) = IDYES;
  end;
end;

{ server.txt neben dem Programm: DocuBuilder liest daraus den gemeinsamen Ordner.
  Vor dem Kopieren der Dateien schreiben, damit "--setup-machine" ([Run]) ihn schon kennt. }
procedure WriteServerFile;
var
  F: String;
begin
  F := ExpandConstant('{app}\server.txt');
  if SelectedMode = MODE_SINGLE then
  begin
    DeleteFile(F);
    Exit;
  end;
  ForceDirectories(ExpandConstant('{app}'));
  SaveStringToFile(F, '# Gemeinsamer Ordner von DocuBuilder (vom Installationsprogramm angelegt)' + #13#10 +
                      PathPage.Values[0] + #13#10, False);
end;

{ Server: Ordner als \\<Rechner>\DocuBuilder freigeben (Ändern für angemeldete Benutzer; was
  jeder darf, regeln die Ordnerrechte, die DocuBuilder mit --setup-machine setzt). }
procedure CreateShare;
var
  ResultCode: Integer;
begin
  Exec('powershell.exe', '-NoProfile -ExecutionPolicy Bypass -Command "' +
       'if (-not (Get-SmbShare -Name ''' + SHARE_NAME + ''' -ErrorAction SilentlyContinue)) { ' +
       '$u = ([Security.Principal.SecurityIdentifier]''S-1-5-11'').Translate([Security.Principal.NTAccount]).Value; ' +
       '$a = ([Security.Principal.SecurityIdentifier]''S-1-5-32-544'').Translate([Security.Principal.NTAccount]).Value; ' +
       'New-SmbShare -Name ''' + SHARE_NAME + ''' -Path ''' + PathPage.Values[0] + ''' -ChangeAccess $u -FullAccess $a | Out-Null }"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    WriteServerFile
  else if (CurStep = ssPostInstall) and IsServerMode then
  begin
    ForceDirectories(PathPage.Values[0]);
    CreateShare;
    if not WizardSilent then
      MsgBox('Der gemeinsame Ordner ist freigegeben als:' + #13#10#13#10 +
             '\\' + GetComputerNameString + '\' + SHARE_NAME + #13#10#13#10 +
             'Diesen Pfad bei der Installation auf den Arbeitsplätzen angeben. ' +
             'Bitte DocuBuilder hier einmal starten und die Ersteinrichtung (Admin-Passwort, ' +
             'Benutzer, Bereiche) vornehmen, bevor die Arbeitsplätze es benutzen.',
             mbInformation, MB_OK);
  end;
end;

{ Update aus DocuBuilder: /WAITPID=<Prozess-ID> – erst weitermachen, wenn das alte DocuBuilder
  beendet ist (höchstens 60 Sekunden; danach greift CloseApplications). }
procedure WaitForOldProgram;
var
  Pid, Handle: Cardinal;
begin
  Pid := StrToIntDef(ExpandConstant('{param:WAITPID|0}'), 0);
  if Pid = 0 then
    Exit;
  Handle := OpenProcess(SYNCHRONIZE, 0, Pid);
  if Handle <> 0 then
  begin
    WaitForSingleObject(Handle, 60000);
    CloseHandle(Handle);
  end;
end;

function IsUpdate: Boolean;
begin
  Result := ExpandConstant('{param:UPDATE|0}') = '1';
end;

{ Variablen entfernen, die eine einzelne PyInstaller-.exe (das alte DocuBuilder) an das Setup
  vererbt hat. Sonst sucht die neu gestartete DocuBuilder.exe (z. B. --setup-machine) ihre
  Dateien im bereits gelöschten Temp-Ordner des alten Programms. }
procedure RemoveInheritedPyInstallerEnv;
var
  Names: TArrayOfString;
  I: Integer;
begin
  Names := ['_PYI_APPLICATION_HOME_DIR', '_PYI_ARCHIVE_FILE', '_PYI_PARENT_PROCESS_LEVEL',
            '_PYI_SPLASH_IPC', '_MEIPASS2'];
  for I := 0 to GetArrayLength(Names) - 1 do
    SetEnvironmentVariableW(Names[I], 0);
  if Pos('\_MEI', GetEnv('TCL_LIBRARY')) > 0 then
    SetEnvironmentVariableW('TCL_LIBRARY', 0);
  if Pos('\_MEI', GetEnv('TK_LIBRARY')) > 0 then
    SetEnvironmentVariableW('TK_LIBRARY', 0);
end;

function InitializeSetup: Boolean;
begin
  RemoveInheritedPyInstallerEnv;
  WaitForOldProgram;
  Result := True;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Dir: String;
  ResultCode: Integer;
begin
  if CurUninstallStep <> usPostUninstall then
    Exit;
  Dir := ExpandConstant('{commonappdata}\DocuBuilder');
  if not DirExists(Dir) or UninstallSilent then
    Exit;
  if MsgBox('Sollen auch alle Einstellungen gelöscht werden?' + #13#10 +
            '(Einträge, Benutzer und Passwörter für alle Windows-Benutzer)' + #13#10#13#10 +
            'Bei "Nein" bleiben sie für eine spätere Neuinstallation erhalten.',
            mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
  begin
    Exec(ExpandConstant('{cmd}'), '/c attrib -R -H "' + Dir + '\*" & rmdir /s /q "' + Dir + '"',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    DelTree(ExpandConstant('{userappdata}\DocuBuilder'), True, True, True);
  end;
end;
