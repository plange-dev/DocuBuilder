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
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

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
