; KnobMixer Inno Setup Script
#define MyAppName      "KnobMixer"
#define MyAppVersion   "2.7.4"
#define MyAppPublisher "KnobMixer"
#define MyAppURL       "https://github.com/KnobMixer/KnobMixer"
#define MyAppExeName   "KnobMixer.exe"

[Setup]
AppId={{A3F8B2C1-4D5E-4F6A-8B9C-0D1E2F3A4B5C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=dist_installer
OutputBaseFilename=KnobMixer_Setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName=KnobMixer

; Do NOT use CloseApplications=yes — KnobMixer intercepts WM_CLOSE to minimize
; to tray instead of closing, so the graceful close silently fails and the
; installer errors out. We kill the process explicitly in [Code] BeforeInstall
; which runs before any file operations begin.
CloseApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "startupicon"; Description: "Start KnobMixer when Windows starts"; GroupDescription: "Startup"

[Files]
Source: "dist\KnobMixer\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\KnobMixer\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}";        FileName: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall KnobMixer"; FileName: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";  FileName: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Launch after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "taskkill.exe"; Parameters: "/f /im KnobMixer.exe"; Flags: runhidden; RunOnceId: "KillApp"
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v KnobMixer /f"; Flags: runhidden; RunOnceId: "RemoveStartup"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "KnobMixer"; \
  ValueData: """{app}\{#MyAppExeName}"""; \
  Tasks: startupicon; Flags: uninsdeletevalue

[Code]
// Kill KnobMixer BEFORE any file operations begin.
// This is the correct place — CloseApplications=yes sends WM_CLOSE
// which KnobMixer intercepts to minimize rather than close.
// BeforeInstall on the first file entry runs before files are touched.
procedure KillKnobMixer();
var
  ResultCode: Integer;
begin
  // Kill gracefully first, then force if needed
  Exec('taskkill.exe', '/im KnobMixer.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(800);
  // Force kill in case graceful didn't work
  Exec('taskkill.exe', '/f /im KnobMixer.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(400);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    KillKnobMixer();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataPath: String;
begin
  if CurUninstallStep = usUninstall then
    KillKnobMixer();

  if CurUninstallStep = usPostUninstall then
  begin
    AppDataPath := ExpandConstant('{userappdata}\KnobMixer');
    if DirExists(AppDataPath) then
    begin
      if MsgBox('Remove KnobMixer settings and data?' + #13#10 +
                '(Hotkeys, groups, and preferences)' + #13#10#13#10 +
                'Click Yes to delete, No to keep them.',
                mbConfirmation, MB_YESNO) = IDYES then
        DelTree(AppDataPath, True, True, True);
    end;
  end;
end;
