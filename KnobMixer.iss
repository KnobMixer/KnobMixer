; KnobMixer Inno Setup Script
#define MyAppName      "KnobMixer"
#define MyAppVersion   "2.2"
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

; Automatically close the running app before install/uninstall
CloseApplications=yes
CloseApplicationsFilter=KnobMixer.exe
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "startupicon"; Description: "Start KnobMixer when Windows starts"; GroupDescription: "Startup"
Name: "deleteconfig"; Description: "Remove my settings and data on uninstall"; GroupDescription: "Uninstall options"; Flags: unchecked

[Files]
Source: "dist\KnobMixer\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\KnobMixer\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}";        FileName: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall KnobMixer"; FileName: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";  FileName: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Close any running instance before files are copied
Filename: "taskkill.exe"; Parameters: "/f /im KnobMixer.exe"; Flags: runhidden; Check: IsAppRunning

; Launch after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Kill the app first so files can be deleted cleanly
Filename: "taskkill.exe"; Parameters: "/f /im KnobMixer.exe"; Flags: runhidden; RunOnceId: "KillApp"

; Remove from startup registry
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v KnobMixer /f"; Flags: runhidden; RunOnceId: "RemoveStartup"

[UninstallDelete]
; Always remove these on uninstall
Type: filesandordirs; Name: "{app}"

[Registry]
; Startup with Windows (if task selected during install)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "KnobMixer"; \
  ValueData: """{app}\{#MyAppExeName}"""; \
  Tasks: startupicon; Flags: uninsdeletevalue

[Code]
function IsAppRunning: Boolean;
var
  ResultCode: Integer;
begin
  Exec('tasklist.exe', '/fi "imagename eq KnobMixer.exe" /fo csv /nh', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := (ResultCode = 0);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataPath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Only delete AppData config if user ticked the option during install
    // (We store this choice in registry during install)
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
