; KnobMixer Inno Setup Script
; Compile with Inno Setup 6+ from https://jrsoftware.org/isinfo.php

#define MyAppName      "KnobMixer"
#define MyAppVersion   "1.0"
#define MyAppPublisher "KnobMixer"
#define MyAppURL       "https://github.com"
#define MyAppExeName   "KnobMixer.exe"

[Setup]
AppId={{A3F8B2C1-4D5E-4F6A-8B9C-0D1E2F3A4B5C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
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
CloseApplications=yes

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
; Launch app after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Registry]
; Startup with Windows (if task selected)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "KnobMixer"; \
  ValueData: """{app}\{#MyAppExeName}"""; \
  Tasks: startupicon; Flags: uninsdeletevalue

[UninstallRun]
; Remove from startup on uninstall
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v KnobMixer /f"; Flags: runhidden; RunOnceId: "RemoveStartup"
