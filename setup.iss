; License Tracker - Inno Setup installer script
; Per-user install (no admin rights): DB is stored next to the EXE,
; so the app must live in a writable folder (LocalAppData).

#define AppName "License Tracker"
#define AppVersion "1.1.16"
#define AppExe "LicenseTracker.exe"
#define AppPublisher "License Tracker"

[Setup]
AppId={{7F3C9A6E-4B21-4E15-8F9D-A1C2E9B7D543}}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#AppExe}
OutputDir=installer
OutputBaseFilename=LicenseTracker_Setup
Compression=lzma2/max
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
ShowLanguageDialog=no
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; Flags: checkedonce

[Run]
Filename: "{app}\{#AppExe}"; Description: "Запустить {#AppName}"; Flags: nowait postinstall skipifsilent
