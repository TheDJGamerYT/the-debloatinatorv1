#define AppName "BloatGuard"
#define AppVersion "1.0.1"
#define Publisher "DJ Studios"
#define AppExe "BloatGuard.exe"
#define AgentExe "BloatGuardAgent.exe"

[Setup]
AppId={{9C8F6E1A-9E8A-4F51-A7E7-CE8E4B12B0C1}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
UninstallDisplayIcon={app}\{#AppExe}
OutputBaseFilename=BloatGuard-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
WizardStyle=modern
DisableProgramGroupPage=no
DisableDirPage=no

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "autostart"; Description: "Start with Windows (run Agent on logon)"; GroupDescription: "Optional tasks:"; Flags: unchecked

[Files]
; Copy EVERYTHING from the app folder
Source: "..\dist\BloatGuard\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Copy EVERYTHING from the agent folder
Source: "..\dist\BloatGuardAgent\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{commonappdata}\BloatGuard"; Permissions: users-modify

[Icons]
Name: "{group}\BloatGuard"; Filename: "{app}\{#AppExe}"
Name: "{group}\BloatGuard (Open Data Folder)"; Filename: "{cmd}"; Parameters: "/c start "" ""%ProgramData%\BloatGuard"""; IconFilename: "{app}\{#AppExe}"
Name: "{commonstartup}\BloatGuard Agent"; Filename: "{app}\{#AgentExe}"; Tasks: autostart

[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{commonappdata}\BloatGuard"
