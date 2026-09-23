; GameVault Inno Setup script
#define MyAppName "GameVault"
#define MyAppVersion "1.0.0"
#define MyAppExeName "GameVault.exe"

[Setup]
AppId={{8F4F7E4D-8B1A-4D9B-9F2E-7A7D4A6C9D31}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\GameVault
DefaultGroupName=GameVault
OutputDir=installer
OutputBaseFilename=GameVault_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=icon.ico
DisableProgramGroupPage=yes

[Files]
Source: "dist\GameVault\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autodesktop}\GameVault"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autoprograms}\GameVault"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch GameVault"; Flags: nowait postinstall skipifsilent
