; GameVault Inno Setup script
;
; UPDATES: running a newer GameVault_Setup.exe on top of an existing install
; upgrades it in place -- no need to uninstall first. This works because
; AppId below is ALWAYS the same. NEVER change AppId, or Windows will treat
; the new installer as a different program and install it side by side.
;
; The version is normally passed in by the build script (/DMyAppVersion=...),
; so every build gets a new version number automatically.
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppName "GameVault"
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
; On an update, reuse the folder of the existing install and skip the folder page.
UsePreviousAppDir=yes
DisableDirPage=auto
; We close GameVault ourselves (see PrepareToInstall below), so skip the
; "applications are using files" dialog.
CloseApplications=no
RestartApplications=no

[InstallDelete]
; PyInstaller's runtime folder is wiped before copying the new one, so files
; that no longer exist in the new build don't pile up between versions.
; User data lives in %APPDATA%\GameVault and is NOT touched.
Type: filesandordirs; Name: "{app}\_internal"

[Files]
Source: "dist\GameVault\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; GameVault rewrites its own shortcuts when you change the app icon, so an
; update must not put the default icon back (Inno has no "only if missing"
; flag for icons, so a Check function is used instead).
; The Desktop shortcut is only created on a fresh install (not re-created
; on every update if you deleted it on purpose).
Name: "{autodesktop}\GameVault"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"; Check: not IsUpgrade
Name: "{autoprograms}\GameVault"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppExeName}"; Check: StartMenuShortcutMissing

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch GameVault"; Flags: nowait postinstall skipifsilent

[Code]
const
  UninstallKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1';

// Version of the copy that is already installed ('' if GameVault isn't installed yet).
function GetInstalledVersion: String;
var
  Key: String;
begin
  Result := '';
  Key := ExpandConstant(UninstallKey);
  if not RegQueryStringValue(HKLM, Key, 'DisplayVersion', Result) then
    if not RegQueryStringValue(HKCU, Key, 'DisplayVersion', Result) then
      Result := '';
end;

function IsUpgrade: Boolean;
begin
  Result := GetInstalledVersion <> '';
end;

function StartMenuShortcutMissing: Boolean;
begin
  Result := not FileExists(ExpandConstant('{autoprograms}\GameVault.lnk'));
end;

// Make sure GameVault (and its embedded browser child processes) isn't running,
// otherwise Windows keeps its files locked and they can't be replaced.
function IsGameVaultRunning: Boolean;
var
  ResultCode: Integer;
  TmpFile: String;
  Contents: AnsiString;
begin
  Result := False;
  TmpFile := ExpandConstant('{tmp}\gv_check.txt');
  Exec(ExpandConstant('{cmd}'),
       '/C tasklist /FI "IMAGENAME eq {#MyAppExeName}" > "' + TmpFile + '"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if LoadStringFromFile(TmpFile, Contents) then
    Result := Pos(Lowercase('{#MyAppExeName}'), Lowercase(Contents)) > 0;
end;

procedure CloseGameVault;
var
  ResultCode: Integer;
  i: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM {#MyAppExeName}', '',
       SW_HIDE, ewWaitUntilTerminated, ResultCode);
  // Poll instead of trusting a fixed short sleep. This matters most when
  // the update was triggered from *inside* GameVault itself (it has to
  // kill its own running instance, plus its embedded WebView2 browser
  // child processes) -- a flat 800ms wasn't always long enough for every
  // file handle to actually release, so a few files (like VERSION) could
  // silently keep their old content even though the install "succeeded".
  for i := 1 to 20 do
  begin
    if not IsGameVaultRunning then
      break;
    Sleep(300);
  end;
  Sleep(500);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  CloseGameVault;
  Result := '';
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    CloseGameVault;
end;

// Show "Update" wording instead of "Install" when GameVault is already installed.
procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpReady then
  begin
    if IsUpgrade then
    begin
      WizardForm.NextButton.Caption := 'Update';
      WizardForm.ReadyLabel.Caption :=
        'Setup will update GameVault from version ' + GetInstalledVersion +
        ' to version {#MyAppVersion}.' + #13#10 + #13#10 +
        'Your library, categories and settings are kept. ' +
        'GameVault will be closed if it is running.' + #13#10 + #13#10 +
        'Click Update to continue.';
    end;
  end;
end;
