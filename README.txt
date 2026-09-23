GAMEVAULT - REAL WINDOWS INSTALLER
===================================

1. Put this whole folder somewhere, for example:
   F:\GameVault\InstallerBuild

2. Double-click:
   build_GameVault_Installer.bat

3. The script will:
   - install/check Python dependencies
   - extract the built-in GameVault icon
   - build GameVault as a proper PyInstaller ONEDIR app
   - find your Inno Setup installation
   - create the installer

4. Final installer:
   installer\GameVault_Setup.exe

The installed app goes to:
   C:\Program Files\GameVault

The app's user data is kept under:
   %APPDATA%\GameVault

This build deliberately uses ONEDIR instead of ONEFILE. That avoids the
self-modifying PyInstaller archive problem that caused the previous EXE
corruption.

The app already contains the shortcut/icon synchronization logic: changing
the icon from inside GameVault updates/recreates the normal Desktop and
Start Menu shortcuts and attempts to update existing taskbar pins.
