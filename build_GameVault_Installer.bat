@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ==========================================
echo        GameVault Installer Builder
echo ==========================================
echo.

REM Find Python
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY (
    where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo ERROR: Python was not found in PATH.
    echo Install Python 3.11+ and enable "Add Python to PATH".
    pause
    exit /b 1
)

if not exist "VERSION" (
    echo ERROR: VERSION file not found next to this BAT.
    echo Create a file named VERSION containing something like 1.0.0
    pause
    exit /b 1
)

echo [1/5] Checking required Python packages...
%PY% -m pip install flask requests pywebview psutil icoextract pillow pywin32 pyinstaller
if errorlevel 1 (
    echo.
    echo ERROR: Could not install Python packages.
    pause
    exit /b 1
)

echo.
echo [2/5] Creating application icon...
%PY% make_icon.py
if errorlevel 1 (
    echo ERROR: Could not create icon.ico
    pause
    exit /b 1
)

echo.
echo [3/5] Building GameVault with PyInstaller (onedir)...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

%PY% -m PyInstaller --noconfirm --clean --onedir --windowed ^
 --name GameVault ^
 --icon=icon.ico ^
 --add-data "VERSION;." ^
 --hidden-import=win32com.client ^
 --hidden-import=pythoncom ^
 --hidden-import=pywintypes ^
 --collect-all webview ^
 GameVault_app.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    echo.
    pause
    exit /b 1
)

if not exist "dist\GameVault\GameVault.exe" (
    echo ERROR: GameVault.exe was not created.
    pause
    exit /b 1
)

echo.
echo [4/5] Finding Inno Setup compiler...
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC (
    where ISCC.exe >nul 2>&1 && set "ISCC=ISCC.exe"
)

if not defined ISCC (
    echo ERROR: Inno Setup was not found.
    echo.
    echo If you installed it in a custom folder, add its folder to PATH,
    echo or edit this BAT and set ISCC to the full path of ISCC.exe.
    pause
    exit /b 1
)

echo Found: %ISCC%

REM The version comes from the VERSION file (same one bundled into the app
REM and used by the GitHub workflow), so the installer version and the
REM version shown inside GameVault always match. Bump VERSION to ship an update.
set "APPVER="
for /f "usebackq delims=" %%i in ("VERSION") do if not defined APPVER set "APPVER=%%i"
if not defined APPVER (
    echo ERROR: VERSION file is empty.
    pause
    exit /b 1
)
echo Version: %APPVER%

echo.
echo [5/5] Building GameVault_Setup.exe...
if exist installer rmdir /s /q installer
mkdir installer

"%ISCC%" /DMyAppVersion=%APPVER% "GameVault.iss"
if errorlevel 1 (
    echo.
    echo ERROR: Inno Setup failed.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo SUCCESS!
echo Installer:
echo %CD%\installer\GameVault_Setup.exe
echo ==========================================
echo.
pause
