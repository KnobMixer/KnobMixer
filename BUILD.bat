@echo off
title KnobMixer Builder
color 0A
echo.
echo  ================================================
echo   KnobMixer Builder
echo  ================================================
echo.

:: ── Clean old builds first ───────────────────────────────────────────────────
echo  Cleaning old build files...
if exist dist          rmdir /s /q dist
if exist build         rmdir /s /q build
if exist dist_installer rmdir /s /q dist_installer
if exist KnobMixer.spec del /q KnobMixer.spec
if exist icon.ico      del /q icon.ico
echo  [OK] Clean.

:: ── Check Python ─────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [!] Python not found. Download from https://python.org
    echo      Make sure to tick "Add Python to PATH" during install.
    pause & exit /b 1
)
echo  [OK] Python found.

:: Extract app version from knob_mixer.py so release version only lives in one place
set "APP_VER="
for /f "tokens=3" %%V in ('findstr /B "APP_VER" knob_mixer.py') do set "APP_VER=%%~V"
if "%APP_VER%"=="" (
    echo  [!] Could not read APP_VER from knob_mixer.py
    pause & exit /b 1
)
echo  [OK] Version detected: %APP_VER%

:: ── Install packages ─────────────────────────────────────────────────────────
echo  Installing packages...
python -m pip install pyinstaller pycaw comtypes keyboard pystray Pillow psutil -q
if errorlevel 1 ( echo  [!] Install failed. & pause & exit /b 1 )
echo  [OK] Packages ready.

:: ── Generate icon ────────────────────────────────────────────────────────────
echo  Generating icon...
python make_icon.py
if errorlevel 1 ( echo  [!] Icon failed. & pause & exit /b 1 )

:: ── Build exe ────────────────────────────────────────────────────────────────
echo.
echo  Building KnobMixer.exe (1-2 minutes)...
python -m PyInstaller --name KnobMixer --onedir --windowed --icon=icon.ico ^
    --add-data "icon.ico;." ^
    --hidden-import pycaw --hidden-import pycaw.magic ^
    --hidden-import comtypes --hidden-import comtypes.client ^
    --hidden-import keyboard --hidden-import pystray ^
    --hidden-import PIL --hidden-import psutil --hidden-import winsound ^
    --noconfirm knob_mixer.py

if errorlevel 1 (
    echo  [!] Build failed. See error above.
    pause & exit /b 1
)
echo  [OK] EXE ready in dist\KnobMixer\

:: ── Build installer ──────────────────────────────────────────────────────────
set ISCC=""
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if %ISCC%=="" (
    echo.
    echo  Inno Setup not found — skipping installer.
    echo  Your app is at:  dist\KnobMixer\KnobMixer.exe
    pause & exit /b 0
)

echo  Building installer...
mkdir dist_installer 2>nul
%ISCC% /DMyAppVersion=%APP_VER% KnobMixer.iss
if errorlevel 1 ( echo  [!] Installer failed. & pause & exit /b 1 )

echo.
echo  ================================================
echo   DONE!
echo   Installer: dist_installer\KnobMixer_Setup.exe
echo  ================================================
pause
