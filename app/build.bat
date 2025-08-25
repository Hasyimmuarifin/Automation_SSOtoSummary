@echo off
title Build EXE with PyInstaller
echo ==========================================
echo Building EXE from main.py ...
echo ==========================================

REM Masuk ke direktori app
cd /d "%~dp0"

REM Hapus folder build & dist lama
rmdir /s /q build
rmdir /s /q dist
del main.spec

REM Jalankan PyInstaller
python -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --icon=assets\barageng.ico ^
    --name "Automation SSO" ^
    --add-data "assets;assets" ^
    --add-data "style;style" ^
    --add-data "config;config" ^
    --add-data "gui;gui" ^
    --add-data "logic;logic" ^
    --add-data "utils;utils" ^
    main.py

echo ==========================================
echo Build Selesai!
echo File EXE ada di: dist\Automation SSO.exe
echo ==========================================
pause