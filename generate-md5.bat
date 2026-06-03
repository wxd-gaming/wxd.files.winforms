@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo.
echo Current directory: %cd%
echo.

if not exist "wxd.files.winforms.exe" (
    echo [ERROR] wxd.files.winforms.exe not found.
    pause
    exit /b 1
)

set CURRENT_VER=
if exist "md5.txt" (
    set /p CURRENT_VER=<md5.txt
)
if defined CURRENT_VER (
    echo Current version: !CURRENT_VER!
)

set /p VERSION=Enter new version (e.g. 1.0.0.5):
if "!VERSION!"=="" (
    echo [ERROR] Version cannot be empty.
    pause
    exit /b 1
)

echo.
echo Computing MD5...
echo.

for /f %%i in ('powershell -NoProfile -Command "(Get-FileHash 'wxd.files.winforms.exe' -Algorithm MD5).Hash.ToLower()"') do set MD5=%%i

> md5.txt echo !VERSION!
>> md5.txt echo !MD5!

echo.
echo ========================
echo  md5.txt generated!
echo  Version: !VERSION!
echo  MD5    : !MD5!
echo ========================
echo.
echo Preview:
echo --------
type md5.txt
echo --------
echo.
pause
