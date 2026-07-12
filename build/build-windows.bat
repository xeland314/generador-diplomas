@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%.."
set PROJECT_DIR=%CD%

echo === Building generador-diplomas for Windows ===

echo --- Installing dependencies ---
uv sync --group dev

echo --- Running PyInstaller ---
uv run --group dev pyinstaller --noconfirm --clean --windowed --name generador-diplomas --distpath "%PROJECT_DIR%\dist" --workpath "%PROJECT_DIR%\build\build-windows" --specpath "%PROJECT_DIR%\build" main.py

echo.
echo === Build complete ===
echo Binary: %PROJECT_DIR%\dist\generador-diplomas\generador-diplomas.exe

endlocal
pause
