@echo off
setlocal

set SCRIPT_DIR=%~dp0
set PROJECT_DIR=%SCRIPT_DIR%..

echo === Building generador-diplomas for Windows ===

cd /d "%PROJECT_DIR%"

echo --- Installing dependencies ---
call uv sync --group dev

echo --- Running PyInstaller ---
call uv run --group dev pyinstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --name generador-diplomas ^
    --distpath "%PROJECT_DIR%\dist" ^
    --workpath "%PROJECT_DIR%\build\build-windows" ^
    --specpath "%SCRIPT_DIR%" ^
    main.py

echo.
echo === Build complete ===
echo Binary: %PROJECT_DIR%\dist\generador-diplomas\generador-diplomas.exe

endlocal
pause
