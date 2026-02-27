@echo off
title Serial Data Logger — EXE Builder
echo =========================================
echo  Serial Data Logger — EXE Builder
echo =========================================
echo.

REM ── 1. Check Python ──────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install from https://www.python.org/downloads/
    echo         Make sure "Add Python to PATH" is ticked during install.
    pause & exit /b 1
)

REM ── 2. Install / upgrade dependencies ────────────────────────────────────
echo [1/3] Installing dependencies...
python -m pip install --upgrade pip --quiet
python -m pip install pyinstaller pyserial openpyxl --quiet
if errorlevel 1 (
    echo [ERROR] pip install failed. Check your internet connection.
    pause & exit /b 1
)

REM ── 3. Build the exe ──────────────────────────────────────────────────────
echo [2/3] Building executable (this takes ~30 seconds)...

REM Put serial_logger.py in the same folder as this script, then run it.
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "SerialDataLogger" ^
    --add-data "%~dp0serial_logger.py;." ^
    "%~dp0serial_logger.py.py"

if errorlevel 1 (
    echo [ERROR] Build failed — see output above for details.
    pause & exit /b 1
)

REM ── 4. Done ───────────────────────────────────────────────────────────────
echo [3/3] Done!
echo.
echo  Your executable is here:
echo  %~dp0dist\SerialDataLogger.exe
echo.
echo  You can copy SerialDataLogger.exe anywhere and share it.
echo  The recipient does NOT need Python installed.
echo.
pause
