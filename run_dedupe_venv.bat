@echo off
setlocal enabledelayedexpansion
title SKU Deduper (venv)

echo ==========================================
echo SKU Deduper (venv) - Keep Highest Price
echo ==========================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "VENV_DIR=%SCRIPT_DIR%venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

if not exist "%SCRIPT_DIR%dedupe_sku_keep_max_price.py" (
  echo [ERROR] Missing: dedupe_sku_keep_max_price.py
  pause
  exit /b 1
)

if not exist "%SCRIPT_DIR%requirements.txt" (
  echo [ERROR] Missing: requirements.txt
  pause
  exit /b 1
)

REM Find a working python.exe (avoid py launcher issues)
set "SYS_PY="
for /f "delims=" %%i in ('where python 2^>nul') do (
  set "SYS_PY=%%i"
  goto :HAVE_PY
)

echo [WARN] "python" not found in PATH.
set /p "SYS_PY=Paste full path to python.exe (example: C:\Python312\python.exe): "
if "%SYS_PY%"=="" (
  echo [ERROR] No python.exe provided.
  pause
  exit /b 1
)

:HAVE_PY
if not exist "%SYS_PY%" (
  echo [ERROR] python.exe not found: "%SYS_PY%"
  pause
  exit /b 1
)

echo [INFO] Using python: "%SYS_PY%"
echo.

REM Create venv if missing
if not exist "%VENV_PY%" (
  echo [INFO] Creating venv: "%VENV_DIR%"
  "%SYS_PY%" -m venv "%VENV_DIR%"
  if %errorlevel% neq 0 (
    echo [ERROR] Failed to create venv.
    pause
    exit /b 1
  )
)

REM Install deps inside venv
echo [INFO] Installing/updating requirements...
"%VENV_PY%" -m pip install --upgrade pip
if %errorlevel% neq 0 (
  echo [ERROR] pip upgrade failed.
  pause
  exit /b 1
)

"%VENV_PY%" -m pip install -r "%SCRIPT_DIR%requirements.txt"
if %errorlevel% neq 0 (
  echo [ERROR] pip install failed.
  pause
  exit /b 1
)

echo.
set /p "INPUT=Enter Excel filename or full path (e.g. products.xlsx): "
if "%INPUT%"=="" (
  echo [ERROR] No file entered.
  pause
  exit /b 1
)

echo.
echo [INFO] Running...
"%VENV_PY%" "%SCRIPT_DIR%dedupe_sku_keep_max_price.py" "%INPUT%"

echo.
echo [INFO] Finished with exit code: %errorlevel%
pause
exit /b %errorlevel%
