@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
cd /d "%ROOT%" || (
  echo ERROR: Failed to switch to script directory.
  exit /b 1
)

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py -3"
if not defined PYEXE (
  where python >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
  echo ERROR: Python launcher not found. Install Python first.
  exit /b 1
)

if not exist "cioms_to_feedback_excel.py" (
  echo ERROR: cioms_to_feedback_excel.py not found in current folder.
  exit /b 1
)

echo [1/4] Checking Python...
%PYEXE% --version
if errorlevel 1 (
  echo ERROR: Python is not working.
  exit /b 1
)

echo [2/4] Installing build dependencies...
%PYEXE% -m pip install --upgrade pip
if errorlevel 1 (
  echo ERROR: Failed to upgrade pip.
  exit /b 1
)
%PYEXE% -m pip install --upgrade pyinstaller pyyaml openpyxl pypdf requests
if errorlevel 1 (
  echo ERROR: Failed to install dependencies.
  exit /b 1
)

echo [3/4] Building EXE...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "cioms_to_feedback_excel.spec" del /f /q "cioms_to_feedback_excel.spec"

%PYEXE% -m PyInstaller --noconfirm --clean --onefile --name cioms_to_feedback_excel --add-data "cioms_field_mapping.json;." --add-data "cioms_field_mapping.yaml;." "cioms_to_feedback_excel.py"
if errorlevel 1 (
  echo ERROR: PyInstaller build failed.
  exit /b 1
)

echo [4/4] Done.
echo EXE: "%ROOT%dist\cioms_to_feedback_excel.exe"
echo.
echo Usage:
echo   1) Put template xlsx and pdf files next to EXE.
echo   2) Run EXE for batch processing.

endlocal
