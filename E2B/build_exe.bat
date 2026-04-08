@echo off
setlocal

REM 一键将 cioms_to_feedback_excel.py 打包为 exe（Windows）
REM 使用方法：
REM   1) 双击本文件，或在 cmd 中执行：build_exe.bat
REM   2) 生成文件位于 dist\cioms_to_feedback_excel.exe

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo [1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python not found. Please install Python and add to PATH.
  exit /b 1
)

echo [2/4] Installing build dependencies...
python -m pip install --upgrade pyinstaller pyyaml openpyxl pypdf requests
if errorlevel 1 (
  echo ERROR: Failed to install dependencies.
  exit /b 1
)

echo [3/4] Building exe...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "cioms_to_feedback_excel.spec" del /f /q "cioms_to_feedback_excel.spec"

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --name cioms_to_feedback_excel ^
  --add-data "cioms_field_mapping.json;." ^
  "cioms_to_feedback_excel.py"

if errorlevel 1 (
  echo ERROR: PyInstaller build failed.
  exit /b 1
)

echo [4/4] Done.
echo EXE: "%SCRIPT_DIR%dist\cioms_to_feedback_excel.exe"
echo.
echo Tip:
echo   Put "数据反馈结果模板.xlsx" and PDF files next to the exe,
echo   then run the exe for batch processing.

endlocal
