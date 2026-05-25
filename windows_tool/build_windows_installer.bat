@echo off
setlocal

cd /d %~dp0\..

if not exist .venv (
  py -3 -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r windows_tool\requirements-windows.txt

pyinstaller --noconfirm --clean --onefile --windowed --name "ESI-Tool" --paths . ^
  --hidden-import windows_tool.esi_pipeline ^
  --hidden-import analyze_2025_esi ^
  --hidden-import generate_2025_esi_doc ^
  --hidden-import generate_2025_esi_process_doc ^
  --hidden-import generate_author_flow_doc ^
  --hidden-import generate_esi_stats_workbook ^
  windows_tool\esi_gui_app.py
if errorlevel 1 exit /b %errorlevel%

set "MAKENSIS=%ProgramFiles(x86)%\NSIS\makensis.exe"
if not exist "%MAKENSIS%" set "MAKENSIS=%ProgramFiles%\NSIS\makensis.exe"

if not exist "%MAKENSIS%" (
  echo 未找到 NSIS，请先安装 NSIS 后再运行此脚本。
  exit /b 1
)

"%MAKENSIS%" /V4 windows_tool\esi_installer.nsi
if errorlevel 1 exit /b %errorlevel%

endlocal
