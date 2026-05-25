@echo off
setlocal

cd /d %~dp0\..

if not exist .venv (
  py -3 -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r windows_tool\requirements-windows.txt

pyinstaller --noconfirm --clean --onefile --windowed --name "ESI统计工具" --paths . ^
  --hidden-import windows_tool.esi_pipeline ^
  --hidden-import analyze_2025_esi ^
  --hidden-import generate_2025_esi_doc ^
  --hidden-import generate_2025_esi_process_doc ^
  --hidden-import generate_author_flow_doc ^
  --hidden-import generate_esi_stats_workbook ^
  windows_tool\esi_gui_app.py

set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe

if not exist "%ISCC%" (
  echo 未找到 Inno Setup 6，请先安装后再运行此脚本。
  exit /b 1
)

pushd windows_tool
"%ISCC%" esi_installer.iss
popd

endlocal
