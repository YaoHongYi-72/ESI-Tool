@echo off
setlocal

cd /d %~dp0\..

if not exist .venv (
  py -3 -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r windows_tool\requirements-windows.txt

pyinstaller --noconfirm windows_tool\esi_tool.spec

set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe

if not exist "%ISCC%" (
  echo 未找到 Inno Setup 6，请先安装后再运行此脚本。
  exit /b 1
)

"%ISCC%" windows_tool\esi_installer.iss

endlocal
