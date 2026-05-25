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

endlocal
