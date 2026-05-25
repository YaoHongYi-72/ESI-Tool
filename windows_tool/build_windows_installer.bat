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

set SEVENZIP=%ProgramFiles%\7-Zip\7z.exe
if not exist "%SEVENZIP%" set SEVENZIP=%ProgramFiles(x86)%\7-Zip\7z.exe
set SFXMODULE=%ProgramFiles%\7-Zip\7zS.sfx
if not exist "%SFXMODULE%" set SFXMODULE=%ProgramFiles(x86)%\7-Zip\7zS.sfx

if not exist "%SEVENZIP%" (
  echo 未找到 7-Zip，请先安装后再运行此脚本。
  exit /b 1
)

if not exist "%SFXMODULE%" (
  echo 未找到 7zS.sfx，请确认 7-Zip 安装完整。
  exit /b 1
)

if not exist release mkdir release
if not exist release\installer_payload mkdir release\installer_payload

copy /Y "dist\ESI统计工具.exe" "release\installer_payload\ESI统计工具.exe" >nul
copy /Y "README.md" "release\installer_payload\README.md" >nul
copy /Y "windows_tool\install.cmd" "release\installer_payload\install.cmd" >nul

"%SEVENZIP%" a -t7z "release\installer_payload.7z" "release\installer_payload\*" >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$text = Get-Content 'windows_tool/sfx_config.txt' -Raw -Encoding utf8; " ^
  "[System.IO.File]::WriteAllText('release/sfx_config.txt', $text, [System.Text.UTF8Encoding]::new($true)); " ^
  "$sfx = [System.IO.File]::ReadAllBytes('%SFXMODULE%'); " ^
  "$cfg = [System.IO.File]::ReadAllBytes('release/sfx_config.txt'); " ^
  "$arc = [System.IO.File]::ReadAllBytes('release/installer_payload.7z'); " ^
  "$out = [System.IO.File]::Create('release/ESI统计工具-Setup.exe'); " ^
  "try { $out.Write($sfx,0,$sfx.Length); $out.Write($cfg,0,$cfg.Length); $out.Write($arc,0,$arc.Length) } finally { $out.Dispose() }"

endlocal
