@echo off
setlocal

set "APP_NAME=ESI统计工具"
set "APP_EXE=ESI-Tool.exe"
set "INSTALL_DIR=%LOCALAPPDATA%\Programs\%APP_NAME%"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
set "DESKTOP_DIR=%USERPROFILE%\Desktop"

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

copy /Y "%APP_EXE%" "%INSTALL_DIR%\%APP_EXE%" >nul
copy /Y "README.md" "%INSTALL_DIR%\README.md" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell; " ^
  "$target = Join-Path $env:LOCALAPPDATA 'Programs\ESI统计工具\ESI-Tool.exe'; " ^
  "$workdir = Split-Path $target; " ^
  "$desktop = Join-Path $env:USERPROFILE 'Desktop\ESI统计工具.lnk'; " ^
  "$menu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\ESI统计工具.lnk'; " ^
  "$s1 = $ws.CreateShortcut($desktop); $s1.TargetPath = $target; $s1.WorkingDirectory = $workdir; $s1.Save(); " ^
  "$s2 = $ws.CreateShortcut($menu); $s2.TargetPath = $target; $s2.WorkingDirectory = $workdir; $s2.Save();"

start "" "%INSTALL_DIR%\%APP_EXE%"

endlocal
