Unicode True
Name "ESI统计工具"
OutFile "..\release\ESI-Tool-Setup.exe"
InstallDir "$LOCALAPPDATA\Programs\ESI统计工具"
RequestExecutionLevel user
ShowInstDetails show
AutoCloseWindow true
XPStyle on
SetCompressor /SOLID lzma

Section "Install"
  SetShellVarContext current
  SetOutPath "$INSTDIR"
  File /oname=ESI-Tool.exe "..\dist\ESI统计工具.exe"
  File "..\README.md"
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  CreateShortcut "$DESKTOP\ESI统计工具.lnk" "$INSTDIR\ESI-Tool.exe"
  CreateShortcut "$SMPROGRAMS\ESI统计工具.lnk" "$INSTDIR\ESI-Tool.exe"

  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ESI统计工具" "DisplayName" "ESI统计工具"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ESI统计工具" "DisplayIcon" "$INSTDIR\ESI-Tool.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ESI统计工具" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ESI统计工具" "QuietUninstallString" "$INSTDIR\Uninstall.exe /S"

  ExecShell "" "$INSTDIR\ESI-Tool.exe"
SectionEnd

Section "Uninstall"
  SetShellVarContext current
  Delete "$DESKTOP\ESI统计工具.lnk"
  Delete "$SMPROGRAMS\ESI统计工具.lnk"
  Delete "$INSTDIR\README.md"
  Delete "$INSTDIR\ESI-Tool.exe"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ESI统计工具"
SectionEnd
