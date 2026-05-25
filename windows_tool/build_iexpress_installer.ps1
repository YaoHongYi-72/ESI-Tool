param(
    [Parameter(Mandatory = $true)]
    [string]$ExePath,

    [string]$ReadmePath = "README.md",
    [string]$InstallCmdPath = "windows_tool/install.cmd",
    [string]$ReleaseDir = "release",
    [string]$AsciiSetupName = "ESI-Tool-Setup.exe",
    [string]$FinalSetupName = "ESI-Tool-Setup.exe"
)

$ErrorActionPreference = "Stop"

function Resolve-RequiredPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (!(Test-Path $PathValue)) {
        throw "$Label not found: $PathValue"
    }

    return (Resolve-Path $PathValue).Path
}

$iexpressPath = Join-Path $env:SystemRoot "System32\iexpress.exe"
if (!(Test-Path $iexpressPath)) {
    throw "IExpress not found: $iexpressPath"
}

$exeResolved = Resolve-RequiredPath -PathValue $ExePath -Label "Application executable"
$readmeResolved = Resolve-RequiredPath -PathValue $ReadmePath -Label "README"
$installResolved = Resolve-RequiredPath -PathValue $InstallCmdPath -Label "Installer command script"

New-Item -ItemType Directory -Force -Path $ReleaseDir | Out-Null
$releaseResolved = (Resolve-Path $ReleaseDir).Path
$payloadDir = Join-Path $releaseResolved "installer_payload"
New-Item -ItemType Directory -Force -Path $payloadDir | Out-Null

Copy-Item -Force $exeResolved (Join-Path $payloadDir "ESI-Tool.exe")
Copy-Item -Force $readmeResolved (Join-Path $payloadDir "README.md")
Copy-Item -Force $installResolved (Join-Path $payloadDir "install.cmd")

$targetAscii = Join-Path $releaseResolved $AsciiSetupName
$targetFinal = Join-Path $releaseResolved $FinalSetupName
$sedPath = Join-Path $releaseResolved "installer.sed"
$payloadSource = "$payloadDir\"

$sedLines = @(
    "[Version]"
    "Class=IEXPRESS"
    "SEDVersion=3"
    "[Options]"
    "PackagePurpose=InstallApp"
    "ShowInstallProgramWindow=0"
    "HideExtractAnimation=1"
    "UseLongFileName=1"
    "InsideCompressed=1"
    "CAB_FixedSize=0"
    "CAB_ResvCodeSigning=0"
    "RebootMode=I"
    "InstallPrompt=%InstallPrompt%"
    "DisplayLicense=%DisplayLicense%"
    "FinishMessage=%FinishMessage%"
    "TargetName=%TargetName%"
    "FriendlyName=%FriendlyName%"
    "AppLaunched=%AppLaunched%"
    "PostInstallCmd=%PostInstallCmd%"
    "AdminQuietInstCmd=%AdminQuietInstCmd%"
    "UserQuietInstCmd=%UserQuietInstCmd%"
    "SourceFiles=SourceFiles"
    "Strings=Strings"
    "[Strings]"
    "InstallPrompt="
    "DisplayLicense="
    "FinishMessage="
    "TargetName=$targetAscii"
    'FriendlyName="ESI Tool Installer"'
    'AppLaunched=cmd.exe /d /s /c ""install.cmd""'
    "PostInstallCmd=<None>"
    "AdminQuietInstCmd="
    "UserQuietInstCmd="
    'FILE0="install.cmd"'
    'FILE1="ESI-Tool.exe"'
    'FILE2="README.md"'
    "[SourceFiles]"
    "SourceFiles0=$payloadSource"
    "[SourceFiles0]"
    "%FILE0%="
    "%FILE1%="
    "%FILE2%="
)

Set-Content -Path $sedPath -Value $sedLines -Encoding Ascii

if (Test-Path $targetAscii) {
    Remove-Item -Force $targetAscii
}
if (Test-Path $targetFinal) {
    Remove-Item -Force $targetFinal
}

& $iexpressPath /N /Q $sedPath
$iexpressExit = $LASTEXITCODE
Write-Host "IExpress exit code: $iexpressExit"

if (!(Test-Path $targetAscii)) {
    Write-Host "Generated SED file:"
    Get-Content -Path $sedPath | ForEach-Object { Write-Host $_ }
    Write-Host "Release directory contents:"
    Get-ChildItem -Force -Path $releaseResolved | ForEach-Object { Write-Host $_.FullName }
    throw "IExpress did not produce installer: $targetAscii"
}

if ($targetAscii -ne $targetFinal) {
    Move-Item -Force $targetAscii $targetFinal
}
Write-Host "Installer created: $targetFinal"
