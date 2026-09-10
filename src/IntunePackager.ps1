<#
.SYNOPSIS
    Intune App Packager - Automated Win32 (.intunewin) packager & detection script generator.
.DESCRIPTION
    Inspects Windows installer binaries (MSI, Inno Setup, NSIS, WiX, InstallShield),
    detects silent install/uninstall switches, compiles encrypted .intunewin packages
    using Microsoft Win32 Content Prep Tool, and generates production-ready PowerShell
    detection scripts for Microsoft Intune.
.PARAMETER SourceFolder
    Directory containing the installer and any associated support files.
.PARAMETER SetupFile
    Filename of the setup executable (e.g. ZoomInstaller.exe or Setup.msi) inside SourceFolder.
.PARAMETER OutputFolder
    Directory where the compiled .intunewin package, detection script, and summary will be saved.
.PARAMETER SkipPackaging
    If specified, inspects installer and generates detection script & summary without compiling .intunewin.
.EXAMPLE
    .\src\IntunePackager.ps1 -SourceFolder "C:\Installers\Zoom" -SetupFile "ZoomInstaller.exe" -OutputFolder "C:\IntunePackages"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$SetupFile,

    [Parameter(Mandatory = $false, Position = 2)]
    [string]$OutputFolder = ".\output",

    [Parameter(Mandatory = $false)]
    [switch]$SkipPackaging,

    [Parameter(Mandatory = $false)]
    [string]$CustomInstallSwitches,

    [Parameter(Mandatory = $false)]
    [string]$CustomUninstallSwitches
)

# Set Strict Mode & Error Handling
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    switch ($Level) {
        "INFO"    { Write-Host "[$timestamp] [INFO] $Message" -ForegroundColor Cyan }
        "SUCCESS" { Write-Host "[$timestamp] [PASS] $Message" -ForegroundColor Green }
        "WARN"    { Write-Host "[$timestamp] [WARN] $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "[$timestamp] [FAIL] $Message" -ForegroundColor Red }
    }
}

function Get-MsiInformation {
    param([string]$FilePath)

    $msiInfo = @{
        Type            = "MSI"
        ProductName     = ""
        ProductVersion  = ""
        Manufacturer    = ""
        ProductCode     = ""
        UpgradeCode     = ""
        InstallerType   = "Windows Installer (MSI)"
        SilentInstall   = "msiexec.exe /i `"$SetupFile`" /qn /norestart"
        SilentUninstall = ""
    }

    try {
        $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $database = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $windowsInstaller, @($FilePath, 0))

        $query = "SELECT Property, Value FROM Property"
        $view = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $database, @($query))
        $null = $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)

        while ($record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)) {
            $prop = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, @(1))
            $val  = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, @(2))

            switch ($prop) {
                "ProductName"    { $msiInfo.ProductName = $val }
                "ProductVersion" { $msiInfo.ProductVersion = $val }
                "Manufacturer"   { $msiInfo.Manufacturer = $val }
                "ProductCode"    { $msiInfo.ProductCode = $val }
                "UpgradeCode"    { $msiInfo.UpgradeCode = $val }
            }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($database) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null
    }
    catch {
        Write-Log "Unable to parse MSI properties through COM: $($_.Exception.Message)" "WARN"
    }

    if ($msiInfo.ProductCode) {
        $msiInfo.SilentUninstall = "msiexec.exe /x $($msiInfo.ProductCode) /qn /norestart"
    } else {
        $msiInfo.SilentUninstall = "msiexec.exe /x `"$SetupFile`" /qn /norestart"
    }

    return $msiInfo
}

function Get-ExeInformation {
    param([string]$FilePath)

    $exeInfo = @{
        Type            = "EXE"
        ProductName     = ""
        ProductVersion  = ""
        Manufacturer    = ""
        ProductCode     = ""
        InstallerType   = "Standard Windows Executable"
        SilentInstall   = ""
        SilentUninstall = ""
    }

    try {
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($FilePath)
        $exeInfo.ProductName = if ($versionInfo.ProductName) { $versionInfo.ProductName.Trim() } else { $versionInfo.FileDescription.Trim() }
        $exeInfo.ProductVersion = if ($versionInfo.ProductVersion) { $versionInfo.ProductVersion.Trim() } else { $versionInfo.FileVersion.Trim() }
        $exeInfo.Manufacturer = if ($versionInfo.CompanyName) { $versionInfo.CompanyName.Trim() } else { "" }
    }
    catch {
        Write-Log "Failed reading FileVersionInfo: $($_.Exception.Message)" "WARN"
    }

    if (-not $exeInfo.ProductName) {
        $exeInfo.ProductName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    }

    # Detect Installer Framework by reading partial binary signatures
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        $readBytes = [Math]::Min($stream.Length, 2097152) # Read up to first 2MB
        $buffer = New-Object byte[] $readBytes
        $null = $stream.Read($buffer, 0, $readBytes)
        $stream.Close()
        $asciiString = [System.Text.Encoding]::ASCII.GetString($buffer)

        if ($asciiString -match "Inno Setup") {
            $exeInfo.InstallerType = "Inno Setup"
            $exeInfo.SilentInstall = "`"$SetupFile`" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-"
            $exeInfo.SilentUninstall = "`"%ProgramFiles%\$($exeInfo.ProductName)\unins000.exe`" /VERYSILENT /NORESTART"
        }
        elseif ($asciiString -match "NullsoftInst" -or $asciiString -match "NSIS") {
            $exeInfo.InstallerType = "Nullsoft Scriptable Install System (NSIS)"
            $exeInfo.SilentInstall = "`"$SetupFile`" /S"
            $exeInfo.SilentUninstall = "`"%ProgramFiles%\$($exeInfo.ProductName)\Uninstall.exe`" /S"
        }
        elseif ($asciiString -match "InstallShield") {
            $exeInfo.InstallerType = "InstallShield"
            $exeInfo.SilentInstall = "`"$SetupFile`" /s /v`"/qn /norestart`""
            $exeInfo.SilentUninstall = "`"$SetupFile`" /uninst /s /v`"/qn /norestart`""
        }
        elseif ($asciiString -match "WixBundle" -or $asciiString -match "WixBurn") {
            $exeInfo.InstallerType = "WiX Burn Chainer"
            $exeInfo.SilentInstall = "`"$SetupFile`" /quiet /norestart"
            $exeInfo.SilentUninstall = "`"$SetupFile`" /uninstall /quiet /norestart"
        }
        elseif ($asciiString -match "7-Zip") {
            $exeInfo.InstallerType = "7-Zip Self-Extracting Archive"
            $exeInfo.SilentInstall = "`"$SetupFile`" -y /gm2"
            $exeInfo.SilentUninstall = ""
        }
        else {
            $exeInfo.InstallerType = "Standard Windows Executable"
            $exeInfo.SilentInstall = "`"$SetupFile`" /S /silent /quiet /norestart"
            $exeInfo.SilentUninstall = ""
        }
    }
    catch {
        Write-Log "Unable to inspect binary headers: $($_.Exception.Message)" "WARN"
        $exeInfo.InstallerType = "Generic Setup"
        $exeInfo.SilentInstall = "`"$SetupFile`" /silent /norestart"
    }

    return $exeInfo
}

function Ensure-ContentPrepTool {
    param([string]$ToolsDir)

    $prepExe = Join-Path $ToolsDir "IntuneWinAppUtil.exe"
    if (Test-Path $prepExe) {
        return $prepExe
    }

    # Check script folder
    $scriptDir = Split-Path -Parent $PSCommandPath
    $prepExeInScript = Join-Path $scriptDir "IntuneWinAppUtil.exe"
    if (Test-Path $prepExeInScript) {
        return $prepExeInScript
    }

    # Check PATH
    $inPath = Get-Command "IntuneWinAppUtil.exe" -ErrorAction SilentlyContinue
    if ($inPath) {
        return $inPath.Source
    }

    # Automated download
    if (-not (Test-Path $ToolsDir)) {
        New-Item -ItemType Directory -Path $ToolsDir -Force | Out-Null
    }

    $downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
    Write-Log "Downloading official Microsoft Intune Content Prep Tool..." "INFO"

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $downloadUrl -OutFile $prepExe -UseBasicParsing
        Write-Log "Successfully acquired IntuneWinAppUtil.exe" "SUCCESS"
        return $prepExe
    }
    catch {
        Write-Log "Failed to download IntuneWinAppUtil.exe: $($_.Exception.Message)" "WARN"
        return $null
    }
}

function New-IntuneDetectionScript {
    param(
        [hashtable]$Info,
        [string]$OutputPath
    )

    $appName = $Info.ProductName
    $appVersion = $Info.ProductVersion
    $productCode = $Info.ProductCode

    $scriptContent = @"
<#
.SYNOPSIS
    Automated Intune Detection Rule for $appName
.DESCRIPTION
    Generated by CG Technology Intune App Packager.
    Standard Intune Exit Codes:
        exit 0 -> Application is installed
        exit 1 -> Application is NOT installed
#>

`$appName    = "$appName"
`$appVersion = "$appVersion"
`$productCode = "$productCode"

# 1. Check Windows Uninstall Registry Keys (64-bit and 32-bit WOW64)
`$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# Direct MSI ProductCode probe if available
if (`$productCode) {
    `$msiPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\`$productCode",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\`$productCode"
    )
    foreach (`$mp in `$msiPaths) {
        if (Test-Path `$mp) {
            Write-Output "Found matching MSI ProductCode: `$productCode"
            exit 0
        }
    }
}

# Inspect all registered installed applications
foreach (`$regPath in `$registryPaths) {
    `$installedItems = Get-ItemProperty `$regPath -ErrorAction SilentlyContinue | Where-Object {
        (`$_.DisplayName -like "*`$appName*") -or (`$_.PSChildName -eq `$productCode)
    }

    foreach (`$item in `$installedItems) {
        if (`$item.DisplayName) {
            if (`$appVersion) {
                # Check version comparison if version is specified
                if (`$item.DisplayVersion -and ([System.Version]::TryParse(`$item.DisplayVersion, [ref]`$null)) -and ([System.Version]::TryParse(`$appVersion, [ref]`$null))) {
                    if ([System.Version]`$item.DisplayVersion -ge [System.Version]`$appVersion) {
                        Write-Output "Installed: `$(`$item.DisplayName) (v`$(`$item.DisplayVersion))"
                        exit 0
                    }
                } else {
                    Write-Output "Installed: `$(`$item.DisplayName)"
                    exit 0
                }
            } else {
                Write-Output "Installed: `$(`$item.DisplayName)"
                exit 0
            }
        }
    }
}

# 2. Application not detected
exit 1
"@

    $scriptContent | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Generated PowerShell detection script: $OutputPath" "SUCCESS"
}

function New-IntuneDeploymentSummary {
    param(
        [hashtable]$Info,
        [string]$IntuneWinFile,
        [string]$DetectionScriptFile,
        [string]$OutputPath
    )

    $summary = @"
================================================================================
CG TECHNOLOGY INTUNE APP PACKAGER - DEPLOYMENT GUIDE
================================================================================
Target Application : $($Info.ProductName)
Detected Version   : $($Info.ProductVersion)
Publisher          : $($Info.Manufacturer)
Installer Engine   : $($Info.InstallerType)
Package Output     : $IntuneWinFile
Detection Script   : $DetectionScriptFile
Generated On       : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
================================================================================

1. APP INFORMATION (Step 1 in Microsoft Intune Admin Center):
   - Name                 : $($Info.ProductName)
   - Description          : Enterprise deployment package for $($Info.ProductName)
   - Publisher            : $(if ($Info.Manufacturer) { $Info.Manufacturer } else { "Vendor" })
   - App Version          : $($Info.ProductVersion)

2. PROGRAM SETTINGS (Step 2):
   - Install Command      : $($Info.SilentInstall)
   - Uninstall Command    : $($Info.SilentUninstall)
   - Install Behavior     : System
   - Device Restart       : Determine behavior based on return codes (or No specific action)

3. REQUIREMENTS (Step 3):
   - Operating System     : 64-bit
   - Minimum OS Version   : Windows 10 21H2 (19044) or Windows 11 22H2+
   - Disk Space Required  : 250 MB

4. DETECTION RULES (Step 4):
   - Rule Format          : Use a custom detection script
   - Script File          : $([System.IO.Path]::GetFileName($DetectionScriptFile))
   - Run Script As 32-bit : No
   - Enforce Signature    : No

5. RETURN CODES (Step 5):
   - 0                    : Success
   - 1707                 : Success
   - 3010                 : Soft Reboot Required
   - 1641                 : Hard Reboot Required
   - 1618                 : Fast Retry (Another installation in progress)

================================================================================
Generated by CG Technology Automation Toolkit (https://cg-technology.github.io)
================================================================================
"@

    $summary | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Generated Intune deployment summary: $OutputPath" "SUCCESS"
}

# ==============================================================================
# Execution Flow
# ==============================================================================

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Magenta
Write-Host " CG Technology Intune App Packager & Detection Script Generator" -ForegroundColor White
Write-Host "==================================================================" -ForegroundColor Magenta
Write-Host ""

# Validate inputs
$fullSourcePath = [System.IO.Path]::GetFullPath($SourceFolder)
if (-not (Test-Path $fullSourcePath)) {
    throw "Source folder does not exist: $fullSourcePath"
}

$fullSetupPath = Join-Path $fullSourcePath $SetupFile
if (-not (Test-Path $fullSetupPath)) {
    throw "Setup file does not exist inside source folder: $fullSetupPath"
}

$fullOutputPath = [System.IO.Path]::GetFullPath($OutputFolder)
if (-not (Test-Path $fullOutputPath)) {
    New-Item -ItemType Directory -Path $fullOutputPath -Force | Out-Null
}

Write-Log "Source Folder : $fullSourcePath" "INFO"
Write-Log "Setup File    : $SetupFile" "INFO"
Write-Log "Output Folder : $fullOutputPath" "INFO"

# Step 1: Inspect Installer
Write-Log "Inspecting setup binary: $SetupFile..." "INFO"
$ext = [System.IO.Path]::GetExtension($SetupFile).ToLower()

$appInfo = if ($ext -eq ".msi") {
    Get-MsiInformation -FilePath $fullSetupPath
} else {
    Get-ExeInformation -FilePath $fullSetupPath
}

if ($appInfo -is [array]) {
    $appInfo = $appInfo | Where-Object { $_ -is [System.Collections.IDictionary] } | Select-Object -Last 1
}

if ($CustomInstallSwitches) {
    $appInfo.SilentInstall = $CustomInstallSwitches
}
if ($CustomUninstallSwitches) {
    $appInfo.SilentUninstall = $CustomUninstallSwitches
}

Write-Log "Detected Application : $($appInfo.ProductName)" "SUCCESS"
Write-Log "Detected Version     : $($appInfo.ProductVersion)" "SUCCESS"
Write-Log "Installer Engine     : $($appInfo.InstallerType)" "SUCCESS"
Write-Log "Silent Install       : $($appInfo.SilentInstall)" "SUCCESS"
Write-Log "Silent Uninstall     : $($appInfo.SilentUninstall)" "SUCCESS"

# Step 2: Generate Detection Script
$cleanAppName = ($appInfo.ProductName -replace '[^a-zA-Z0-9_\-]', '')
if (-not $cleanAppName) { $cleanAppName = "App" }
$detectionScriptName = "Detect-$cleanAppName.ps1"
$detectionScriptPath = Join-Path $fullOutputPath $detectionScriptName
New-IntuneDetectionScript -Info $appInfo -OutputPath $detectionScriptPath

# Step 3: Package .intunewin
$intuneWinFileName = "$([System.IO.Path]::GetFileNameWithoutExtension($SetupFile)).intunewin"
$intuneWinPath = Join-Path $fullOutputPath $intuneWinFileName

if ($SkipPackaging) {
    Write-Log "SkipPackaging flag supplied. Skipping .intunewin compilation." "WARN"
} else {
    $toolsDir = Join-Path $fullOutputPath "tools"
    $prepTool = Ensure-ContentPrepTool -ToolsDir $toolsDir

    if ($prepTool -and (Test-Path $prepTool)) {
        Write-Log "Invoking Microsoft Content Prep Tool ($prepTool)..." "INFO"
        $prepArgs = @(
            "-c", "`"$fullSourcePath`"",
            "-s", "`"$SetupFile`"",
            "-o", "`"$fullOutputPath`"",
            "-q"
        )

        $proc = Start-Process -FilePath $prepTool -ArgumentList ($prepArgs -join " ") -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0 -and (Test-Path $intuneWinPath)) {
            $pkgSize = (Get-Item $intuneWinPath).Length
            $pkgSizeMB = [Math]::Round($pkgSize / 1MB, 2)
            Write-Log "Successfully compiled Intune Win32 package: $intuneWinPath ($pkgSizeMB MB)" "SUCCESS"
        } else {
            Write-Log "Content Prep Tool exited with code: $($proc.ExitCode)" "WARN"
        }
    } else {
        Write-Log "IntuneWinAppUtil.exe not found. .intunewin compilation skipped." "WARN"
    }
}

# Step 4: Generate Summary File
$summaryPath = Join-Path $fullOutputPath "Intune-Deployment-Summary-$cleanAppName.txt"
New-IntuneDeploymentSummary -Info $appInfo -IntuneWinFile $intuneWinPath -DetectionScriptFile $detectionScriptPath -OutputPath $summaryPath

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Green
Write-Host " Packaging Workflow Completed Successfully!" -ForegroundColor Green
Write-Host " Output Directory: $fullOutputPath" -ForegroundColor Yellow
Write-Host "==================================================================" -ForegroundColor Green
Write-Host ""
