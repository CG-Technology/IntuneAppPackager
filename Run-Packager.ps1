<#
.SYNOPSIS
    CG Technology Intune App Packager - Unified Launcher
.DESCRIPTION
    Launches the WPF Packager Studio or runs headless CLI mode if parameters are supplied.
#>

param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$SourceFolder,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$SetupFile,

    [Parameter(Mandatory = $false, Position = 2)]
    [string]$OutputFolder = ".\output",

    [Parameter(Mandatory = $false)]
    [switch]$SkipPackaging,

    [Parameter(Mandatory = $false)]
    [switch]$CLI
)

$scriptDir = $PSScriptRoot

if ($SourceFolder -and $SetupFile) {
    # Run CLI mode
    & "$scriptDir\src\IntunePackager.ps1" -SourceFolder $SourceFolder -SetupFile $SetupFile -OutputFolder $OutputFolder -SkipPackaging:$SkipPackaging
} elseif ($CLI) {
    Get-Help "$scriptDir\src\IntunePackager.ps1" -Detailed
} else {
    # Launch Desktop GUI
    & "$scriptDir\src\Run-PackagerGUI.ps1"
}
