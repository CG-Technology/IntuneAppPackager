# Intune App Packager

[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-blue.svg)](https://www.microsoft.com/windows)
[![Architecture](https://img.shields.io/badge/engine-PowerShell%20%7C%20WPF-indigo.svg)](https://learn.microsoft.com/powershell/)
[![Packaging](https://img.shields.io/badge/intunewin-Win32%20Prep%20Tool-purple.svg)](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Organization](https://img.shields.io/badge/org-CG%20Technology-slate.svg)](https://github.com/CG-Technology)

An automated Win32 (`.intunewin`) packaging engine and dynamic PowerShell detection script generator engineered for systems administrators, enterprise IT teams, and Managed Service Providers (MSPs) deploying software via Microsoft Intune.

---

## Key Features

- **Automated `.intunewin` Compilation**: Automatically wraps Microsoft's Win32 Content Prep Tool (`IntuneWinAppUtil.exe`) into an intuitive GUI and headless batch CLI workflow.
- **Installer Engine Inspection**:
  - **MSI Installers**: Automatically queries the Windows Installer COM database to extract `ProductName`, `ProductVersion`, `Manufacturer`, and `ProductCode` (GUID).
  - **EXE Installers**: Inspects PE headers and binary signatures to detect installer types:
    - **Inno Setup**: Recommends `/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-`
    - **Nullsoft (NSIS)**: Recommends `/S`
    - **WiX Burn**: Recommends `/quiet /norestart`
    - **InstallShield**: Recommends `/s /v"/qn /norestart"`
- **Production-Ready Detection Script Generator**: Auto-generates clean PowerShell detection rules (`Detect-<AppName>.ps1`) testing 64-bit and 32-bit (WOW6432Node) uninstall registries with standard Intune exit codes (`exit 0` = Installed, `exit 1` = Not Found).
- **Intune Admin Center Deployment Guide**: Produces a summary file (`Intune-Deployment-Summary.txt`) with copy-paste values for Program Settings, Detection Rules, Return Codes, and Requirements.
- **Dual-Mode (GUI & CLI)**: Interactive WPF Desktop Studio or scriptable PowerShell engine for automated CI/CD and bulk packaging pipelines.

---

## Quick Start

### 1. Launch Interactive GUI Studio
```powershell
.\Run-Packager.ps1
```

1. Select your **Source Folder** (containing installer assets).
2. Choose your **Setup File** (`.exe` or `.msi`).
3. Set your **Output Folder**.
4. Click **Package & Generate Detection**.

### 2. Headless CLI Packaging
```powershell
# Full packaging (.intunewin compilation + detection script + deployment guide)
.\Run-Packager.ps1 -SourceFolder "C:\Installers\Zoom" -SetupFile "ZoomInstaller.exe" -OutputFolder "C:\IntunePackages"

# Fast detection rule generation only (skipping .intunewin build)
.\Run-Packager.ps1 -SourceFolder "C:\Installers\App" -SetupFile "Setup.msi" -SkipPackaging
```

---

## Output Artifacts

For each packaged application, Intune App Packager produces 3 production-ready deliverables:

| File | Purpose |
|------|---------|
| `<AppName>.intunewin` | The encrypted Win32 payload ready to upload to Microsoft Intune Admin Center. |
| `Detect-<AppName>.ps1` | Standalone PowerShell detection script matching DisplayName, DisplayVersion, or MSI GUID. |
| `Intune-Deployment-Summary-<AppName>.txt` | Step-by-step summary containing exact install/uninstall parameters, detection rules, and return codes. |

---

## Generated Detection Script Example

```powershell
# 1. Inspect Windows Uninstall Registry Keys (64-bit and 32-bit WOW64)
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($regPath in $registryPaths) {
    $item = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -like "*Zoom*"
    }
    if ($item) {
        Write-Output "Installed: $($item.DisplayName) (v$($item.DisplayVersion))"
        exit 0
    }
}

# Not detected
exit 1
```

---

## Requirements

- **Operating System**: Windows 10 or Windows 11 (x64)
- **PowerShell**: Windows PowerShell 5.1 or PowerShell 7+
- **Prerequisites**: Zero external runtime dependencies. Automatically acquires Microsoft's `IntuneWinAppUtil.exe` if not already installed.

---

## License

Released under the [MIT License](LICENSE). Copyright &copy; 2026 CG Technology.
