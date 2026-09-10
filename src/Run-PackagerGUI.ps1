<#
.SYNOPSIS
    CG Technology Intune App Packager - Interactive Desktop GUI
.DESCRIPTION
    WPF Graphical interface for packaging Win32 applications and auto-generating
    detection scripts for Microsoft Intune.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CG Technology - Intune App Packager &amp; Detection Studio"
        Height="680" Width="840" MinHeight="600" MinWidth="750"
        WindowStartupLocation="CenterScreen"
        Background="#0B0F19" Foreground="#F8FAFC" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#94A3B8"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#111827"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Margin="0,0,0,20" Padding="0,0,0,16" BorderBrush="#1E293B" BorderThickness="0,0,0,1">
            <StackPanel>
                <TextBlock Text="INTUNE APP PACKAGER" FontSize="11" FontWeight="Bold" Foreground="#818CF8" Margin="0,0,0,4"/>
                <TextBlock Text="Win32 Package (.intunewin) &amp; Detection Script Generator" FontSize="20" FontWeight="Bold" Foreground="#FFFFFF"/>
                <TextBlock Text="Inspect setup files, detect silent switches, and bundle ready-to-upload Intune applications." FontSize="13" Foreground="#94A3B8" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Form Fields -->
        <Border Grid.Row="1" Background="#0F172A" BorderBrush="#1E293B" BorderThickness="1" CornerRadius="8" Padding="18" Margin="0,0,0,16">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="140"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="95"/>
                </Grid.ColumnDefinitions>

                <!-- Source Folder -->
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Source Folder:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,0,12"/>
                <TextBox Grid.Row="0" Grid.Column="1" x:Name="TxtSourceFolder" Margin="0,0,10,12" VerticalAlignment="Center"/>
                <Button Grid.Row="0" Grid.Column="2" x:Name="BtnBrowseSource" Content="Browse..." Margin="0,0,0,12"/>

                <!-- Setup File -->
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Setup File:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,0,12"/>
                <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtSetupFile" Margin="0,0,10,12" VerticalAlignment="Center" ToolTip="Filename of installer inside Source Folder"/>
                <Button Grid.Row="1" Grid.Column="2" x:Name="BtnBrowseSetup" Content="Select..." Margin="0,0,0,12"/>

                <!-- Output Folder -->
                <TextBlock Grid.Row="2" Grid.Column="0" Text="Output Folder:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,0,12"/>
                <TextBox Grid.Row="2" Grid.Column="1" x:Name="TxtOutputFolder" Margin="0,0,10,12" VerticalAlignment="Center"/>
                <Button Grid.Row="2" Grid.Column="2" x:Name="BtnBrowseOutput" Content="Browse..." Margin="0,0,0,12"/>

                <!-- Action Strip -->
                <StackPanel Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0">
                    <CheckBox x:Name="ChkSkipPackaging" Content="Generate Detection Script Only (Skip .intunewin)" Foreground="#94A3B8" VerticalAlignment="Center" Margin="0,0,16,0"/>
                    <Button x:Name="BtnStartPackaging" Content="Package &amp; Generate Detection" Background="#4F46E5" Foreground="#FFFFFF" BorderThickness="0" Padding="18,8" FontSize="14" FontWeight="Bold"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Log Output Console -->
        <Border Grid.Row="2" Background="#060911" BorderBrush="#1E293B" BorderThickness="1" CornerRadius="8" Padding="12">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Text="OUTPUT &amp; DIAGNOSTIC CONSOLE" FontSize="11" FontWeight="Bold" Foreground="#64748B" Margin="0,0,0,8"/>
                <TextBox Grid.Row="1" x:Name="TxtConsoleLog" Background="Transparent" Foreground="#E2E8F0" FontFamily="Consolas, Cascadia Code, Courier New" FontSize="12" IsReadOnly="True" VerticalScrollBarVisibility="Auto" BorderThickness="0" TextWrapping="Wrap"/>
            </Grid>
        </Border>

        <!-- Footer -->
        <Grid Grid.Row="3" Margin="0,16,0,0">
            <TextBlock Text="CG Technology | https://cg-technology.github.io" FontSize="12" Foreground="#475569" VerticalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnOpenOutput" Content="Open Output Folder" Margin="0,0,10,0"/>
                <Button x:Name="BtnClearLog" Content="Clear Log"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Element References
$txtSourceFolder   = $window.FindName("TxtSourceFolder")
$btnBrowseSource   = $window.FindName("BtnBrowseSource")
$txtSetupFile      = $window.FindName("TxtSetupFile")
$btnBrowseSetup     = $window.FindName("BtnBrowseSetup")
$txtOutputFolder   = $window.FindName("TxtOutputFolder")
$btnBrowseOutput   = $window.FindName("BtnBrowseOutput")
$chkSkipPackaging  = $window.FindName("ChkSkipPackaging")
$btnStartPackaging = $window.FindName("BtnStartPackaging")
$txtConsoleLog     = $window.FindName("TxtConsoleLog")
$btnOpenOutput     = $window.FindName("BtnOpenOutput")
$btnClearLog       = $window.FindName("BtnClearLog")

# Default values
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$defaultOutput = Join-Path (Split-Path -Parent $scriptDir) "output"
$txtOutputFolder.Text = $defaultOutput

function Append-Log([string]$msg) {
    $time = (Get-Date).ToString("HH:mm:ss")
    $txtConsoleLog.AppendText("[$time] $msg`r`n")
    $txtConsoleLog.ScrollToEnd()
}

Append-Log "Intune App Packager Studio ready."
Append-Log "Select a source folder and installer executable to begin."

# Browse Source Folder
$btnBrowseSource.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select Source Folder containing Installer"
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSourceFolder.Text = $fbd.SelectedPath
        Append-Log "Selected source directory: $($fbd.SelectedPath)"
    }
})

# Browse Setup File
$btnBrowseSetup.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select Setup Executable or MSI"
    $ofd.Filter = "Installers (*.exe;*.msi)|*.exe;*.msi|All Files (*.*)|*.*"
    if ($txtSourceFolder.Text -and (Test-Path $txtSourceFolder.Text)) {
        $ofd.InitialDirectory = $txtSourceFolder.Text
    }
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $file = $ofd.FileName
        $dir = [System.IO.Path]::GetDirectoryName($file)
        $fileName = [System.IO.Path]::GetFileName($file)

        if (-not $txtSourceFolder.Text) {
            $txtSourceFolder.Text = $dir
        }
        $txtSetupFile.Text = $fileName
        Append-Log "Selected setup executable: $fileName"
    }
})

# Browse Output Folder
$btnBrowseOutput.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select Output Folder for .intunewin and Detection Script"
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputFolder.Text = $fbd.SelectedPath
        Append-Log "Selected output directory: $($fbd.SelectedPath)"
    }
})

# Open Output Folder
$btnOpenOutput.Add_Click({
    $out = $txtOutputFolder.Text
    if ($out -and (Test-Path $out)) {
        Start-Process "explorer.exe" -ArgumentList "`"$out`""
    } else {
        [System.Windows.MessageBox]::Show("Output directory does not exist yet.", "Notice", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
})

# Clear Log
$btnClearLog.Add_Click({
    $txtConsoleLog.Clear()
})

# Run Packaging
$btnStartPackaging.Add_Click({
    $src = $txtSourceFolder.Text.Trim()
    $setup = $txtSetupFile.Text.Trim()
    $out = $txtOutputFolder.Text.Trim()
    $skip = $chkSkipPackaging.IsChecked

    if (-not $src -or -not (Test-Path $src)) {
        [System.Windows.MessageBox]::Show("Please specify a valid Source Folder.", "Input Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    if (-not $setup) {
        [System.Windows.MessageBox]::Show("Please specify the Setup File name.", "Input Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $fullSetup = Join-Path $src $setup
    if (-not (Test-Path $fullSetup)) {
        [System.Windows.MessageBox]::Show("The setup file was not found in the source directory:`n$fullSetup", "File Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    $btnStartPackaging.IsEnabled = $false
    Append-Log "Starting packaging workflow..."
    Append-Log "Source: $src"
    Append-Log "Setup: $setup"
    Append-Log "Output: $out"

    try {
        $cliScript = Join-Path $scriptDir "IntunePackager.ps1"
        $cliArgs = @(
            "-SourceFolder", "`"$src`"",
            "-SetupFile", "`"$setup`"",
            "-OutputFolder", "`"$out`""
        )
        if ($skip) {
            $cliArgs += "-SkipPackaging"
        }

        # Run CLI in PowerShell process capturing output
        $procInfo = New-Object System.Diagnostics.ProcessStartInfo
        $procInfo.FileName = "powershell.exe"
        $procInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$cliScript`" $($cliArgs -join ' ')"
        $procInfo.RedirectStandardOutput = $true
        $procInfo.RedirectStandardError = $true
        $procInfo.UseShellExecute = $false
        $procInfo.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($procInfo)
        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()

        if ($stdout) {
            foreach ($line in ($stdout -split "`r?`n")) {
                if ($line) { Append-Log $line }
            }
        }
        if ($stderr) {
            Append-Log "[STDERR] $stderr"
        }

        if ($proc.ExitCode -eq 0) {
            Append-Log "Packaging and detection rule generation complete!"
            [System.Windows.MessageBox]::Show("Packaging completed successfully!`nArtifacts generated in output directory.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            Append-Log "Process finished with exit code: $($proc.ExitCode)"
        }
    }
    catch {
        Append-Log "[ERROR] $($_.Exception.Message)"
    }
    finally {
        $btnStartPackaging.IsEnabled = $true
    }
})

# Show GUI
$null = $window.ShowDialog()
