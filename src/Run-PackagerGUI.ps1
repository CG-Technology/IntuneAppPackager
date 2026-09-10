<#
.SYNOPSIS
    CG Technology Intune App Packager - Interactive Desktop GUI
.DESCRIPTION
    WPF Graphical interface for packaging Win32 applications and auto-generating
    detection scripts for Microsoft Intune.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

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

        <!-- Progress Bar & Active Status Strip -->
        <Border Grid.Row="2" x:Name="BorderProgress" Visibility="Collapsed" Background="#0F172A" BorderBrush="#1E293B" BorderThickness="1" CornerRadius="8" Padding="14,12" Margin="0,0,0,14">
            <StackPanel>
                <Grid Margin="0,0,0,8">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <Ellipse Width="8" Height="8" Fill="#818CF8" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock x:Name="TxtStatus" Text="Packaging application in progress..." FontSize="12" FontWeight="SemiBold" Foreground="#F8FAFC"/>
                    </StackPanel>
                    <TextBlock x:Name="TxtElapsed" Text="00:00" FontSize="12" FontWeight="Bold" Foreground="#818CF8" HorizontalAlignment="Right"/>
                </Grid>
                <ProgressBar x:Name="ProgressBar" Height="6" IsIndeterminate="True" 
                             Background="#111827" Foreground="#6366F1" BorderBrush="#334155" BorderThickness="1"/>
            </StackPanel>
        </Border>

        <!-- Log Output Console -->
        <Border Grid.Row="3" Background="#060911" BorderBrush="#1E293B" BorderThickness="1" CornerRadius="8" Padding="12">
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
        <Grid Grid.Row="4" Margin="0,16,0,0">
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
$borderProgress    = $window.FindName("BorderProgress")
$txtStatus         = $window.FindName("TxtStatus")
$txtElapsed        = $window.FindName("TxtElapsed")
$progressBar       = $window.FindName("ProgressBar")
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

function Select-FolderDialog([string]$description, [string]$initialPath) {
    try {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = $description
        $fbd.ShowNewFolderButton = $true
        if ($initialPath -and (Test-Path $initialPath)) {
            $fbd.SelectedPath = $initialPath
        }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $fbd.SelectedPath
        }
    }
    catch {
        # Fallback to Shell.Application COM object if WinForms has issues
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.BrowseForFolder(0, $description, 0, 0)
            if ($folder) {
                return $folder.Self.Path
            }
        } catch {}
    }
    return $null
}

# Browse Source Folder
$btnBrowseSource.Add_Click({
    $selected = Select-FolderDialog "Select Source Folder containing Installer" $txtSourceFolder.Text
    if ($selected) {
        $txtSourceFolder.Text = $selected
        Append-Log "Selected source directory: $selected"
    }
})

# Browse Setup File
$btnBrowseSetup.Add_Click({
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Title = "Select Setup Executable or MSI"
    $ofd.Filter = "Installers (*.exe;*.msi)|*.exe;*.msi|All Files (*.*)|*.*"
    if ($txtSourceFolder.Text -and (Test-Path $txtSourceFolder.Text)) {
        $ofd.InitialDirectory = $txtSourceFolder.Text
    }
    if ($ofd.ShowDialog($window) -eq $true) {
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
    $selected = Select-FolderDialog "Select Output Folder for .intunewin and Detection Script" $txtOutputFolder.Text
    if ($selected) {
        $txtOutputFolder.Text = $selected
        Append-Log "Selected output directory: $selected"
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
    $borderProgress.Visibility = [System.Windows.Visibility]::Visible
    $txtStatus.Text = "Initializing packaging workflow..."
    $txtElapsed.Text = "00:00"

    Append-Log "=================================================="
    Append-Log "Starting packaging workflow..."
    Append-Log "Source: $src"
    Append-Log "Setup:  $setup"
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

        $tempOut    = [System.IO.Path]::GetTempFileName()
        $tempErr    = [System.IO.Path]::GetTempFileName()
        $tempStatus = [System.IO.Path]::GetTempFileName()

        $cliArgs += @("-StatusFile", "`"$tempStatus`"")

        $proc = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$cliScript`" $($cliArgs -join ' ')" `
            -RedirectStandardOutput $tempOut `
            -RedirectStandardError $tempErr `
            -PassThru -NoNewWindow

        $outStream = [System.IO.File]::Open($tempOut, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $errStream = [System.IO.File]::Open($tempErr, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $outReader = New-Object System.IO.StreamReader($outStream, [System.Text.Encoding]::UTF8)
        $errReader = New-Object System.IO.StreamReader($errStream, [System.Text.Encoding]::UTF8)

        $script:packagingState = @{
            Process    = $proc
            OutStream  = $outStream
            ErrStream  = $errStream
            OutReader  = $outReader
            ErrReader  = $errReader
            StartTime  = [System.DateTime]::Now
            TempOut    = $tempOut
            TempErr    = $tempErr
            TempStatus = $tempStatus
            Timer      = $null
        }

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [System.TimeSpan]::FromMilliseconds(200)
        $script:packagingState.Timer = $timer

        $timer.Add_Tick({
            if (-not $script:packagingState) { return }

            # Update elapsed timer
            $elapsed = [System.DateTime]::Now - $script:packagingState.StartTime
            $txtElapsed.Text = "$($elapsed.Minutes.ToString('00')):$($elapsed.Seconds.ToString('00'))"

            # Stream newly written standard output
            $outReader = $script:packagingState.OutReader
            if ($outReader) {
                $newOut = $outReader.ReadToEnd()
                if ($newOut) {
                    foreach ($line in ($newOut -split "`r?`n")) {
                        if ($line) {
                            Append-Log $line
                            if ($line -match "\[INFO\]\s*(.+)") {
                                $txtStatus.Text = $Matches[1]
                            } elseif ($line -match "\[PASS\]\s*(.+)") {
                                $txtStatus.Text = $Matches[1]
                            } elseif ($line -match "\[WARN\]\s*(.+)") {
                                $txtStatus.Text = $Matches[1]
                            }
                        }
                    }
                }
            }

            # Stream newly written standard error
            $errReader = $script:packagingState.ErrReader
            if ($errReader) {
                $newErr = $errReader.ReadToEnd()
                if ($newErr) {
                    foreach ($line in ($newErr -split "`r?`n")) {
                        if ($line) {
                            Append-Log "[STDERR] $line"
                        }
                    }
                }
            }

            # Check process completion
            $proc = $script:packagingState.Process
            if ($proc -and $proc.HasExited) {
                $t = $script:packagingState.Timer
                if ($t) { $t.Stop() }

                # Final read flush
                if ($outReader) {
                    $finalOut = $outReader.ReadToEnd()
                    if ($finalOut) {
                        foreach ($line in ($finalOut -split "`r?`n")) {
                            if ($line) { Append-Log $line }
                        }
                    }
                    $outReader.Close()
                }
                if ($errReader) {
                    $finalErr = $errReader.ReadToEnd()
                    if ($finalErr) {
                        foreach ($line in ($finalErr -split "`r?`n")) {
                            if ($line) { Append-Log "[STDERR] $line" }
                        }
                    }
                    $errReader.Close()
                }

                if ($script:packagingState.OutStream) { $script:packagingState.OutStream.Close() }
                if ($script:packagingState.ErrStream) { $script:packagingState.ErrStream.Close() }

                $tempOut    = $script:packagingState.TempOut
                $tempErr    = $script:packagingState.TempErr
                $tempStatus = $script:packagingState.TempStatus

                $isSuccess = $false
                if ($tempStatus -and (Test-Path $tempStatus)) {
                    $statusVal = ([System.IO.File]::ReadAllText($tempStatus)).Trim()
                    if ($statusVal -eq "0") {
                        $isSuccess = $true
                    }
                }

                # Resilient fallback checks
                if (-not $isSuccess) {
                    if ($proc.ExitCode -eq 0) {
                        $isSuccess = $true
                    } elseif ($txtConsoleLog.Text -match "Packaging Workflow Completed Successfully!") {
                        $isSuccess = $true
                    }
                }

                Remove-Item $tempOut, $tempErr, $tempStatus -Force -ErrorAction SilentlyContinue

                $exitCode = $proc.ExitCode
                $totalSec = [Math]::Round(($elapsed).TotalSeconds, 1)

                $script:packagingState = $null

                $borderProgress.Visibility = [System.Windows.Visibility]::Collapsed
                $btnStartPackaging.IsEnabled = $true

                if ($isSuccess) {
                    Append-Log "Packaging workflow completed in $totalSec seconds!"
                    [System.Windows.MessageBox]::Show("Packaging completed successfully in $totalSec seconds!`nArtifacts generated in output directory.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    $exitDisplay = if ($exitCode -ne $null) { $exitCode } else { "Non-zero / Error" }
                    Append-Log "Process finished with status: $exitDisplay"
                    [System.Windows.MessageBox]::Show("Packaging process encountered an issue (status: $exitDisplay).`nReview log for details.", "Packaging Issue", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                }
            }
        })

        $timer.Start()
    }
    catch {
        Append-Log "[ERROR] $($_.Exception.Message)"
        $borderProgress.Visibility = [System.Windows.Visibility]::Collapsed
        $btnStartPackaging.IsEnabled = $true
    }
})

# Clean up on window close
$window.Add_Closing({
    if ($script:packagingState) {
        try {
            if ($script:packagingState.Timer) { $script:packagingState.Timer.Stop() }
            if ($script:packagingState.Process -and -not $script:packagingState.Process.HasExited) {
                $script:packagingState.Process.Kill()
            }
            if ($script:packagingState.OutReader) { $script:packagingState.OutReader.Close() }
            if ($script:packagingState.ErrReader) { $script:packagingState.ErrReader.Close() }
            if ($script:packagingState.OutStream) { $script:packagingState.OutStream.Close() }
            if ($script:packagingState.ErrStream) { $script:packagingState.ErrStream.Close() }
            Remove-Item $script:packagingState.TempOut, $script:packagingState.TempErr, $script:packagingState.TempStatus -Force -ErrorAction SilentlyContinue
        } catch {}
    }
})

# Show GUI
$null = $window.ShowDialog()
