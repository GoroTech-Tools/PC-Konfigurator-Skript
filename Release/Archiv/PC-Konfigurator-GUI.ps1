param(
    [switch]$DryRun,
    [switch]$TestUI
)

# ============================================================================
# PC-Konfigurator-GUI
# GUI-Nachfolger des konsolenbasierten PC-Konfigurator (PowerShell/Windows Presentation Foundation).
# Die Windows Presentation Foundation stellt die interaktive Oberfläche für
# den geführten Assistenten, Eingaben, Bestätigungen und Live-Protokolle bereit.
# Die komplette fachliche Logik wurde 1:1 aus PC-Konfigurator.ps1 (Original-
# Projekt) übernommen. Alle interaktiven Read-Host-Konsolenabfragen wurden
# durch einen Windows-Presentation-Foundation-Einrichtungsassistenten (Wizard) ersetzt.
# ============================================================================

# 32-Bit-Office liefert nur eine Win32-Typbibliothek; ein 64-Bit-Host bricht sonst mit
# TYPE_E_CANTLOADLIBRARY (0x80029C4A) beim ersten Office-COM-Zugriff ab.
# A ps2exe executable cannot be relaunched with PowerShell's -File parameter.
# Only perform the original 32-bit Office relaunch when running the .ps1 source.
$isScriptHost = -not [string]::IsNullOrWhiteSpace($PSCommandPath) -and
    [IO.Path]::GetExtension($PSCommandPath) -ieq '.ps1'

# Eine Release-EXE kann direkt aus einem entpackten Ordner gestartet werden.
# Beim ersten Start werden alle Laufzeitdateien nach LocalAppData kopiert und
# die Anwendung von dort erneut gestartet. Im Zielordner läuft sie danach ohne
# Bezug zum Release- oder OneDrive-Ordner.
function Show-PreparationWindow {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop

    $window = New-Object System.Windows.Window
    $window.Title = 'PC-Konfigurator-GUI wird vorbereitet'
    $window.Width = 560
    $window.Height = 390
    $window.WindowStartupLocation = 'CenterScreen'
    $window.ResizeMode = 'NoResize'
    $window.WindowStyle = 'SingleBorderWindow'
    $window.ShowInTaskbar = $true

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = '24'

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = 'PC-Konfigurator-GUI wird vorbereitet'
    $title.FontSize = 20
    $title.FontWeight = 'SemiBold'
    $title.Margin = '0,0,0,16'
    [void]$panel.Children.Add($title)

    $message = New-Object System.Windows.Controls.TextBlock
    $message.Text = 'Zur Vorbereitung wird der PC-Konfigurator in Ihrem System hinterlegt. Es geht gleich weiter.'
    $message.TextWrapping = 'Wrap'
    $message.FontSize = 14
    $message.Margin = '0,0,0,18'
    [void]$panel.Children.Add($message)

    $status = New-Object System.Windows.Controls.TextBlock
    $status.Text = 'Vorbereitung wird gestartet ...'
    $status.FontWeight = 'SemiBold'
    $status.Foreground = [System.Windows.Media.Brushes]::DarkSlateGray
    $status.Margin = '0,0,0,8'
    [void]$panel.Children.Add($status)

    $activityLog = New-Object System.Windows.Controls.ListBox
    $activityLog.Height = 150
    $activityLog.IsHitTestVisible = $false
    $activityLog.FontFamily = 'Consolas'
    $activityLog.FontSize = 11
    [void]$panel.Children.Add($activityLog)

    $window.Content = $panel
    $window.Show()
    $window.UpdateLayout()
    $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)

    return [pscustomobject]@{
        Window = $window
        Status = $status
        ActivityLog = $activityLog
    }
}

function Set-PreparationWindowStatus {
    param(
        [Parameter(Mandatory = $true)]$PreparationWindow,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $PreparationWindow.Status.Text = $Message
    $PreparationWindow.Window.UpdateLayout()
    $PreparationWindow.Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Add-PreparationWindowEntry {
    param(
        [Parameter(Mandatory = $true)]$PreparationWindow,
        [Parameter(Mandatory = $true)][string]$Message
    )

    [void]$PreparationWindow.ActivityLog.Items.Add($Message)
    $PreparationWindow.ActivityLog.ScrollIntoView($Message)
    $PreparationWindow.Window.UpdateLayout()
    $PreparationWindow.Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Get-AppendedPayload {
    param(
        [Parameter(Mandatory = $true)][string]$ExecutablePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    $payloadMarker = [Text.Encoding]::ASCII.GetBytes('PCKGUI-PAYLOAD-1')
    $footerLength = $payloadMarker.Length + 8
    $source = [IO.File]::Open($ExecutablePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        if ($source.Length -lt $footerLength) {
            throw 'Das eingebettete Laufzeitpaket wurde nicht gefunden.'
        }

        $source.Position = $source.Length - $footerLength
        $payloadLengthBytes = New-Object byte[] 8
        [void]$source.Read($payloadLengthBytes, 0, $payloadLengthBytes.Length)
        $payloadLength = [BitConverter]::ToInt64($payloadLengthBytes, 0)
        $markerBytes = New-Object byte[] $payloadMarker.Length
        [void]$source.Read($markerBytes, 0, $markerBytes.Length)
        $markerText = [Text.Encoding]::ASCII.GetString($markerBytes)
        if ($markerText -ne 'PCKGUI-PAYLOAD-1') {
            throw 'Das eingebettete Laufzeitpaket ist ungültig.'
        }

        $payloadOffset = $source.Length - $footerLength - $payloadLength
        if ($payloadLength -le 0 -or $payloadOffset -lt 0) {
            throw 'Die Größe des eingebetteten Laufzeitpakets ist ungültig.'
        }

        $source.Position = $payloadOffset
        $destination = [IO.File]::Open($DestinationPath, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $buffer = New-Object byte[] 1048576
            $remaining = $payloadLength
            while ($remaining -gt 0) {
                $bytesRead = $source.Read($buffer, 0, [Math]::Min($buffer.Length, $remaining))
                if ($bytesRead -le 0) { throw 'Das Laufzeitpaket konnte nicht vollständig gelesen werden.' }
                $destination.Write($buffer, 0, $bytesRead)
                $remaining -= $bytesRead
            }
        } finally {
            $destination.Dispose()
        }
    } finally {
        $source.Dispose()
    }
}

function Expand-PayloadWithProgress {
    param(
        [Parameter(Mandatory = $true)][string]$PayloadPath,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [Parameter(Mandatory = $true)]$PreparationWindow
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $destinationRoot = [IO.Path]::GetFullPath($DestinationPath).TrimEnd('\') + '\'
    $archive = [IO.Compression.ZipFile]::OpenRead($PayloadPath)
    try {
        $fileEntries = @($archive.Entries | Where-Object { -not [string]::IsNullOrEmpty($_.Name) })
        $entryIndex = 0
        foreach ($entry in $fileEntries) {
            $entryIndex++
            $targetPath = [IO.Path]::GetFullPath((Join-Path $DestinationPath $entry.FullName))
            if (-not $targetPath.StartsWith($destinationRoot, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Ungültiger Pfad im Laufzeitpaket: $($entry.FullName)"
            }

            Set-PreparationWindowStatus -PreparationWindow $PreparationWindow -Message "Laufzeitdateien werden eingerichtet ($entryIndex/$($fileEntries.Count)) ..."
            Add-PreparationWindowEntry -PreparationWindow $PreparationWindow -Message "Entpacke: $($entry.FullName)"
            New-Item -ItemType Directory -Path (Split-Path -Path $targetPath -Parent) -Force | Out-Null
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
        }
    } finally {
        $archive.Dispose()
    }
}

if (-not $isScriptHost) {
    $currentExecutable = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $currentRoot = Split-Path -Path $currentExecutable -Parent
    $localAppDataRoot = Join-Path $env:LOCALAPPDATA 'PC-Konfigurator-GUI'
    $installedExecutable = Join-Path $localAppDataRoot 'PC-Konfigurator-GUI.exe'
    $currentRootFull = [IO.Path]::GetFullPath($currentRoot).TrimEnd('\')
    $localAppDataRootFull = [IO.Path]::GetFullPath($localAppDataRoot).TrimEnd('\')
    $embeddedPayload = Join-Path $localAppDataRoot '_embedded-payload.zip'
    if ($currentRootFull -ine $localAppDataRootFull) {
        $preparationWindow = $null
        try {
            New-Item -ItemType Directory -Path $localAppDataRoot -Force | Out-Null
            $preparationWindow = Show-PreparationWindow
            Add-PreparationWindowEntry -PreparationWindow $preparationWindow -Message 'Lese eingebettetes Laufzeitpaket ...'
            Get-AppendedPayload -ExecutablePath $currentExecutable -DestinationPath $embeddedPayload
            Expand-PayloadWithProgress -PayloadPath $embeddedPayload -DestinationPath $localAppDataRoot -PreparationWindow $preparationWindow
            Set-PreparationWindowStatus -PreparationWindow $preparationWindow -Message 'Vorbereitung abgeschlossen. Der Assistent wird gestartet ...'
            Remove-Item -LiteralPath $embeddedPayload -Force
        } catch {
            if ($preparationWindow -and $preparationWindow.Window) {
                $preparationWindow.Window.Close()
            }
            Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
            [System.Windows.MessageBox]::Show(
                "Die eingebetteten Laufzeitdateien konnten nicht entpackt werden:`r`n$($_.Exception.Message)",
                'PC-Konfigurator-GUI', 'OK', 'Error') | Out-Null
            exit 1
        } finally {
            if ($preparationWindow -and $preparationWindow.Window) {
                $preparationWindow.Window.Close()
            }
        }
    }

    $sourceHasAssets = Test-Path (Join-Path $currentRoot 'Datei-Vorlagen')
    $localAssetsReady = Test-Path (Join-Path $localAppDataRoot 'Datei-Vorlagen')
    if ($currentRootFull -ine $localAppDataRootFull -and ($sourceHasAssets -or $localAssetsReady)) {
        try {
            New-Item -ItemType Directory -Path $localAppDataRoot -Force | Out-Null
            Copy-Item -LiteralPath $currentExecutable -Destination $installedExecutable -Force
            foreach ($runtimeDirectory in @('Datei-Vorlagen', 'Fonts', 'docs', 'src')) {
                $sourceDirectory = Join-Path $currentRoot $runtimeDirectory
                if ($sourceHasAssets -and (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
                    Copy-Item -LiteralPath $sourceDirectory -Destination $localAppDataRoot -Recurse -Force
                }
            }
            $runtimeSourceRoot = Join-Path $currentRoot 'src'
            foreach ($runtimeFile in @('Pin-Desktop-Schnellzugriff.ps1', 'RELEASE-NOTES.txt', 'RELEASE-VERSION.txt')) {
                $sourceFile = Join-Path $runtimeSourceRoot $runtimeFile
                if (Test-Path -LiteralPath $sourceFile -PathType Leaf) {
                    Copy-Item -LiteralPath $sourceFile -Destination $localAppDataRoot -Force
                }
            }
            $readmeSource = Join-Path $currentRoot 'README.MD'
            if (Test-Path -LiteralPath $readmeSource -PathType Leaf) {
                Copy-Item -LiteralPath $readmeSource -Destination $localAppDataRoot -Force
            }
        } catch {
            Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
            [System.Windows.MessageBox]::Show(
                "Die Anwendung konnte nicht nach LocalAppData bereitgestellt werden:`r`n$($_.Exception.Message)",
                'PC-Konfigurator-GUI', 'OK', 'Error') | Out-Null
            exit 1
        }

        $relaunchArgs = @()
        if ($DryRun) { $relaunchArgs += '-DryRun' }
        if ($TestUI) { $relaunchArgs += '-TestUI' }
        if ($relaunchArgs.Count -gt 0) {
            $installedProcess = Start-Process -FilePath $installedExecutable -ArgumentList $relaunchArgs -Wait -PassThru
        } else {
            $installedProcess = Start-Process -FilePath $installedExecutable -Wait -PassThru
        }
        exit $installedProcess.ExitCode
    }
}
if ([Environment]::Is64BitProcess -and -not $env:PCK_ARCH_RELAUNCH) {
    $wordTypeLibKey = 'Registry::HKEY_CLASSES_ROOT\TypeLib\{00020905-0000-0000-C000-000000000046}'
    $needs32Bit = $false
    if (Test-Path $wordTypeLibKey) {
        foreach ($versionKey in Get-ChildItem $wordTypeLibKey -ErrorAction SilentlyContinue) {
            $localeKey = Join-Path $versionKey.PSPath '0'
            if (-not (Test-Path $localeKey)) { continue }
            $platforms = (Get-ChildItem $localeKey -ErrorAction SilentlyContinue).PSChildName
            if (($platforms -contains 'Win32') -and ($platforms -notcontains 'Win64')) {
                $needs32Bit = $true
            }
        }
    }

    if ($needs32Bit) {
        $powerShell32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        if (Test-Path $powerShell32) {
            $env:PCK_ARCH_RELAUNCH = '1'
            if ($isScriptHost) {
                $relaunchArgs = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
                if ($DryRun) { $relaunchArgs += '-DryRun' }
                if ($TestUI) { $relaunchArgs += '-TestUI' }
                $process = Start-Process -FilePath $powerShell32 -ArgumentList $relaunchArgs -NoNewWindow -Wait -PassThru
            } else {
                $relaunchTarget = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
                # Auch beim kompilierten ps2exe-Host müssen die Schalter an den
                # 32-Bit-Child-Prozess weitergereicht werden. Sonst bleibt der
                # Parent nach -TestUI bzw. DryRun als unsichtbarer Restprozess
                # bestehen und der Child startet mit anderem Verhalten.
                $relaunchArgs = @()
                if ($DryRun) { $relaunchArgs += '-DryRun' }
                if ($TestUI) { $relaunchArgs += '-TestUI' }
                if ($relaunchArgs.Count -gt 0) {
                    $process = Start-Process -FilePath $relaunchTarget -ArgumentList $relaunchArgs -Wait -PassThru
                } else {
                    $process = Start-Process -FilePath $relaunchTarget -Wait -PassThru
                }
            }
            exit $process.ExitCode
        }
        # Kein 32-Bit-PowerShell gefunden - Skript läuft im 64-Bit-Prozess weiter;
        # Office-COM-Zugriffe können in diesem Fall fehlschlagen (wird protokolliert).
    }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xml

# ----------------------------------------------------------------------------
# Globale Pfade
# HINWEIS: Bewusst "PC-Konfigurator-GUI" (statt "PC-Konfigurator"), damit Logs
# dieses Projekts nicht mit denen des Konsolen-Originalprojekts kollidieren.
# ----------------------------------------------------------------------------
$script:documentsPath = [Environment]::GetFolderPath('MyDocuments')
$script:logDir = Join-Path $script:documentsPath "PC-Konfigurator-GUI\Logs"
$script:logFile = Join-Path $script:logDir "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss_fff')_$PID.log"
$script:robocopyLogDir = Join-Path $script:documentsPath "PC-Konfigurator-GUI\Robocopy-Logs"
$script:BackupTargetPath = $null

# Robuste Ermittlung des Skript-/Anwendungsverzeichnisses: $PSScriptRoot ist bei einer
# mit ps2exe kompilierten EXE i. d. R. das Verzeichnis der EXE. Als Fallback wird der
# Pfad des aktuellen Prozesses (der kompilierten EXE) verwendet.
$script:ScriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
} else {
    try {
        Split-Path -Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent
    } catch {
        (Get-Location).Path
    }
}

# Thread-sichere Warteschlange: Hintergrund-Runspaces schreiben hier ihre Zeilen hinein,
# ein DispatcherTimer im GUI-Thread liest sie regelmäßig aus und schreibt sie in die Log-TextBox.
$script:UiQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# ----------------------------------------------------------------------------
# Einheitliche Konsolenfarben (aus dem Original übernommen) + Spiegelung in die
# UI-Log-Queue, damit Ausgaben auch im GUI sichtbar werden.
# Standard/Bestätigung = Weiß, Eingabeaufforderung = Grün,
# Eingaben/Informationen = Cyan, Fehler = Rot.
# ----------------------------------------------------------------------------
function Write-Host {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        [object[]]$Object,
        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor,
        [string]$Separator,
        [switch]$NoNewline
    )

    $mappedColor = if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
        switch ($ForegroundColor.ToString()) {
            'Red' { [ConsoleColor]::Red }
            'Green' { [ConsoleColor]::White }
            'Cyan' { [ConsoleColor]::Cyan }
            'Yellow' { [ConsoleColor]::Green }
            default { [ConsoleColor]::White }
        }
    } else {
        [ConsoleColor]::White
    }

    $nativeParameters = @{
        Object = $Object
        ForegroundColor = $mappedColor
    }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $nativeParameters.BackgroundColor = $BackgroundColor }
    if ($PSBoundParameters.ContainsKey('Separator')) { $nativeParameters.Separator = $Separator }
    if ($NoNewline) { $nativeParameters.NoNewline = $true }

    try {
        Microsoft.PowerShell.Utility\Write-Host @nativeParameters
    } catch {
        # In ps2exe -noConsole-EXEs existiert keine Konsole - Fehler hier ignorieren.
        $null = $_.Exception.Message
    }

    try {
        $text = ($Object -join ' ')
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $script:UiQueue.Enqueue($text)
        }
    } catch {
        $null = $_.Exception.Message
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    try {
        if ([string]::IsNullOrWhiteSpace($script:logDir)) {
            $script:logDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "PC-Konfigurator-GUI\Logs"
        }
        if ([string]::IsNullOrWhiteSpace($script:logFile)) {
            $script:logFile = Join-Path $script:logDir "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss_fff')_$PID.log"
        }

        if (-not (Test-Path -LiteralPath $script:logDir)) {
            New-Item -ItemType Directory -Path $script:logDir -Force -ErrorAction Stop | Out-Null
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $entry = "$timestamp [$Level] - $Message"
        Add-Content -LiteralPath $script:logFile -Value $entry -Encoding UTF8 -ErrorAction Stop
    } catch {
        $fallbackDir = Join-Path $env:TEMP 'PC-Konfigurator-GUI\Logs'
        $fallbackFile = Join-Path $fallbackDir 'Fallback.log'
        try {
            New-Item -ItemType Directory -Path $fallbackDir -Force -ErrorAction Stop | Out-Null
            $fallbackEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] - $Message (Primärlog fehlgeschlagen: $($_.Exception.Message))"
            Add-Content -LiteralPath $fallbackFile -Value $fallbackEntry -Encoding UTF8 -ErrorAction Stop
        } catch { $null = $_.Exception.Message }
        Write-Host -ForegroundColor Cyan "LOGFALLBACK [$Level] $Message"
    }
    try {
        $script:UiQueue.Enqueue("[$Level] $Message")
    } catch {
        $null = $_.Exception.Message
    }
}

function Clear-OldLogs {
    param (
        [string]$logDirPath = $script:logDir
    )
    if (Test-Path $logDirPath) {
        $logFiles = Get-ChildItem -Path $logDirPath -Filter "Log_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 3) {
            $filesToDelete = $logFiles | Select-Object -Skip 3
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force
            }
        }
    }
}

function Write-DryRunAction {
    param(
        [Parameter(Mandatory = $true)][string]$Message
    )
    Write-Host -ForegroundColor Cyan "[DRYRUN] $Message"
    Write-Log "DRYRUN: $Message" "INFO"
}

# --- Systemanforderungen prüfen (schnell, wird synchron im GUI-Thread ausgeführt) ---
function Test-SystemRequirements {
    Write-Log "Überprüfung der Windows- und Office-Version gestartet" "INFO"

    # === WINDOWS-VERSION PRÜFEN ===
    try {
        $windowsVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
        $windowsBuild = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
        $windowsCaption = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption

        Write-Log "Ermittelte Windows-Version: $windowsVersion (Build: $windowsBuild)" "INFO"
        Write-Log "Windows-Edition: $windowsCaption" "INFO"

        $windowsOK = $false
        if ($windowsVersion -match "^10\.") {
            if ([int]$windowsBuild -ge 18362) {
                $windowsOK = $true
                Write-Log "Windows 10 mit ausreichendem Build erkannt (Build $windowsBuild >= 18362)" "INFO"
            } else {
                Write-Log "Windows 10 Build zu alt (Build $windowsBuild < 18362)" "WARN"
            }
        } elseif ($windowsVersion -match "^11\.") {
            $windowsOK = $true
            Write-Log "Windows 11 erkannt" "INFO"
        } else {
            Write-Log "Nicht unterstützte Windows-Version: $windowsVersion" "WARN"
        }
    } catch {
        Write-Log "Fehler beim Ermitteln der Windows-Version: $($_.Exception.Message)" "ERROR"
        $windowsOK = $false
    }

    # === OFFICE-VERSION PRÜFEN ===
    $officeVersion = $null
    $officeMajorVersion = 0
    $officeOK = $false

    $clickToRunKeys = @(
        "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )

    foreach ($officeKey in $clickToRunKeys) {
        if (Test-Path $officeKey) {
            try {
                $regValue = Get-ItemProperty -Path $officeKey -ErrorAction SilentlyContinue
                if ($regValue.ProductVersion) {
                    $officeVersion = $regValue.ProductVersion
                    $officeMajorVersion = [int]($officeVersion.Split(".")[0])
                    Write-Log "ClickToRun Office gefunden: $officeVersion (Pfad: $officeKey)" "INFO"
                    break
                }
            } catch {
                Write-Log "Fehler beim Lesen von $officeKey : $($_.Exception.Message)" "WARN"
            }
        }
    }

    if (-not $officeVersion) {
        $msiKeys = @(
            "HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot",
            "HKLM:\Software\Microsoft\Office\15.0\Common\InstallRoot",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot",
            "HKLM:\Software\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot"
        )

        foreach ($msiKey in $msiKeys) {
            if (Test-Path $msiKey) {
                try {
                    $msiPath = (Get-ItemProperty -Path $msiKey -ErrorAction SilentlyContinue).Path
                    if ($msiPath -and (Test-Path $msiPath)) {
                        if ($msiKey -match "16\.0") {
                            $officeVersion = "16.0 (MSI-Installation)"
                            $officeMajorVersion = 16
                        } elseif ($msiKey -match "15\.0") {
                            $officeVersion = "15.0 (MSI-Installation)"
                            $officeMajorVersion = 15
                        }
                        Write-Log "MSI Office gefunden: $officeVersion (Pfad: $msiKey)" "INFO"
                        break
                    }
                } catch {
                    Write-Log "Fehler beim Prüfen von $msiKey : $($_.Exception.Message)" "WARN"
                }
            }
        }
    }

    if (-not $officeVersion) {
        $commonPaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office16\WINWORD.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16\WINWORD.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office15\WINWORD.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office15\WINWORD.EXE"
        )

        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                try {
                    $fileVersion = (Get-ItemProperty $path).VersionInfo.ProductVersion
                    if ($path -match "Office16") {
                        $officeVersion = "16.0 (Executable gefunden - Version: $fileVersion)"
                        $officeMajorVersion = 16
                    } elseif ($path -match "Office15") {
                        $officeVersion = "15.0 (Executable gefunden - Version: $fileVersion)"
                        $officeMajorVersion = 15
                    }
                    Write-Log "Office durch Executable gefunden: $officeVersion (Pfad: $path)" "INFO"
                    break
                } catch {
                    Write-Log "Fehler beim Prüfen der Datei $path : $($_.Exception.Message)" "WARN"
                }
            }
        }
    }

    if ($officeVersion) {
        $officeOK = $officeMajorVersion -ge 15
        Write-Log "Office-Validierung: Version $officeMajorVersion >= 15: $officeOK" "INFO"
    } else {
        Write-Log "Keine Office-Installation gefunden" "WARN"
        Write-Log "HINWEIS: Skript wird trotz fehlender Office-Erkennung fortgesetzt (möglicherweise Office 365 Web)" "WARN"
        $officeOK = $true
    }

    Write-Log "=== SYSTEMPRÜFUNG ZUSAMMENFASSUNG ===" "INFO"
    Write-Log "Windows OK: $windowsOK (Version: $windowsVersion, Build: $windowsBuild)" "INFO"
    Write-Log "Office OK: $officeOK (Version: $officeVersion)" "INFO"

    if ($windowsOK) {
        Write-Log "Systemanforderungen erfüllt - Anwendung wird fortgesetzt" "INFO"
        return $true
    } else {
        Write-Log "Systemanforderungen nicht erfüllt - Windows-Version zu alt" "ERROR"
        Write-Log "Erforderlich: Windows 10 (Build 18362+) oder Windows 11" "ERROR"
        return $false
    }
}

# ============================================================================
# Invoke-PCKonfiguratorPipeline
# ----------------------------------------------------------------------------
# Enthält die GESAMTE automatisierte Konfigurationslogik aus dem Original-
# Konsolenskript PC-Konfigurator.ps1 (1:1 übernommen), verpackt als eine große
# Funktion mit verschachtelten Hilfsfunktionen. Alle vormals interaktiven
# Read-Host-Abfragen wurden bereits VORHER im GUI-Assistenten beantwortet und
# werden hier über den $Params-Parameter hereingereicht.
#
# Diese Funktion wird NICHT im GUI-Thread ausgeführt, sondern über eine
# separate Runspace im Hintergrund gestartet (siehe Start-BackgroundPipeline),
# damit lange Robocopy-/Office-COM-Vorgänge die Benutzeroberfläche nicht
# blockieren. Da Runspaces innerhalb desselben Prozesses laufen, kann die
# thread-sichere $UiQueue (ConcurrentQueue) problemlos als Referenz mit
# übergeben werden; ein DispatcherTimer im GUI-Thread liest sie aus.
# ============================================================================
function Invoke-PCKonfiguratorPipeline {
    param(
        [Parameter(Mandatory = $true)]$Params
    )

    # --- Aus den Wizard-Eingaben übernommene / abgeleitete Werte ---
    $logDir = $Params.LogDir
    $logFile = $Params.LogFile
    $robocopyLogDir = $Params.RobocopyLogDir
    $ScriptRoot = $Params.ScriptRoot
    $AssetRoot = if (Test-Path -LiteralPath (Join-Path $ScriptRoot 'Datei-Vorlagen') -PathType Container) {
        $ScriptRoot
    } else {
        Split-Path $ScriptRoot -Parent
    }
    $UiQueue = $Params.UiQueue
    $DryRun = [bool]$Params.DryRun
    $BackupTargetPath = $null

    # ------------------------------------------------------------------
    # Write-Host-Override und Write-Log (wie im Hauptskript), damit auch
    # innerhalb der Hintergrund-Runspace sämtliche Ausgaben in die Log-Datei
    # und in die $UiQueue (und damit in die GUI-TextBox) gespiegelt werden.
    # ------------------------------------------------------------------
    function Write-Host {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
            [object[]]$Object,
            [ConsoleColor]$ForegroundColor,
            [ConsoleColor]$BackgroundColor,
            [string]$Separator,
            [switch]$NoNewline
        )
        try {
            $text = ($Object -join ' ')
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $UiQueue.Enqueue($text)
            }
        } catch {
            $null = $_.Exception.Message
        }
    }

    function Write-Log {
        param(
            [Parameter(Mandatory=$true)][string]$Message,
            [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
        )
        try {
            if (-not (Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $entry = "$timestamp [$Level] - $Message"
            Add-Content -Path $logFile -Value $entry -Encoding UTF8
        } catch {
            Write-Host -ForegroundColor Cyan "LOGFALLBACK [$Level] $Message"
        }
        try {
            $UiQueue.Enqueue("[$Level] $Message")
        } catch {
            $null = $_.Exception.Message
        }
    }

    function Clear-OldLogs {
        param (
            [string]$logDirPath = $logDir
        )
        if (Test-Path $logDirPath) {
            $logFiles = Get-ChildItem -Path $logDirPath -Filter "Log_*.log" | Sort-Object LastWriteTime -Descending
            if ($logFiles.Count -gt 3) {
                $filesToDelete = $logFiles | Select-Object -Skip 3
                foreach ($file in $filesToDelete) {
                    Remove-Item -Path $file.FullName -Force
                }
            }
        }
    }

    function Write-DryRunAction {
        param(
            [Parameter(Mandatory = $true)][string]$Message
        )
        Write-Host -ForegroundColor Cyan "[DRYRUN] $Message"
        Write-Log "DRYRUN: $Message" "INFO"
    }

    function Assert-RegistryValue {
        param(
            [Parameter(Mandatory = $true)][string]$Path,
            [Parameter(Mandatory = $true)][string]$Name,
            [Parameter(Mandatory = $true)][object]$Expected,
            [string]$Description = $Name
        )

        if (-not (Test-Path -Path $Path)) {
            Write-Log "Validierung fehlgeschlagen: Pfad fehlt ($Path) beim Wert $Description" "WARN"
            return $false
        }

        try {
            $actual = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
            if ($actual -eq $Expected) {
                Write-Log "Validierung OK: $Description = $actual in $Path" "INFO"
                return $true
            }

            Write-Log "Validierung fehlgeschlagen: $Description in $Path -> erwartet '$Expected', ist '$actual'" "WARN"
            return $false
        } catch {
            Write-Log "Validierung fehlgeschlagen: $Description in $Path konnte nicht gelesen werden ($($_.Exception.Message))" "WARN"
            return $false
        }
    }

    function Set-ExplorerRecentAndFrequentState {
        param(
            [ValidateSet(0,1)][int]$ShowRecent = 0,
            [ValidateSet(0,1)][int]$ShowFrequent = 0
        )

        try {
            $regPath = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }

            Set-ItemProperty -Path $regPath -Name 'ShowRecent' -Value $ShowRecent -Type DWord -Force
            Set-ItemProperty -Path $regPath -Name 'ShowFrequent' -Value $ShowFrequent -Type DWord -Force
            Write-Log -Message "Explorer-Optionen gesetzt: ShowRecent=$ShowRecent, ShowFrequent=$ShowFrequent." -Level 'INFO'
            return $true
        } catch {
            Write-Log -Message "Fehler beim Setzen der Explorer-Optionen: $($_.Exception.Message)" -Level 'ERROR'
            return $false
        }
    }

    function Disable-ExplorerRecentAndFrequent {
        <#
        Deaktiviert im Windows-Explorer die Anzeige von:
        - Zuletzt verwendete Dateien
        - Häufig verwendete Ordner
        durch Setzen der entsprechenden Registry-Werte. Ein Explorer-Neustart wird bewusst nicht sofort erzwungen, damit während der laufenden Konfiguration keine Explorer-Sitzung unnötig unterbrochen wird.
        #>
        try {
            if (Set-ExplorerRecentAndFrequentState -ShowRecent 0 -ShowFrequent 0) {
                Write-Log -Message "Explorer-Neustart für 'ShowRecent'/'ShowFrequent' bewusst zurückgestellt; spätere Explorer-Aktualisierung übernimmt die Anwendung." -Level 'INFO'
            }
        } catch {
            Write-Log -Message "Fehler beim Setzen der Explorer-Optionen: $($_.Exception.Message)" -Level 'ERROR'
        }
    }

    function Set-DesktopInQuickAccess {
        <#
        Prüft, ob der Desktop des aktuell angemeldeten Benutzers im Schnellzugriff angeheftet ist.
        Falls nicht, wird der Desktop an den Schnellzugriff angeheftet.
        #>
        try {
            $helperCandidates = @(
                (Join-Path $ScriptRoot 'Pin-Desktop-Schnellzugriff.ps1'),
                (Join-Path (Split-Path $ScriptRoot -Parent) 'Pin-Desktop-Schnellzugriff.ps1'),
                (Join-Path (Split-Path (Split-Path $ScriptRoot -Parent) -Parent) 'Pin-Desktop-Schnellzugriff.ps1')
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

            $helperPath = $helperCandidates |
                Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
                Select-Object -First 1

            if (-not $helperPath) {
                Write-Log -Message 'Pin-Desktop-Schnellzugriff.ps1 wurde nicht gefunden.' -Level 'WARN'
                Write-Host -ForegroundColor Cyan "Hilfsskript für den Desktop-Schnellzugriff wurde nicht gefunden - Prüfung übersprungen."
                return
            }

            Write-Log -Message "Desktop-Schnellzugriff wird über Hilfsskript ausgeführt: $helperPath" -Level 'INFO'
            $helperResult = & $helperPath

            if ($helperResult -and $helperResult.Success) {
                $statusText = if ($helperResult.AlreadyPinned) { 'bereits vorhanden' } else { 'angeheftet' }
                Write-Log -Message "Desktop-Schnellzugriff erfolgreich verarbeitet ($statusText): $($helperResult.DesktopPath)" -Level 'INFO'
                return
            }

            Write-Log -Message 'Hilfsskript für den Desktop-Schnellzugriff lieferte keinen Erfolgstatus zurück.' -Level 'WARN'
        }
        catch {
            Write-Host -ForegroundColor Red "Fehler beim Prüfen des Schnellzugriffs für den Desktop: $_"
            Write-Log -Message "Fehler bei der Schnellzugriff-Prüfung für Desktop: $($_.Exception.Message)" -Level 'ERROR'
        }
    }

    Write-Log "=== Invoke-PCKonfiguratorPipeline gestartet ===" "INFO"
    Disable-ExplorerRecentAndFrequent

    # ------------------------------------------------------------------
    # sync() - Robocopy-Synchronisation von Datei-Vorlagen
    # ------------------------------------------------------------------
    function Clear-OldRobocopyLogs {
        param (
            [string]$robocopyLogDirPath = $robocopyLogDir
        )
        if (Test-Path $robocopyLogDirPath) {
            $logFiles = Get-ChildItem -Path $robocopyLogDirPath -Filter "*.log" | Sort-Object LastWriteTime -Descending
            if ($logFiles.Count -gt 3) {
                $filesToDelete = $logFiles | Select-Object -Skip 3
                foreach ($file in $filesToDelete) {
                    Remove-Item -Path $file.FullName -Force
                }
            }
        }
    }

    function sync() {
        # Zielverzeichnis
        $roboCopyBackupPath = $BackupTargetPath

        # Wie viele Instanzen von Robocopy sollen verwendet werden?
        $maxThreads = 5

        $excludeFiles = @(
            "Thumbs.db"
            "Muell.txt"
        )

        $excludeDirectories = @(
            '$Recycle.Bin'
            "System Volume Information"
        )

        $sourceDirectories = @(
            (Join-Path $AssetRoot 'Datei-Vorlagen')
        )

        function Get-DriveLetter {
            param (
                [Parameter(Mandatory = $true)]
                [string]$Path
            )
            if ($Path -match '^[A-Za-z]:\\') {
                $driveLetter = ($Path -split ':')[0]
                return $driveLetter
            } else {
                throw "Falsche Pfadangabe. Der Pfad muss mit einem Laufwerksbuchstaben gefolgt von einem Doppelpunkt und Backslash beginnen (z. B. C:\)."
            }
        }

        try {
            $driveLetterSync = Get-DriveLetter -Path $roboCopyBackupPath
        } catch {
            Write-Host -foregroundcolor red "Fehler bei der Verarbeitung von '$roboCopyBackupPath': $_"
            Write-Log "Fehler bei der Verarbeitung von '$roboCopyBackupPath': $_" "ERROR"
            return
        }

        $excludeFiles = $excludeFiles | Select-Object -Unique
        $excludeDirectories = $excludeDirectories | Select-Object -Unique
        $sourceDirectories = $sourceDirectories | Select-Object -Unique

        $quotedFiles = @()
        foreach ($file in $excludeFiles) {
            $quotedFiles += '"' + $file + '"'
        }
        $singleLineFiles = $quotedFiles -join ' '

        $quotedDirectories = @()
        foreach ($dir in $excludeDirectories) {
            $quotedDirectories += '"' + $dir + '"'
        }
        $singleLineDirectories = $quotedDirectories -join ' '

        # Robocopy-Log-Ordner im neuen Log-Pfad anlegen
        if (!(Test-Path $robocopyLogDir)) {
            New-Item -ItemType Directory -Path $robocopyLogDir -Force | Out-Null
        }

        # Blacklist für Laufwerkswurzeln
        $blacklist = @('A:\', 'B:\', 'C:\', 'D:\', 'E:\', 'F:\', 'G:\', 'H:\', 'I:\', 'J:\',
                       'K:\', 'L:\', 'M:\', 'N:\', 'O:\', 'P:\', 'Q:\', 'R:\', 'S:\', 'T:\',
                       'U:\', 'V:\', 'W:\', 'X:\', 'Y:\', 'Z:\')

        function IsPathInBlacklist([string]$path) {
            $normalizedPath = $path.TrimEnd('\') + '\'
            foreach ($root in $blacklist) {
                if ($normalizedPath -eq $root) {
                    return $true
                }
            }
            return $false
        }

        foreach ($path in $sourceDirectories) {
            if (IsPathInBlacklist $path) {
                Write-Host -ForegroundColor Red "Die Synchronisation eines kompletten Laufwerkes ist nicht vorgesehen; es muss sich um Ordner handeln."
                Write-Log "Abbruch: Quellpfad '$path' ist eine Laufwerkswurzel." "ERROR"
                return
            } elseif (Test-Path -Path $path -PathType Leaf) {
                Write-Host -ForegroundColor Red "Die Synchronisation einer einzelnen Datei ist nicht vorgesehen."
                Write-Log "Abbruch: Quellpfad '$path' ist eine Datei." "ERROR"
                return
            } elseif ($path -eq $roboCopyBackupPath) {
                Write-Host -ForegroundColor Red "Quelle und Ziel sind identisch. Das ist so nicht vorgesehen"
                Write-Log "Abbruch: Quelle und Ziel sind identisch ('$path')." "ERROR"
                return
            } elseif (-not(Test-Path -Path $path)) {
                Write-Host -ForegroundColor Red "Der Pfad '$path' ist nicht vorhanden oder nicht erreichbar."
                Write-Log "Abbruch: Quellpfad '$path' nicht erreichbar." "ERROR"
                return
            }
        }

        # Check if the drive letter exists
        $driveExists = Get-PSDrive -Name $driveLetterSync -ErrorAction SilentlyContinue

        if ($driveExists) {
            # create backup folder if not exists
            If(!(test-path -PathType container $roboCopyBackupPath)) {
                New-Item -ItemType Directory -Path $roboCopyBackupPath -Force
            }

            # Start the timer
            $startTime = [System.Diagnostics.Stopwatch]::StartNew()

            $jobs = @()
            $maxConcurrentJobs = $maxThreads

            foreach ($source in $sourceDirectories) {
                while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge $maxConcurrentJobs) {
                    Start-Sleep -Seconds 1
                }

                # === PRE-FLIGHT CHECKS ===
                Write-Log "Starte Synchronisation von: $source" "INFO"

                if (-not (Test-Path $source)) {
                    Write-Host -ForegroundColor Red "FEHLER: Quellverzeichnis nicht gefunden: $source"
                    Write-Log "Quellverzeichnis nicht gefunden: $source" "ERROR"
                    continue
                }

                $sourceFiles = Get-ChildItem -Path $source -Recurse -File -ErrorAction SilentlyContinue
                if (-not $sourceFiles -or $sourceFiles.Count -eq 0) {
                    Write-Host -ForegroundColor Cyan "WARNUNG: Quellverzeichnis ist leer: $source"
                    Write-Log "Quellverzeichnis ist leer: $source" "WARN"
                    continue
                }

                Write-Host -ForegroundColor Green "Quelldateien gefunden: $($sourceFiles.Count) Dateien in $source"
                Write-Log "Vorbereitung Synchronisation: $($sourceFiles.Count) Dateien gefunden in $source" "INFO"

                try {
                    $testFile = Join-Path $roboCopyBackupPath "test_write_access.tmp"
                    "test" | Out-File -FilePath $testFile -ErrorAction Stop
                    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
                    Write-Log "Schreibzugriff auf Zielverzeichnis bestätigt: $roboCopyBackupPath" "INFO"
                } catch {
                    Write-Host -ForegroundColor Red "FEHLER: Kein Schreibzugriff auf Zielverzeichnis: $roboCopyBackupPath"
                    Write-Log "Kein Schreibzugriff auf Zielverzeichnis: $roboCopyBackupPath - $($_.Exception.Message)" "ERROR"
                    continue
                }

                # === ROBOCOPY-VORBEREITUNG ===
                $cleanPath = $source -creplace '^[A-Za-z]:\\', ''
                $driveLetterOfSource = [System.IO.Path]::GetPathRoot($source)
                $trimmedString = $driveLetterOfSource.Trim(':\\')
                $newPath = $cleanPath -replace '\\', '-'
                $logName = Join-Path -Path $robocopyLogDir -ChildPath ($trimmedString + "-" + $newPath + ".log")
                $destination = $roboCopyBackupPath

                # === NEUE LOGIK FÜR PRAKTIKUM-ORDNER ===
                $praktikumSource = Join-Path -Path $source -ChildPath "Praktikum"
                $praktikumDestination = Join-Path -Path $destination -ChildPath "Praktikum"

                if (Test-Path -Path $praktikumSource) {
                    Write-Log "Praktikum-Ordner gefunden - wird mit kompletter Überschreibung synchronisiert: $praktikumSource" "INFO"

                    if (Test-Path -Path $praktikumDestination) {
                        try {
                            Start-Sleep -Milliseconds 500
                            Remove-Item -Path $praktikumDestination -Recurse -Force -ErrorAction Stop
                            Write-Log "Bestehender Praktikum-Ordner wurde gelöscht: $praktikumDestination" "INFO"
                            Start-Sleep -Milliseconds 1000
                        } catch {
                            Write-Log "Fehler beim Löschen des Praktikum-Ordners: $($_.Exception.Message)" "WARN"
                        }
                    }

                    $praktikumLogName = $logName -replace ".log", "-Praktikum.log"

                    $praktikumJob = Start-Job -ScriptBlock {
                        $src = $using:praktikumSource
                        $dest = $using:praktikumDestination
                        $excFile = $using:singleLineFiles
                        $excDirectorie = $using:singleLineDirectories
                        $logPath = $using:praktikumLogName

                        $maxRetries = 3
                        $retryCount = 0
                        $success = $false

                        while ($retryCount -lt $maxRetries -and -not $success) {
                            if ($retryCount -gt 0) {
                                Start-Sleep -Seconds (2 * $retryCount)
                            }

                            $quotedLogPath = "`"$logPath`""

                            $process = Start-Process -FilePath "robocopy.exe" -ArgumentList "`"$src`" `"$dest`" /MIR /J /XJ /DCOPY:DAT /COPY:DAT /MT:8 /R:3 /W:5 /NP /V /XA:S /XF $excFile /XD $excDirectorie /TEE /UNILOG+:$quotedLogPath" -Wait -PassThru -WindowStyle Hidden

                            $exitCode = $process.ExitCode

                            if ($exitCode -le 7) {
                                $success = $true
                            } else {
                                $retryCount++
                                if ($retryCount -lt $maxRetries) {
                                    try {
                                        Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Retry $retryCount for Praktikum sync after exit code $exitCode" -ErrorAction SilentlyContinue
                                    } catch {
                                        $null = $_.Exception.Message
                                    }
                                }
                            }
                        }

                        return $exitCode
                    }

                    $jobs += $praktikumJob
                }

                $job = Start-Job -ScriptBlock {
                    $src = $using:source
                    $dest = $using:destination
                    $excFile = $using:singleLineFiles
                    $excDirectorie = $using:singleLineDirectories
                    $logPath = $using:logName

                    $maxRetries = 3
                    $retryCount = 0
                    $success = $false

                    while ($retryCount -lt $maxRetries -and -not $success) {
                        if ($retryCount -gt 0) {
                            Start-Sleep -Seconds (2 * $retryCount)
                        }

                        $quotedLogPath = "`"$logPath`""

                        $process = Start-Process -FilePath "robocopy.exe" -ArgumentList "`"$src`" `"$dest`" /XO /E /J /XJ /DCOPY:DAT /COPY:DAT /MT:8 /R:3 /W:5 /NP /V /XA:S /XF $excFile /XD $excDirectorie /XO /XX /TEE /UNILOG+:$quotedLogPath" -Wait -PassThru -WindowStyle Hidden

                        $exitCode = $process.ExitCode

                        if ($exitCode -le 7) {
                            $success = $true
                        } else {
                            $retryCount++
                            if ($retryCount -lt $maxRetries) {
                                try {
                                    Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Retry $retryCount for standard sync after exit code $exitCode" -ErrorAction SilentlyContinue
                                } catch {
                                    $null = $_.Exception.Message
                                }
                            }
                        }
                    }

                    return $exitCode
                }

                $jobs += $job
            }

            while (($jobs | Where-Object { $_.State -ne 'Completed' }).Count -gt 0) {
                Start-Sleep -Seconds 1
            }

            foreach ($job in $jobs) {
                $jobResult = Receive-Job -Job $job -Wait
                $exitCode = $jobResult

                if ($exitCode -eq 0) {
                    Write-Host " "
                    Write-Host -ForegroundColor Green "Eine Synchronisation ist nicht erforderlich."
                    Write-Log "Eine Synchronisation ist nicht erforderlich." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 1) {
                    Write-Host " "
                    Write-Host -ForegroundColor Green "Die Synchronisation wurde erfolgreich abgeschlossen."
                    Write-Log "Die Synchronisation wurde erfolgreich abgeschlossen." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 2) {
                    Write-Host " "
                    Write-Log "Es gibt zusätzliche Dateien im Zielverzeichnis, keine neuen kopiert." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 3) {
                    Write-Host " "
                    Write-Host -ForegroundColor Cyan "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis."
                    Write-Log "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 16) {
                    Write-Host " "
                    Write-Host -ForegroundColor Red "⚠ Robocopy Fehler 16: Versuche alternativen Kopiervorgang..."
                    Write-Log "Robocopy Fehler 16 - starte Fallback-Kopiervorgang" "WARN"

                    try {
                        $sourceExists = Test-Path (Join-Path $AssetRoot 'Datei-Vorlagen')
                        if ($sourceExists -and (Get-ChildItem -Path (Join-Path $AssetRoot 'Datei-Vorlagen') -ErrorAction SilentlyContinue)) {
                            if (-not (Test-Path $BackupTargetPath)) {
                                New-Item -ItemType Directory -Path $BackupTargetPath -Force | Out-Null
                            }

                            Write-Host -ForegroundColor Cyan "  ➤ Kopiere Dateien mit alternativer Methode..."
                            Copy-Item -Path (Join-Path $AssetRoot 'Datei-Vorlagen\*') -Destination $BackupTargetPath -Recurse -Force -ErrorAction Stop

                            Write-Host -foregroundcolor Green "  ✓ FALLBACK ERFOLGREICH: Alle Dateien wurden kopiert!"
                            Write-Log "Fallback-Kopiervorgang erfolgreich abgeschlossen." "INFO"

                            $copiedFiles = (Get-ChildItem -Path $BackupTargetPath -Recurse -File -ErrorAction SilentlyContinue).Count
                            Write-Host -foregroundcolor Green "  ➤ $copiedFiles Dateien erfolgreich synchronisiert"

                        } else {
                            Write-Host -foregroundcolor red "  ✗ PROBLEM: Quellverzeichnis ist leer oder nicht zugänglich."
                            Write-Log "Quellverzeichnis für Kopiervorgang nicht verfügbar." "ERROR"
                        }
                    } catch {
                        Write-Host -foregroundcolor red "  ✗ Fallback-Kopiervorgang fehlgeschlagen: $($_.Exception.Message)"
                        Write-Log "Fallback-Kopiervorgang fehlgeschlagen: $($_.Exception.Message)" "ERROR"
                    }

                    Write-Host " "
                } else {
                    Write-Host " "
                    Write-Host -foregroundcolor red "Fehlercode $exitCode. Lesen Sie dazu https://learn.microsoft.com/en-us/troubleshoot/windows-server/backup-and-storage/return-codes-used-robocopy-utility"
                    Write-Host -foregroundcolor red "Lesen Sie auch die entsprechende Log-Datei --> $robocopyLogDir"
                    Write-Log "Fehlercode $exitCode beim Robocopy-Lauf." "ERROR"
                    Write-Host " "
                }
                Remove-Job -Job $job
            }

            $startTime.Stop()
            $elapsedTime = $startTime.Elapsed
            $formattedTime = "{0:D2} Stunden, {1:D2} Minuten, {2:D2} Sekunden, {3:D3} Millisekunden" -f $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds, $elapsedTime.Milliseconds
            Write-Log "Synchronisation abgeschlossen. Dauer: $formattedTime" "INFO"
        } else {
            Write-Host -ForegroundColor Red "Das Laufwerk $driveLetterSync ist nicht vorhanden."
            Write-Log "Das Laufwerk $driveLetterSync ist nicht vorhanden." "ERROR"
        }

        Clear-OldRobocopyLogs
        Clear-OldLogs
    }
    # Ende der Funktion sync()

    function Install-SelectedOfficeTemplates {
        param(
            [Parameter(Mandatory = $true)][string]$FontName,
            [Parameter(Mandatory = $true)][int]$FontSizeWord,
            [Parameter(Mandatory = $true)][int]$FontSizeExcel
        )

        $sourceRoot = Join-Path $AssetRoot 'Datei-Vorlagen\Sonstiges\Standards'
        $backupRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "PC-Konfigurator-GUI\Backups\Vorlagen_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $templates = @{
            Word = @{ Source = "Normal-$FontName-$FontSizeWord.dotm"; Destination = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm' }
            Excel = @{ Source = "Mappe-$FontName-$FontSizeExcel.xltx"; Destination = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\Mappe.xltx' }
            Outlook = @{ Source = "NormalEmail-$FontName-$FontSizeWord.dotm"; Destination = Join-Path $env:APPDATA 'Microsoft\Templates\NormalEmail.dotm' }
        }

        foreach ($template in $templates.GetEnumerator()) {
            $sourcePath = Join-Path $sourceRoot $template.Value.Source
            if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                throw "Vorbereitete $($template.Key)-Vorlage fehlt: $sourcePath"
            }
        }

        Get-Process WINWORD, EXCEL, OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        foreach ($template in $templates.GetEnumerator()) {
            $sourcePath = Join-Path $sourceRoot $template.Value.Source
            $destinationPath = $template.Value.Destination
            New-Item -ItemType Directory -Path (Split-Path $destinationPath -Parent) -Force | Out-Null
            if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
                New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
                Copy-Item -LiteralPath $destinationPath -Destination (Join-Path $backupRoot (Split-Path $destinationPath -Leaf)) -Force -ErrorAction Stop
                Write-Log "Vorhandene $($template.Key)-Vorlage gesichert: $destinationPath" "INFO"
            }
            Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force -ErrorAction Stop
            Write-Log "Vorbereitete $($template.Key)-Vorlage übernommen: $sourcePath -> $destinationPath" "INFO"
        }
    }

    function Sync-OutlookSignatures {
        param (
            [Parameter(Mandatory = $true)]
            [string]$BaseTargetPath
        )

        function Test-DirectoryHasContent {
            param([string]$Path)

            if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
                return $false
            }

            return [bool](Get-ChildItem -Path $Path -Force -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)
        }

        $localSignaturesPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
        $backupSignaturesPath = Join-Path $BaseTargetPath 'Signaturen'

        try {
            $localHasContent = Test-DirectoryHasContent -Path $localSignaturesPath

            if ($localHasContent) {
                if (-not (Test-Path -LiteralPath $backupSignaturesPath -PathType Container)) {
                    New-Item -ItemType Directory -Path $backupSignaturesPath -Force | Out-Null
                }

                $null = robocopy "$localSignaturesPath" "$backupSignaturesPath" /E /R:2 /W:2 /XO /FFT /COPY:DAT /DCOPY:DAT /NFL /NDL /NJH /NJS /NP
                $copyToBackupExitCode = $LASTEXITCODE
                if ($copyToBackupExitCode -le 7) {
                    Write-Log "Outlook-Signaturen in Sicherungsordner synchronisiert: $backupSignaturesPath" "INFO"
                } else {
                    Write-Log "Fehler beim Sichern der Outlook-Signaturen (Robocopy ExitCode: $copyToBackupExitCode)." "WARN"
                }
            } else {
                Write-Log "Keine lokalen Outlook-Signaturen zum Sichern gefunden: $localSignaturesPath" "INFO"
            }

            $backupHasContent = Test-DirectoryHasContent -Path $backupSignaturesPath
            if (-not $backupHasContent) {
                Write-Log "Kein Signatur-Backup für Rücksicherung gefunden: $backupSignaturesPath" "INFO"
                return
            }

            if (-not (Test-Path -LiteralPath $localSignaturesPath -PathType Container)) {
                New-Item -ItemType Directory -Path $localSignaturesPath -Force | Out-Null
            }

            if (-not $localHasContent) {
                $null = robocopy "$backupSignaturesPath" "$localSignaturesPath" /E /R:2 /W:2 /FFT /COPY:DAT /DCOPY:DAT /NFL /NDL /NJH /NJS /NP
                $restoreExitCode = $LASTEXITCODE
                if ($restoreExitCode -le 7) {
                    Write-Log "Outlook-Signaturen aus Sicherungsordner wiederhergestellt: $localSignaturesPath" "INFO"
                } else {
                    Write-Log "Fehler bei der Wiederherstellung der Outlook-Signaturen (Robocopy ExitCode: $restoreExitCode)." "WARN"
                }
                return
            }

            $null = robocopy "$backupSignaturesPath" "$localSignaturesPath" /E /R:2 /W:2 /XO /XN /XC /FFT /COPY:DAT /DCOPY:DAT /NFL /NDL /NJH /NJS /NP
            $supplementExitCode = $LASTEXITCODE
            if ($supplementExitCode -le 7) {
                Write-Log "Outlook-Signaturen geprüft; fehlende Backup-Inhalte wurden bei Bedarf ergänzt." "INFO"
            } else {
                Write-Log "Fehler beim Ergänzen fehlender Signaturen aus Backup (Robocopy ExitCode: $supplementExitCode)." "WARN"
            }
        } catch {
            Write-Log "Fehler in der Signaturen-Routine: $($_.Exception.Message)" "WARN"
        }
    }

    # ===== Ziel-Laufwerk für Datei-Vorlagen (aus Wizard-Auswahl übernommen) =====
    $UserdocPath = [Environment]::GetFolderPath('MyDocuments')
    Write-Log "Documents-Pfad ermittelt: $UserdocPath" "INFO"

    ### Initialisierung der Office-Programme
    Write-Host -ForegroundColor Cyan "Bitte kurz warten. Excel, Outlook und Word werden initialisiert."
    Write-Host " "

    if (-not (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
        Start-Process "WINWORD" -WindowStyle Minimized
        Start-Sleep 5
        Stop-Process -Name "WINWORD" -Force
    }
    if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        Start-Process "EXCEL" -WindowStyle Minimized
        Start-Sleep 5
        Stop-Process -Name "EXCEL" -Force
    }
    Write-Log "Outlook-Initialisierung übersprungen: Vorlagen und Registry-Einstellungen werden ohne gestarteten Outlook-Prozess eingerichtet." "INFO"

    try {
        Stop-Process -Name "OfficeClickToRun" -Force -ErrorAction Stop
        Write-Log "OfficeClickToRun erfolgreich beendet." "INFO"
    } catch {
        Write-Log "OfficeClickToRun konnte nicht beendet werden: Zugriff verweigert." "WARN"
    }

    if ($Params.UseDocuments) {
        $driveRoot = $UserdocPath
        $targetTemplatePath = Join-Path $UserdocPath "Datei-Vorlagen"
        Write-Log "Verwende Documents-Verzeichnis: $targetTemplatePath" "INFO"
    } else {
        $driveLetter = $Params.DriveLetter
        $driveRoot = "${driveLetter}:\"
        $targetTemplatePath = Join-Path $driveRoot "Datei-Vorlagen"
        Write-Log "Verwende Laufwerk ${driveLetter}: $targetTemplatePath" "INFO"
    }
    $BackupTargetPath = $targetTemplatePath

    Write-Host -ForegroundColor Cyan "⏳ Schritt 1/7: Outlook-Signaturen werden synchronisiert..."
    Write-Host -ForegroundColor Cyan " "
    Sync-OutlookSignatures -BaseTargetPath $BackupTargetPath
    Write-Host -foregroundcolor Green "  ✓ Schritt 1/7 abgeschlossen: Outlook-Signaturen wurden geprüft und bei Bedarf synchronisiert."

    # ===== Sync-Funktion mit RoboCopy starten =====
    sync

    # ===== Einheitliche Konfiguration der Programme Excel und Word =====
    function Set-OfficeRegistrySettings {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSizeWord = 11,
            [int]$FontSizeExcel = 10,
            $wordSettings = @{
                "DeveloperTools" = 1
                "Ruler" = 1
                "ShowAllFormatting" = 1
                "VisiDrawTableDrs" = 1
                "DOC-PATH" = if (Test-Path "Z:\") { "Z:\" } else { $driveRoot }
                "PersonalTemplates" = $BackupTargetPath
                "DisableBootToOfficeStart" = 1
                "DisableBackstageOpenKeyShortcuts" = 1
                "DefaultFont" = $FontName
                "DefaultFontSize" = $FontSizeWord
                "CorrectCapsLock" = 0
                "AutoFormatApplyBulletedLists" = 0
                "AutoFormatApplyNumberedLists" = 0
                "AutoFormatCapitalizeTableCells" = 0
                "CorrectTableCells" = 0
                "PictureInsertLayout" = 1
                "Font" = $FontName
                "Fontsubstitutes" = ""
            },
            $excelSettings = @{
                "DeveloperTools" = 1
                "DefaultPath" = if (Test-Path "Z:\") { "Z:\" } else { $driveRoot }
                "PersonalTemplates" = $BackupTargetPath
                "DisableBootToOfficeStart" = 1
                "StandardFont" = $FontName
                "StandardFontSize" = $FontSizeExcel
                "Font" = "$FontName,$FontSizeExcel"
                "AltStartupPath" = $BackupTargetPath
                "AutoSaveInterval" = 5
            },
            $windowsSettings = @{
                "HideFileExt" = 0
                "Hidden" = 1
                "ShowSuperHidden" = 0
            }
        )

        $regPathWord = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
        $regPathExcel = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        $regPathWindows = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        function Set-RegistryValues ($regPath, $settings) {
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            if (Test-Path $regPath) {
                foreach ($key in $settings.Keys) {
                    try {
                        $value = $settings[$key]
                        if ($null -eq $value) {
                            Write-Log "Überspringe $key in ${regPath}: Wert ist null." "WARN"
                            continue
                        }

                        if ($value -is [int] -or $value -is [long] -or $value -is [uint32] -or $value -is [bool]) {
                            $valueType = "DWord"
                            $safeValue = [int]$value
                        } elseif ($value -is [string]) {
                            $trimmedValue = $value.Trim()
                            if ($trimmedValue -match '^\d+$') {
                                $valueType = "DWord"
                                $safeValue = [int]$trimmedValue
                            } elseif ($trimmedValue -match '%[^%]+%') {
                                $valueType = "ExpandString"
                                $safeValue = $value
                            } else {
                                $valueType = "String"
                                $safeValue = $value
                            }
                        } else {
                            $valueType = "String"
                            $safeValue = [string]$value
                        }

                        Set-ItemProperty -Path $regPath -Name $key -Value $safeValue -Type $valueType -Force -ErrorAction Stop
                        Write-Log "Gesetzt in $regPath : $key = $safeValue (Typ: $valueType)" "INFO"
                    } catch {
                        Write-Log "Fehler beim Setzen von $key in $regPath : $($_.Exception.Message)" "WARN"
                    }
                }
            } else {
                Write-Log "Pfad nicht gefunden: $regPath" "ERROR"
            }
        }

        Set-RegistryValues $regPathWord $wordSettings
        Set-RegistryValues $regPathExcel $excelSettings
        Set-RegistryValues $regPathWindows $windowsSettings

        Write-Log "Die empfohlenen Office-Einstellungen wurden erfolgreich gesetzt." "INFO"
    }

    # === ERKENNUNG / KONFIGURATION DER EXCEL-QUICK-ACCESS-TOOLBAR ===
    function Get-ExcelQuickAccessToolbarTargets {
        param (
            [string[]]$OfficeVersions = @('16.0', '15.0', '14.0')
        )

        $targets = @()

        foreach ($version in $OfficeVersions) {
            $candidatePaths = @(
                "HKCU:\Software\Microsoft\Office\$version\Excel\QAT",
                "HKCU:\Software\Microsoft\Office\$version\Excel\Ribbon",
                "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
            )

            foreach ($candidatePath in $candidatePaths) {
                $targets += [pscustomobject]@{
                    Version = $version
                    Path = $candidatePath
                    Exists = Test-Path $candidatePath
                    Strategy = if ($candidatePath -match '\\Excel\\QAT$') { 'QAT' }
                        elseif ($candidatePath -match '\\Excel\\Ribbon$') { 'Ribbon' }
                        else { 'Options' }
                }
            }
        }

        $userAppData = [Environment]::GetFolderPath('LocalApplicationData')
        $uiFileCandidates = @(
            (Join-Path $userAppData 'Microsoft\Office\Excel.officeUI'),
            (Join-Path $userAppData 'Microsoft\Office\excel.officeUI'),
            (Join-Path $userAppData 'Microsoft\Excel\Excel16.xlb'),
            (Join-Path $userAppData 'Microsoft\Excel\Excel15.xlb'),
            (Join-Path $userAppData 'Microsoft\Excel\Excel14.xlb')
        )

        foreach ($uiFile in $uiFileCandidates) {
            if ($uiFile) {
                $targets += [pscustomobject]@{
                    Version = 'FileSystem'
                    Path = $uiFile
                    Exists = Test-Path $uiFile
                    Strategy = 'OfficeUI'
                }
            }
        }

        $preferredTarget = $targets | Where-Object { $_.Exists -and $_.Strategy -eq 'QAT' } | Select-Object -First 1
        if (-not $preferredTarget) {
            $preferredTarget = $targets | Where-Object { $_.Exists } | Select-Object -First 1
        }

        return [pscustomobject]@{
            AvailableTargets = $targets
            PreferredTarget = $preferredTarget
            PreferredStrategy = if ($preferredTarget) { $preferredTarget.Strategy } else { 'NotFound' }
        }
    }

    function Install-SelectedOfficeTheme {
        param (
            [Parameter(Mandatory = $true)][string]$FontName,
            [Parameter(Mandatory = $true)][ValidateSet('INN-tegrativ', 'DBK', 'Careli')][string]$Design
        )

        $designFolders = @{
            'INN-tegrativ' = 'Designs_INN-tegrativ'
            'DBK' = 'Designs_DBK'
            'Careli' = 'Designs_Careli'
        }
        $designPrefixes = @{
            'INN-tegrativ' = 'Design_INN-tegrativ-'
            'DBK' = 'Design_DBK-'
            'Careli' = 'Design_Careli-'
        }

        $themeFileName = "$($designPrefixes[$Design])$FontName.thmx"

        $sourceCandidates = @(
            (Join-Path $AssetRoot "Datei-Vorlagen\Sonstiges\$($designFolders[$Design])\$themeFileName")
        ) | Select-Object -Unique
        $sourcePath = $sourceCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1

        if (-not $sourcePath) {
            Write-Log "Office-Theme nicht gefunden: $themeFileName" "WARN"
            return $null
        }

        $themeTargets = @(
            (Join-Path $env:APPDATA 'Microsoft\Templates\Document Themes'),
            (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\Themes')
        )

        foreach ($themeTarget in $themeTargets) {
            try {
                New-Item -ItemType Directory -Path $themeTarget -Force | Out-Null
                $targetPath = Join-Path $themeTarget $themeFileName
                Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
                Write-Log "Office-Theme kopiert: $sourcePath -> $targetPath" "INFO"
            } catch {
                Write-Log "Office-Theme konnte nicht kopiert werden nach $themeTarget`: $($_.Exception.Message)" "WARN"
            }
        }

        return (Resolve-Path -LiteralPath $sourcePath).Path
    }

    function Sync-OfficeQuickAccessToolbarTemplates {
        $templateCandidates = @(
            (Join-Path $AssetRoot 'Datei-Vorlagen\Sonstiges\Symbolleiste Schnellzugriff')
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

        $templateRoot = $templateCandidates |
            Where-Object { Test-Path -LiteralPath $_ -PathType Container } |
            Select-Object -First 1

        if (-not $templateRoot) {
            Write-Log "Keine Vorlagen für Schnellzugriffe im Ordner 'Datei-Vorlagen\Sonstiges\Symbolleiste Schnellzugriff' gefunden." "INFO"
            return $false
        }

        $officeUserRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Microsoft\Office'
        New-Item -ItemType Directory -Path $officeUserRoot -Force | Out-Null

        $copiedFiles = @()
        foreach ($fileName in @('Excel.officeUI', 'Word.officeUI')) {
            $sourcePath = Join-Path $templateRoot $fileName
            $targetPath = Join-Path $officeUserRoot $fileName

            if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                Write-Log "Vorlage für Schnellzugriffe nicht gefunden: $sourcePath" "INFO"
                continue
            }

            Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
            $copiedFiles += $fileName
            Write-Log "Vorlage für Schnellzugriffe kopiert: $sourcePath -> $targetPath" "INFO"
        }

        if ($copiedFiles.Count -gt 0) {
            Write-Log "Vorlagen für Schnellzugriffe wurden in den Office-Benutzerpfad kopiert: $($copiedFiles -join ', ')" "INFO"
            return $true
        }

        return $false
    }

    function Set-ExcelQuickAccessToolbar {
        param (
            [ValidateSet('Standard', 'Seitenansicht', 'Drucken')]
            [string]$Layout = 'Standard',
            [string[]]$OfficeVersions = @('16.0', '15.0', '14.0')
        )

        $layoutCommands = @{
            'Standard' = @('FileSave', 'Undo', 'Redo', 'ViewPageLayoutView', 'PrintPreviewAndPrint')
            'Seitenansicht' = @('ViewPageLayoutView', 'FileSave', 'Undo', 'Redo', 'PrintPreviewAndPrint')
            'Drucken' = @('PrintPreviewAndPrint', 'FileSave', 'Undo', 'Redo', 'ViewPageLayoutView')
        }

        $analysis = Get-ExcelQuickAccessToolbarTargets -OfficeVersions $OfficeVersions
        $commands = $layoutCommands[$Layout]

        $targetPath = if ($analysis.PreferredTarget -and $analysis.PreferredTarget.Path) {
            $analysis.PreferredTarget.Path
        } else {
            "HKCU:\Software\Microsoft\Office\16.0\Excel\QAT"
        }

        if ($targetPath -match '^HKCU:') {
            if (-not (Test-Path $targetPath)) {
                New-Item -Path $targetPath -Force | Out-Null
            }

            for ($index = 0; $index -lt $commands.Count; $index++) {
                $propertyName = "Command$($index + 1)"
                try {
                    Set-ItemProperty -Path $targetPath -Name $propertyName -Value $commands[$index] -Type String -Force -ErrorAction Stop
                    Write-Log "Excel-QAT gesetzt in $targetPath : $propertyName = $($commands[$index])" "INFO"
                } catch {
                    Write-Log "Excel-QAT konnte in $targetPath nicht gesetzt werden ($propertyName): $($_.Exception.Message)" "WARN"
                }
            }
        } else {
            Write-Log "Excel-QAT-Konfiguration wurde erkannt, aber kein passender Registry-Ordner gefunden. Fallback-Strategie: $($analysis.PreferredStrategy); Zielpfad: $targetPath" "INFO"
        }

        Write-Log "Excel-QAT-Strategie: $($analysis.PreferredStrategy); Layout: $Layout; Ziel: $targetPath" "INFO"

        return [pscustomobject]@{
            Layout = $Layout
            Commands = $commands
            TargetPath = $targetPath
            PreferredStrategy = $analysis.PreferredStrategy
            AvailableTargets = $analysis.AvailableTargets
        }
    }

    # === ZUSÄTZLICHE REGISTRY-EINSTELLUNGEN FÜR SCHRIFTARTEN ===
    function Set-FontRegistrySettings {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSizeWord = 11,
            [int]$FontSizeExcel = 10
        )

        Write-Log "Setze spezifische Schriftart-Registry-Einstellungen..." "INFO"

        $officeVersions = @("16.0", "15.0", "14.0")

        foreach ($version in $officeVersions) {
            $wordRegPaths = @(
                "HKCU:\Software\Microsoft\Office\$version\Word\Options",
                "HKCU:\Software\Microsoft\Office\$version\Common\LanguageResources"
            )

            foreach ($regPath in $wordRegPaths) {
                if (-not (Test-Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                }
                if (Test-Path $regPath) {
                    try {
                        Set-ItemProperty -Path $regPath -Name "DefaultFont" -Value $FontName -Type String -Force
                        Set-ItemProperty -Path $regPath -Name "DefaultFontSize" -Value $FontSizeWord -Type DWord -Force
                        Write-Log "Word-Schriftart gesetzt in $regPath" "INFO"
                    } catch {
                        Write-Log "Fehler beim Setzen der Word-Schriftart in $regPath`: $($_.Exception.Message)" "WARN"
                    }
                }
            }

            $excelRegPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
            if (-not (Test-Path $excelRegPath)) {
                New-Item -Path $excelRegPath -Force | Out-Null
            }
            if (Test-Path $excelRegPath) {
                try {
                    Set-ItemProperty -Path $excelRegPath -Name "StandardFont" -Value $FontName -Type String -Force
                    Set-ItemProperty -Path $excelRegPath -Name "StandardFontSize" -Value $FontSizeExcel -Type DWord -Force
                    Write-Log "Excel-Schriftart gesetzt in $excelRegPath" "INFO"
                } catch {
                    Write-Log "Fehler beim Setzen der Excel-Schriftart in $excelRegPath`: $($_.Exception.Message)" "WARN"
                }
            }
        }
    }

    function Test-WordComAvailability {
        $result = [pscustomobject]@{
            IsAvailable = $false
            Version = $null
            Path = $null
            Error = $null
            OfficeInstallRoot = $null
        }

        $candidatePaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office16\WINWORD.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16\WINWORD.EXE",
            "${env:ProgramFiles}\Microsoft Office\root\Office15\WINWORD.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15\WINWORD.EXE"
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($candidatePath in $candidatePaths) {
            if (Test-Path -LiteralPath $candidatePath) {
                $result.OfficeInstallRoot = Split-Path -Path $candidatePath -Parent
                break
            }
        }

        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            if ($null -ne $word) {
                $result.Version = $word.Version
                $result.Path = $word.Path
                try {
                    $word.Visible = $false
                    $result.IsAvailable = $true
                } catch {
                    $result.Error = $_.Exception.Message
                    $result.IsAvailable = $false
                } finally {
                    try {
                        $word.Quit()
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                    } catch {
                        $null = $_
                    }
                }
            }
        } catch {
            $result.Error = $_.Exception.Message
            $result.IsAvailable = $false
        }

        if (-not $result.IsAvailable -and -not $result.Error) {
            $result.Error = 'Word COM-Objekt konnte nicht initialisiert werden.'
        }

        return $result
    }

    function Set-WordAutoCorrectRegistry {
        param (
            $autoCorrectWordSettings = @{
                "AutoFormatAsYouTypeApplyNumberedLists" = 0
                "AutoFormatAsYouTypeApplyBulletedLists" = 0
                "CorrectSentenceCaps" = 1
                "AutoFormatAsYouTypeReplaceHyperlinks" = 0
                "CorrectInitialCaps" = 0
                "AutoFormatAsYouTypeReplaceQuotes" = 1
                "AutoFormatAsYouTypeReplaceSymbols" = 1
                "PasteFormattingOtherApp" = 2
                "PasteFormattingTwoDocumentsNoStyles" = 1
            }
        )

        Write-Log "Beende alle Word-Prozesse vor der Autokorrektur-Konfiguration..." "INFO"
        Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        $regPaths = @(
            "HKCU:\Software\Microsoft\Office\16.0\Word\Options",
            "HKCU:\Software\Microsoft\Office\17.0\Word\Options",
            "HKCU:\Software\Microsoft\Office\18.0\Word\Options"
        )

        $settingsApplied = 0

        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                foreach ($key in $autoCorrectWordSettings.Keys) {
                    try {
                        Set-ItemProperty -Path $regPath -Name $key -Value $autoCorrectWordSettings[$key] -Type DWord -Force
                        Write-Log "Gesetzt in $regPath : $key = $($autoCorrectWordSettings[$key])" "INFO"
                        $settingsApplied++
                    } catch {
                        Write-Log "Fehler beim Setzen von $key in $regPath : $($_.Exception.Message)" "WARN"
                    }
                }
            }
        }

        Start-Sleep -Milliseconds 500

        try {
            Write-Log "Starte Word für COM-basierte Autokorrektur-Konfiguration..." "INFO"

            $wordAvailability = Test-WordComAvailability
            if (-not $wordAvailability.IsAvailable) {
                Write-Log "Word-COM-Initialisierung fehlgeschlagen; Autokorrektur-Konfiguration wird übersprungen. Ursache: $($wordAvailability.Error)" "WARN"
                return
            }

            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false

            $autoCorrect = $word.AutoCorrect

            try {
                $autoCorrect.AutoFormatAsYouTypeApplyNumberedLists = $false
                Write-Log "Nummerierte Listen deaktiviert: $($autoCorrect.AutoFormatAsYouTypeApplyNumberedLists)" "INFO"

                $autoCorrect.AutoFormatAsYouTypeApplyBulletedLists = $false
                Write-Log "Aufzählungslisten deaktiviert: $($autoCorrect.AutoFormatAsYouTypeApplyBulletedLists)" "INFO"

                $autoCorrect.AutoFormatAsYouTypeReplaceHyperlinks = $false
                Write-Log "Hyperlink-Ersetzung deaktiviert: $($autoCorrect.AutoFormatAsYouTypeReplaceHyperlinks)" "INFO"

                $autoCorrect.CorrectInitialCaps = $false
                Write-Log "Anfangsbuchstaben-Korrektur deaktiviert: $($autoCorrect.CorrectInitialCaps)" "INFO"

                $autoCorrect.CorrectSentenceCaps = $false
                Write-Log "Satzanfang-Korrektur aktiviert: $($autoCorrect.CorrectSentenceCaps)" "INFO"

                $autoCorrect.AutoFormatAsYouTypeReplaceQuotes = $true
                $autoCorrect.AutoFormatAsYouTypeReplaceSymbols = $true

            } catch {
                Write-Log "Fehler beim Setzen einzelner AutoCorrect-Eigenschaften: $($_.Exception.Message)" "WARN"
            }

            try {
                $tempDoc = $word.Documents.Add()
                Start-Sleep -Milliseconds 500
                $tempDoc.Close($false)
                Write-Log "Temporäres Dokument zur Aktivierung der Einstellungen erstellt und geschlossen" "INFO"
            } catch {
                Write-Log "Fehler beim Erstellen des temporären Dokuments: $($_.Exception.Message)" "WARN"
            }

            Write-Log "COM-basierte Autokorrektur erfolgreich konfiguriert" "INFO"

        } catch {
            Write-Log "Fehler bei COM-basierter Autokorrektur-Konfiguration: $($_.Exception.Message)" "WARN"
            Write-Log "Fallback auf Registry-Einstellungen" "INFO"
        } finally {
            if ($word) {
                try {
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                } catch {
                    $null = $_.Exception.Message
                }
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }

        Write-Log "=== VERIFIKATION DER AUTOKORREKTUR-EINSTELLUNGEN ===" "INFO"
        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                try {
                    $currentLists = Get-ItemProperty -Path $regPath -Name "AutoFormatAsYouTypeApplyNumberedLists" -ErrorAction SilentlyContinue
                    $currentBullets = Get-ItemProperty -Path $regPath -Name "AutoFormatAsYouTypeApplyBulletedLists" -ErrorAction SilentlyContinue

                    if ($currentLists -or $currentBullets) {
                        Write-Log "Aktuelle Registry-Werte in ${regPath}:" "INFO"

                        if ($currentLists) {
                            Write-Log "  - AutoFormatAsYouTypeApplyNumberedLists: $($currentLists.AutoFormatAsYouTypeApplyNumberedLists)" "INFO"
                        }
                        if ($currentBullets) {
                            Write-Log "  - AutoFormatAsYouTypeApplyBulletedLists: $($currentBullets.AutoFormatAsYouTypeApplyBulletedLists)" "INFO"
                        }
                    }
                } catch {
                    Write-Log "Fehler beim Lesen der Verifikationswerte: $($_.Exception.Message)" "WARN"
                }
                break
            }
        }

        if ($settingsApplied -eq 0) {
            Write-Log "WARNUNG: Keine Word Registry-Pfade gefunden - möglicherweise ist Office nicht korrekt installiert" "WARN"
        } else {
            Write-Log "Autokorrektureinstellungen für Word wurden verarbeitet ($settingsApplied Einstellungen gesetzt)" "INFO"
            Write-Log "HINWEIS: Starten Sie Word neu, um alle Änderungen zu aktivieren" "INFO"
        }
    }

    function Set-ExcelAutoCorrectRegistry {
        param (
            [hashtable]$autoCorrectExcelSettings = @{
                "CorrectSentenceCap" = 0
            }
        )

        $regPaths = @(
            "HKCU:\Software\Microsoft\Office\16.0\Excel\Options",
            "HKCU:\Software\Microsoft\Office\15.0\Excel\Options",
            "HKCU:\Software\Microsoft\Office\14.0\Excel\Options"
        )

        $settingsApplied = 0
        foreach ($regPath in $regPaths) {
            if (-not (Test-Path $regPath)) {
                continue
            }

            foreach ($key in $autoCorrectExcelSettings.Keys) {
                try {
                    Set-ItemProperty -Path $regPath -Name $key -Value $autoCorrectExcelSettings[$key] -Type DWord -Force -ErrorAction Stop
                    Write-Log "Gesetzt in $regPath : $key = $($autoCorrectExcelSettings[$key])" "INFO"
                    $settingsApplied++
                } catch {
                    Write-Log "Fehler beim Setzen von $key in ${regPath}: $($_.Exception.Message)" "WARN"
                }
            }
        }

        if ($settingsApplied -eq 0) {
            Write-Log "Kein Excel-Registry-Pfad für die Autokorrektur gefunden." "WARN"
        }

        $excel = $null
        try {
            Write-Log "Starte Excel für die COM-basierte Autokorrektur-Konfiguration..." "INFO"
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.AutoCorrect.CorrectSentenceCap = $false

            if ($excel.AutoCorrect.CorrectSentenceCap -eq $false) {
                Write-Log "Excel-Autokorrektur verifiziert: CorrectSentenceCap = False" "INFO"
            } else {
                Write-Log "Excel-Autokorrektur konnte nicht verifiziert werden: CorrectSentenceCap ist weiterhin aktiviert." "WARN"
            }
        } catch {
            Write-Log "Excel-COM-Autokorrektur konnte nicht gesetzt werden: $($_.Exception.Message)" "WARN"
        } finally {
            if ($excel) {
                try {
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                } catch {
                    Write-Log "Excel-COM-Objekt konnte nicht vollständig beendet werden: $($_.Exception.Message)" "WARN"
                }
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }

        Write-Log "Autokorrektureinstellungen für Excel wurden für den aktuellen Benutzer verarbeitet." "INFO"
    }

    function Set-ExplorerPrivacyOptions {
        $recentPath = Join-Path $env:APPDATA "Microsoft\Windows\Recent"
        if (Test-Path $recentPath) {
            try {
                $recentShortcutFiles = Get-ChildItem -Path $recentPath -File -Force -ErrorAction Stop

                if ($recentShortcutFiles -and $recentShortcutFiles.Count -gt 0) {
                    $recentShortcutFiles | Remove-Item -Force -ErrorAction Stop
                    Write-Log "Datei-Explorer-Verlauf (nur oberste Recent-Dateien) wurde gelöscht, Schnellzugriff-Pins bleiben erhalten: $recentPath" "INFO"
                } else {
                    Write-Log "Keine löschbaren Recent-Dateien im Explorer-Verlauf gefunden: $recentPath" "INFO"
                }
            } catch {
                Write-Log "Fehler beim schonenden Löschen des Datei-Explorer-Verlaufs: $($_.Exception.Message)" "WARN"
            }
        } else {
            Write-Log "Verlaufspfad nicht gefunden: $recentPath" "WARN"
        }

        [void](Set-ExplorerRecentAndFrequentState -ShowRecent 0 -ShowFrequent 0)
    }

    function Set-ExplorerSearchOptions {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Search"
        $settings = @{
            "SearchSystemDirs"      = 1
            "SearchCompressedFiles" = 1
            "SearchAlways"          = 1
        }

        if ($DryRun) {
            foreach ($key in $settings.Keys) {
                Write-DryRunAction "Explorer-Suchoption: $key = $($settings[$key]) in $regPath"
            }
            return $true
        }

        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }

        $allValid = $true
        foreach ($key in $settings.Keys) {
            try {
                Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key] -Type DWord -Force
                Write-Log "Explorer-Suchoption gesetzt: $key = $($settings[$key])" "INFO"

                if (-not (Assert-RegistryValue -Path $regPath -Name $key -Expected $settings[$key] -Description "Explorer-Suchoption $key")) {
                    $allValid = $false
                }
            } catch {
                Write-Log ("Fehler beim Setzen von $key in " + $regPath + ": " + $_.Exception.Message) "WARN"
                $allValid = $false
            }
        }

        return $allValid
    }

    function Set-OutlookRegistry {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSize = 11,
            [hashtable]$OutlookCalendarSettings = @{
                "WeekNum" = 1
            }
        )

        $newOutlookInstalled = $false
        try {
            $newOutlookInstalled = [bool](Get-AppxPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue)
        } catch {
            Write-Log "Neue Outlook-App konnte nicht erkannt werden: $($_.Exception.Message)" "WARN"
        }
        $newOutlookRunning = [bool](Get-Process -Name 'olk' -ErrorAction SilentlyContinue)
        $classicOutlookRunning = [bool](Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
        if ($newOutlookInstalled -or $newOutlookRunning) {
            Write-Log "Neue Outlook-App erkannt (installiert=$newOutlookInstalled, aktiv=$newOutlookRunning). Lokale MailSettings/NormalEmail.dotm gelten nur für klassisches Outlook; die neue Outlook-App muss über ihre eigenen Microsoft-365-Einstellungen konfiguriert werden." "WARN"
        }
        if ($classicOutlookRunning) {
            Write-Log "Klassisches Outlook erkannt; lokale Outlook-Schriftwerte werden gesetzt." "INFO"
        }

        if ([string]::IsNullOrWhiteSpace($FontName)) { $FontName = "Aptos" }
        if ($FontSize -lt 1) { $FontSize = 11 }

        function Set-OutlookMailSettingsFonts {
            param (
                [Parameter(Mandatory = $true)][string]$Path,
                [Parameter(Mandatory = $true)][string]$FontName,
                [Parameter(Mandatory = $true)][int]$FontSize
            )

            if (-not (Test-Path $Path)) {
                New-Item -Path $Path -Force | Out-Null
            }

            # ComposeFontComplex steuert neue HTML-Nachrichten; ReplyFontComplex
            # Antworten/Weiterleitungen. TextFontComplex gilt für Nur-Text-Mails.
            foreach ($name in @('ComposeFontComplex', 'ReplyFontComplex', 'TextFontComplex')) {
                $existingComplex = (Get-ItemProperty -Path $Path -Name $name -ErrorAction SilentlyContinue).$name
                if ($existingComplex -is [byte[]] -and $existingComplex.Length -gt 0) {
                    $complexFont = [System.Text.Encoding]::UTF8.GetString($existingComplex)
                    $complexFont = [regex]::Replace($complexFont, 'font-size:\s*[\d.]+pt', "font-size:$($FontSize).0pt")
                    $complexFont = [regex]::Replace($complexFont, 'font-family:\s*(?:"[^"]+"|[^;\r\n]+)', "font-family:`"$FontName`",sans-serif")
                } else {
                    $complexFont = "<html>`r`n<style>`r`n p.MsoPlainText, li.MsoPlainText, div.MsoPlainText { font-size:$($FontSize).0pt; font-family:`"$FontName`",sans-serif; }`r`n</style>`r`n</html>`r`n"
                }
                $complexBytes = [System.Text.Encoding]::UTF8.GetBytes($complexFont)
                Set-ItemProperty -Path $Path -Name $name -Value $complexBytes -Type Binary -Force
                Write-Log "Outlook-MailSettings aktualisiert: $name = $FontName / $FontSize pt" "INFO"
            }

            foreach ($name in @('ComposeFontSimple', 'ReplyFontSimple', 'TextFontSimple')) {
                $existing = (Get-ItemProperty -Path $Path -Name $name -ErrorAction SilentlyContinue).$name
                if ($existing -is [byte[]] -and $existing.Length -gt 26) {
                    $bytes = [byte[]]$existing.Clone()
                    $ascii = [System.Text.Encoding]::ASCII
                    $fontOffset = 26
                    $replacementBytes = $ascii.GetBytes($FontName)
                    if ($replacementBytes.Length -le ($bytes.Length - $fontOffset)) {
                        [Array]::Clear($bytes, $fontOffset, $bytes.Length - $fontOffset)
                        [Array]::Copy($replacementBytes, 0, $bytes, $fontOffset, $replacementBytes.Length)
                        Set-ItemProperty -Path $Path -Name $name -Value $bytes -Type Binary -Force
                        Write-Log "Outlook-MailSettings aktualisiert: $name = $FontName" "INFO"
                    } else {
                        Write-Log "Outlook-MailSettings $name nicht geändert: Fontname ist für den Binärbereich zu lang." "WARN"
                    }
                }
            }
        }

        $officeVersions = @("16.0", "15.0", "14.0")
        foreach ($version in $officeVersions) {
            $regPathOutlookCalendarHKCU = "HKCU:\Software\Microsoft\Office\$version\Outlook\Options\Calendar"
            if ($DryRun) {
                foreach ($key in $OutlookCalendarSettings.Keys) {
                    Write-DryRunAction "Outlook-Kalender: $key = $($OutlookCalendarSettings[$key]) in $regPathOutlookCalendarHKCU"
                }
            } else {
                if (-not (Test-Path $regPathOutlookCalendarHKCU)) {
                    New-Item -Path $regPathOutlookCalendarHKCU -Force | Out-Null
                }
                foreach ($key in $OutlookCalendarSettings.Keys) {
                    Set-ItemProperty -Path $regPathOutlookCalendarHKCU -Name $key -Value $OutlookCalendarSettings[$key] -Type DWord -Force
                    Write-Log "Gesetzt in $regPathOutlookCalendarHKCU : $key = $($OutlookCalendarSettings[$key])" "INFO"
                    $null = Assert-RegistryValue -Path $regPathOutlookCalendarHKCU -Name $key -Expected $OutlookCalendarSettings[$key] -Description "Outlook Kalender $key"
                }
            }

            $regPathOutlookOptionsHKCU = "HKCU:\Software\Microsoft\Office\$version\Outlook\Options"
            if ($DryRun) {
                Write-DryRunAction "Outlook-Schrift: NewMailFont = $FontName / NewMailFontSize = $FontSize in $regPathOutlookOptionsHKCU"
                Write-DryRunAction "Outlook-Schrift: ReplyForwardFont = $FontName / ReplyForwardFontSize = $FontSize in $regPathOutlookOptionsHKCU"
                Write-DryRunAction "Outlook-Schrift: DefaultMailFont = $FontName in $regPathOutlookOptionsHKCU"
                continue
            }

            if (-not (Test-Path $regPathOutlookOptionsHKCU)) {
                New-Item -Path $regPathOutlookOptionsHKCU -Force | Out-Null
            }

            $fontSettings = @{
                "NewMailFont" = $FontName
                "NewMailFontSize" = [int]$FontSize
                "ReplyForwardFont" = $FontName
                "ReplyForwardFontSize" = [int]$FontSize
                "DefaultMailFont" = $FontName
            }

            foreach ($entry in $fontSettings.GetEnumerator()) {
                $regType = if ($entry.Key -match "Size$") { "DWord" } else { "String" }
                Set-ItemProperty -Path $regPathOutlookOptionsHKCU -Name $entry.Key -Value $entry.Value -Type $regType -Force
                Write-Log "Outlook-Schrift gesetzt in $regPathOutlookOptionsHKCU : $($entry.Key) = $($entry.Value)" "INFO"
                $null = Assert-RegistryValue -Path $regPathOutlookOptionsHKCU -Name $entry.Key -Expected $entry.Value -Description "Outlook Schrift $($entry.Key)"
            }

            $mailSettingsPath = "HKCU:\Software\Microsoft\Office\$version\Common\MailSettings"
            Set-OutlookMailSettingsFonts -Path $mailSettingsPath -FontName $FontName -FontSize $FontSize

            Write-Log "Outlook-Schrift gesetzt in $regPathOutlookOptionsHKCU : $FontName / $FontSize pt" "INFO"
        }

        Write-Log "Einstellungen für Outlook wurden für den aktuellen Benutzer verarbeitet." "INFO"
    }

    function Install-Fonts {
        param (
            [string]$SourceDirectory
        )

        $targetFolder = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"

        if (-Not (Test-Path -Path $SourceDirectory)) {
            Write-Log "Quellverzeichnis '$SourceDirectory' existiert nicht. Bitte überprüfen Sie den Pfad." "ERROR"
            return
        }

        New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null

        $fontFiles = Get-ChildItem -Path $SourceDirectory -Recurse -Filter "*.ttf"
        $fontFiles += Get-ChildItem -Path $SourceDirectory -Recurse -Filter "*.otf"

        foreach ($fontFile in $fontFiles) {
            try {
                $targetFontPath = Join-Path -Path $targetFolder -ChildPath $fontFile.Name

                Copy-Item -Path $fontFile.FullName -Destination $targetFolder -Force
                Write-Log "Schriftart aktualisiert: $($fontFile.FullName) -> $targetFontPath" "INFO"

            } catch {
                Write-Log "Fehler beim Kopieren von $($fontFile.FullName): $_" "ERROR"
            }
        }

        Write-Log "Die empfohlenen Schriftarten wurden erfolgreich installiert." "INFO"
    }

    function Install-UserFonts {
        param (
            [string]$fontPath = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"
        )

        if (-Not (Test-Path -Path $fontPath)) {
            Write-Log "Der angegebene Ordner existiert nicht: $fontPath" "ERROR"
            return
        }

        $fonts = Get-ChildItem -Path $fontPath -Filter *.ttf -ErrorAction SilentlyContinue

        if ($fonts.Count -eq 0) {
            Write-Log "Keine Schriftdateien im Ordner gefunden: $fontPath" "ERROR"
            return
        }

        foreach ($font in $fonts) {
            try {
                $fontName = $font.Name
                $fontRegistryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
                New-ItemProperty -Path $fontRegistryPath -Name $fontName -Value $font.FullName -PropertyType String -ErrorAction SilentlyContinue
            } catch {
                $null = $_.Exception.Message
            }
        }
        Write-Log "Die neuen Schriftarten wurden für den angemeldeten Benutzer registriert." "INFO"
    }

    # Aufruf: Excel-Autokorrektur wird unabhängig von der Schriftart-Auswahl früh gesetzt (wie im Original).
    Set-ExcelAutoCorrectRegistry

    Set-ExplorerPrivacyOptions
    Set-DesktopInQuickAccess
    Set-ExplorerSearchOptions

    Write-Host "   "
    Write-Host -ForegroundColor Green "Die empfohlenen Anpassungen für Ihre Arbeitsumgebung wurden erfolgreich vorgenommen."
    Write-Host "   "

    ### Optional: Word und Excel neu starten, damit die Änderungen wirksam werden
    Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue

    Install-Fonts -SourceDirectory (Join-Path $AssetRoot 'Fonts')
    Install-UserFonts

    # ===== Im Wizard gewählte Werte übernehmen (statt Read-Host) =====
    $selectedDesign = $Params.SelectedDesign
    Write-Log "Gewähltes Corporate Design: $selectedDesign" "INFO"

    if ($Params.IndividualFonts) {
        $FontName = $Params.FontName
        $ActualFontName = $FontName
        $FontSizeWord = [int]$Params.FontSizeWord
        $FontSizeExcel = [int]$Params.FontSizeExcel
        Write-Log "Gewählte Schriftart (individuell): $FontName" "INFO"
        Write-Log "Gewählte Schriftgröße für Word/Outlook: $FontSizeWord Punkte" "INFO"
        Write-Log "Gewählte Schriftgröße für Excel: $FontSizeExcel Punkte" "INFO"
    } else {
        Write-Host -ForegroundColor Cyan "Schriftart- und Schriftgrößenauswahl übersprungen. Es werden Standardwerte verwendet."
        Write-Host -ForegroundColor Cyan "Schriftart: Aptos, Word/Outlook: 11 Punkte, Excel: 10 Punkte."
        $FontName = "Aptos"
        $ActualFontName = $FontName
        $FontSizeWord = 11
        $FontSizeExcel = 10
        Write-Log "Schriftart/Größen-Auswahl übersprungen. Standardwerte: $FontName, Word: $FontSizeWord, Excel: $FontSizeExcel" "INFO"
    }

    # Erst jetzt sind die finalen (gewählten oder Standard-)Werte zuverlässig vorhanden.
    try {
        Install-SelectedOfficeTemplates -FontName $ActualFontName -FontSizeWord $FontSizeWord -FontSizeExcel $FontSizeExcel
        Write-Host -ForegroundColor Green "Vorbereitete Word-, Excel- und Outlook-Vorlagen wurden übernommen."
    } catch {
        Write-Log "Vorbereitete Office-Vorlagen konnten nicht übernommen werden: $($_.Exception.Message)" "ERROR"
        throw
    }
    $selectedOfficeThemePath = Install-SelectedOfficeTheme -FontName $ActualFontName -Design $selectedDesign
    Set-OfficeRegistrySettings -FontName $ActualFontName -FontSizeWord $FontSizeWord -FontSizeExcel $FontSizeExcel
    $templateSyncResult = Sync-OfficeQuickAccessToolbarTemplates
    if ($templateSyncResult) {
        Write-Log "Vorlagen für Schnellzugriffe aus 'Datei-Vorlagen\Sonstiges\Symbolleiste Schnellzugriff' wurden übernommen." "INFO"
        Write-Host -ForeGroundColor Cyan " "
        Write-Host -ForeGroundColor Cyan "Symbolleisten für den Schnellzugriff für Excel und Word wurden erfolgreich übernommen."
    }
    $excelQatConfig = Set-ExcelQuickAccessToolbar -Layout 'Standard'
    Write-Log "Excel-QAT wurde mit Layout '$($excelQatConfig.Layout)' konfiguriert. Strategie: $($excelQatConfig.PreferredStrategy), Ziel: $($excelQatConfig.TargetPath)" "INFO"
    Set-FontRegistrySettings -FontName $ActualFontName -FontSizeWord $FontSizeWord -FontSizeExcel $FontSizeExcel

    # Outlook bekommt bewusst die gleichen Werte wie Word
    Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Set-OutlookRegistry -FontName $ActualFontName -FontSize $FontSizeWord
    Write-Log "Outlook-Schrift wurde mit Word synchronisiert: $ActualFontName / $FontSizeWord pt" "INFO"

    # === OFFICE-KONFIGURATION STARTEN ===
    Write-Host ""
    Write-Host -ForegroundColor Cyan "Starte Office-Konfiguration..."
    Write-Host -foregroundcolor Cyan "Dies kann 1-2 Minuten dauern. Bitte haben Sie Geduld."
    Write-Host ""

    # --- Funktionen zum Anpassen der Office-Vorlagen ---

    function Set-WordDocumentTheme {
        param (
            [Parameter(Mandatory = $true)]$Document,
            [string]$ThemePath
        )

        if ([string]::IsNullOrWhiteSpace($ThemePath) -or -not (Test-Path -LiteralPath $ThemePath -PathType Leaf)) {
            return $false
        }

        try {
            $Document.ApplyTheme($ThemePath)
            Write-Log "Word-Theme eingebettet: $ThemePath" "INFO"
            return $true
        } catch {
            Write-Log "Word-Theme konnte nicht eingebettet werden: $($_.Exception.Message)" "WARN"
            return $false
        }
    }

    function Set-ExcelWorkbookTheme {
        param (
            [Parameter(Mandatory = $true)]$Workbook,
            [string]$ThemePath
        )

        if ([string]::IsNullOrWhiteSpace($ThemePath) -or -not (Test-Path -LiteralPath $ThemePath -PathType Leaf)) {
            return $false
        }

        try {
            $Workbook.ApplyTheme($ThemePath)
            Write-Log "Excel-Theme eingebettet: $ThemePath" "INFO"
            return $true
        } catch {
            Write-Log "Excel-Theme konnte nicht eingebettet werden: $($_.Exception.Message)" "WARN"
            return $false
        }
    }

    function Set-OutlookTemplateTheme {
        param (
            [Parameter(Mandatory = $true)][string]$ThemePath,
            [string]$FontName = "Aptos",
            [int]$FontSize = 11
        )

        $templatePath = Join-Path $env:APPDATA 'Microsoft\Templates\NormalEmail.dotm'
        if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) {
            Write-Log "NormalEmail.dotm für Theme-Anpassung nicht gefunden: $templatePath" "WARN"
            return
        }

        $word = $null
        $template = $null
        try {
            Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            $template = $word.Documents.Open($templatePath, $false, $false)
            [void](Set-WordDocumentTheme -Document $template -ThemePath $ThemePath)
            $standardStyle = $template.Styles.Item("Standard")
            $standardStyle.Font.Name = $FontName
            $standardStyle.Font.Size = $FontSize
            try {
                $normalStyle = $template.Styles.Item("Normal")
                $normalStyle.Font.Name = $FontName
                $normalStyle.Font.Size = $FontSize
            } catch {
                Write-Log "Normal-Style in NormalEmail.dotm konnte nicht angepasst werden: $($_.Exception.Message)" "WARN"
            }
            $template.Save()
            $template.Close($false)
            Write-Log "Outlook-Theme und Schriftart in NormalEmail.dotm übernommen: $FontName / $FontSize pt." "INFO"
        } catch {
            Write-Log "Outlook-Theme konnte nicht in NormalEmail.dotm übernommen werden: $($_.Exception.Message)" "WARN"
            if ($template) { try { $template.Close($false) } catch { $null = $_.Exception.Message } }
        } finally {
            if ($word) {
                try { $word.Quit() } catch { $null = $_.Exception.Message }
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null } catch { $null = $_.Exception.Message }
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    function Set-WordHeadingStyles {
        param (
            [Parameter(Mandatory = $true)]$Styles,
            [string]$FontName = "Aptos"
        )

        $headingStyles = @(
            @{ Id = -2; Name = 'Überschrift 1'; Size = 16; BottomBorder = $true },
            @{ Id = -3; Name = 'Überschrift 2'; Size = 14; BottomBorder = $true },
            @{ Id = -4; Name = 'Überschrift 3'; Size = 12; BottomBorder = $false },
            @{ Id = -5; Name = 'Überschrift 4'; Size = 11; BottomBorder = $false }
        )

        foreach ($heading in $headingStyles) {
            try {
                $style = $Styles.Item($heading.Id)
                $style.Font.Name = $FontName
                $style.Font.Size = $heading.Size
                $style.Font.Color = 0 # wdColorBlack

                $style.Borders.Enable = $false
                if ($heading.BottomBorder) {
                    $bottomBorder = $style.Borders.Item(3) # wdBorderBottom
                    $bottomBorder.LineStyle = 1 # wdLineStyleSingle
                    $bottomBorder.LineWidth = 4 # wdLineWidth050pt
                    $bottomBorder.Color = 0 # wdColorBlack
                    $bottomBorder.Visible = $true
                }

                Write-Log "$($heading.Name) angepasst: $($heading.Size) pt, Schwarz, untere Linie = $($heading.BottomBorder)" "INFO"
            } catch {
                Write-Log "$($heading.Name) konnte nicht angepasst werden: $($_.Exception.Message)" "WARN"
            }
        }
    }

    function Set-WordCustomizer {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSize = 11
        )

        # Hinweis: Diese Funktion wird - wie im Original-Konsolenskript - nicht direkt
        # aufgerufen. Die eigentliche Normal.dotm-Anpassung erfolgt inline in Schritt 2/7
        # (siehe unten). Die Funktion bleibt für Vollständigkeit/Kompatibilität erhalten.
        Write-Log "Set-WordCustomizer gestartet..." "INFO"
        Write-Host -ForegroundColor Cyan "  ➤ Word wird initialisiert..."

        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1

        $word = $null
        $wordtemplate = $null
        $standardStyle = $null

        $timeout = 30
        $timer = [Diagnostics.Stopwatch]::StartNew()

        try {
            Write-Host -ForegroundColor Cyan "  ➤ Word COM-Objekt wird gestartet..."
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            $word.DisplayAlerts = 0
            Write-Log "Word COM-Objekt erfolgreich gestartet." "INFO"
            Write-Host -ForegroundColor Green "  ✓ Word erfolgreich gestartet"
        } catch {
            Write-Log "Fehler beim Starten von Word COM-Objekt: $($_.Exception.Message)" "ERROR"
            Write-Host -ForegroundColor Red "  ✗ Word konnte nicht gestartet werden: $($_.Exception.Message)"
            return
        }

        Write-Host -ForegroundColor Cyan "  ➤ Normal.dotm wird angepasst..."
        $wordtemplatePath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
        if (Test-Path $wordtemplatePath) {
            try {
                if ($timer.Elapsed.TotalSeconds -gt $timeout) {
                    throw "Timeout erreicht beim Word-Template-Zugriff"
                }

                $wordtemplate = $word.Documents.Open($wordtemplatePath, $false, $false)
                [void](Set-WordDocumentTheme -Document $wordtemplate -ThemePath $selectedOfficeThemePath)
                $standardStyle = $wordtemplate.Styles.Item("Standard")
                $standardStyle.ParagraphFormat.SpaceAfter = 0
                $wordtemplate.DefaultTabStop = 1.0 * 28.35
                $standardStyle.Font.Name = $ActualFontName
                $standardStyle.Font.Size = $FontSize
                $standardStyle.ParagraphFormat.LineSpacingRule = 3  # wdLineSpaceMultiple
                $standardStyle.ParagraphFormat.LineSpacing = 1.1    # 1.1-fach

                try {
                    $normalStyle = $wordtemplate.Styles.Item("Normal")
                    $normalStyle.Font.Name = $ActualFontName
                    $normalStyle.Font.Size = $FontSize
                    $normalStyle.ParagraphFormat.SpaceAfter = 0
                    $normalStyle.ParagraphFormat.LineSpacingRule = 3
                    $normalStyle.ParagraphFormat.LineSpacing = 1.1
                    Write-Log "Normal-Style erfolgreich angepasst." "INFO"
                } catch {
                    Write-Log "Fehler beim Anpassen des Normal-Styles: $($_.Exception.Message)" "WARN"
                }

                $wordtemplate.Save()
                Write-Host -ForegroundColor Green "  ✓ Normal.dotm erfolgreich angepasst"
                $wordtemplate.Close($false)

                if ($standardStyle) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($standardStyle) | Out-Null
                    $standardStyle = $null
                }
                if ($wordtemplate) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordtemplate) | Out-Null
                    $wordtemplate = $null
                }

                Write-Log "Normal.dotm wurde erfolgreich angepasst." "INFO"
            } catch {
                Write-Log "Fehler beim Anpassen der Normal.dotm: $($_.Exception.Message)" "ERROR"
                if ($standardStyle) {
                    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($standardStyle) | Out-Null } catch { $null = $_.Exception.Message }
                    $standardStyle = $null
                }
                if ($wordtemplate) {
                    try { $wordtemplate.Close($false) } catch { $null = $_.Exception.Message }
                    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordtemplate) | Out-Null } catch { $null = $_.Exception.Message }
                    $wordtemplate = $null
                }
            }
        }

        Write-Host -ForegroundColor Cyan "  ➤ Word-Optionen werden gesetzt..."
        try {
            if ($word.Options) {
                $word.Options.DefaultFont = $ActualFontName
                $word.Options.DefaultFontSize = $FontSize
                try {
                    $word.Options.DefaultOpenFormat = 0  # Word-Format
                } catch {
                    $null = $_.Exception.Message
                }
                Write-Log "Word-Standard-Schriftart über COM-Objekt gesetzt: $ActualFontName" "INFO"
                Write-Host -ForegroundColor Green "  ✓ Word-Optionen erfolgreich gesetzt"
            }
        } catch {
            Write-Log "Fehler beim Setzen der Standard-Schriftart über COM: $($_.Exception.Message)" "WARN"
            Write-Host -ForegroundColor Cyan "  ⚠ Word-Optionen konnten nicht gesetzt werden"
        }

        try {
            if ($word.Selection) {
                $word.Selection.Font.Name = $ActualFontName
                $word.Selection.Font.Size = $FontSize
                $word.Selection.ParagraphFormat.SpaceAfter = 0
                $word.Selection.ParagraphFormat.LineSpacingRule = 3
                $word.Selection.ParagraphFormat.LineSpacing = 1.1
                Write-Log "Word-Selection-Default gesetzt" "INFO"
            }
        } catch {
            Write-Log "Fehler beim Setzen der Selection-Defaults: $($_.Exception.Message)" "WARN"
        }

        try {
            if ($word.NormalTemplate) {
                $normalTemplate = $word.NormalTemplate
                $normalStyle = $normalTemplate.Styles.Item("Standard")
                $normalStyle.Font.Name = $ActualFontName
                $normalStyle.Font.Size = $FontSize
                $normalStyle.ParagraphFormat.SpaceAfter = 0
                $normalStyle.ParagraphFormat.LineSpacingRule = 3  # wdLineSpaceMultiple
                $normalStyle.ParagraphFormat.LineSpacing = 1.1   # 1.1-fach
                $normalTemplate.Save()

                if ($normalStyle) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($normalStyle) | Out-Null
                    $normalStyle = $null
                }
                if ($normalTemplate) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($normalTemplate) | Out-Null
                    $normalTemplate = $null
                }

                Write-Log "NormalTemplate direkt angepasst und gespeichert." "INFO"
            }
        } catch {
            Write-Log "Fehler beim direkten Anpassen des NormalTemplate: $($_.Exception.Message)" "WARN"
        }

        try {
            $testDoc = $word.Documents.Add()
            $testRange = $testDoc.Range(0, 0)

            $testRange.Font.Name = $ActualFontName
            $testRange.Font.Size = $FontSize
            $testRange.ParagraphFormat.SpaceAfter = 0
            $testRange.ParagraphFormat.LineSpacingRule = 3
            $testRange.ParagraphFormat.LineSpacing = 1.1

            $testDoc.AttachedTemplate.Save()
            $testDoc.Close($false)

            Write-Log "Test-Dokument mit Standard-Formatierung erstellt und Normal.dotm aktualisiert." "INFO"
        } catch {
            Write-Log "Fehler beim Erstellen des Test-Dokuments: $($_.Exception.Message)" "WARN"
        }

        try {
            if ($word) {
                $word.Quit($false, $false, $false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                $word = $null
            }

            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()

            Start-Sleep -Milliseconds 500

            Write-Log "Word COM-Objekt erfolgreich geschlossen und freigegeben." "INFO"
        } catch {
            Write-Log "Fehler beim Schließen des Word COM-Objekts: $($_.Exception.Message)" "WARN"
        }
    }

    function Set-WordLernsituationenCustomizer {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSize = 11
        )

        Write-Log "Set-WordLernsituationenCustomizer gestartet ..." "INFO"
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            Write-Log "Word COM-Objekt erfolgreich gestartet." "INFO"
        } catch {
            Write-Log "Fehler beim Starten von Word COM-Objekt: $($_.Exception.Message)" "ERROR"
            Write-Host -ForegroundColor Red "⚠ Anpassung von Word-Lernsituationen übersprungen: Word konnte nicht gestartet werden - COM-Fehler."
            return
        }

        $wordtemplatePath = Join-Path $driveRoot "Datei-Vorlagen\Duisdorfer BüroKonzept KG\Lernsituationen\Lernsituationen DBK.dotx"
        if (-not (Test-Path $wordtemplatePath)) {
            Write-Log "Lernsituationen DBK.dotx nicht gefunden unter: $wordtemplatePath" "ERROR"
            $word.Quit()
            return
        }

        try {
            try {
                $wordtemplate = $word.Documents.Open($wordtemplatePath)
            } catch {
                Write-Log "Fehler beim Öffnen der Lernsituationen DBK.dotx: $($_.Exception.Message)" "ERROR"
                $word.Quit()
                return
            }
            [void](Set-WordDocumentTheme -Document $wordtemplate -ThemePath $selectedOfficeThemePath)
            $standardStyle = $wordtemplate.Styles.Item("Standard")
            $standardStyle.ParagraphFormat.SpaceAfter = 0
            $wordtemplate.DefaultTabStop = 1.0 * 28.35
            $standardStyle.Font.Name = $FontName
            $standardStyle.Font.Size = $FontSize
            $standardStyle.ParagraphFormat.LineSpacingRule = 3  # wdLineSpaceMultiple
            $standardStyle.ParagraphFormat.LineSpacing = 1.1   # 1.1-fach
            Set-WordHeadingStyles -Styles $wordtemplate.Styles -FontName $FontName
            $wordtemplate.Save()
            $wordtemplate.Close($false)
            Write-Log "Die Lernsituationen DBK.dotx wurde erfolgreich angepasst." "INFO"
        } catch {
            Write-Log "Fehler beim Ändern der Lernsituationen DBK.dotx: $($_.Exception.Message)" "ERROR"
        } finally {
            if ($word) {
                $word.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    function Set-ExcelCustomizer {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSize = 10
        )

        Write-Log "Set-ExcelCustomizer gestartet..." "INFO"

        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            try {
                $excel.StandardFont = $FontName
                $excel.StandardFontSize = $FontSize
                Write-Log "Excel-Standard-Schriftart über COM-Objekt gesetzt." "INFO"
            } catch {
                Write-Log "Fehler beim Setzen der Standard-Schriftart über COM: $($_.Exception.Message)" "WARN"
            }

            $excelTemplatePath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Excel\XLSTART\Mappe.xltx')

            if (Test-Path $excelTemplatePath) {
                try {
                    $workbook = $excel.Workbooks.Open($excelTemplatePath, $null, $false)
                    [void](Set-ExcelWorkbookTheme -Workbook $workbook -ThemePath $selectedOfficeThemePath)

                    $style = $workbook.Styles.Item("Normal")
                    $style.Font.Name = $FontName
                    $style.Font.Size = $FontSize

                    foreach ($worksheet in $workbook.Worksheets) {
                        $worksheet.Cells.Font.Name = $FontName
                        $worksheet.Cells.Font.Size = $FontSize
                    }

                    $temporaryTemplatePath = Join-Path ([IO.Path]::GetDirectoryName($excelTemplatePath)) ("Mappe.xltx.$([guid]::NewGuid()).tmp")
                    $workbook.SaveAs($temporaryTemplatePath, 54) # Excel Template Format
                    $workbook.Close($false)
                    Move-Item -LiteralPath $temporaryTemplatePath -Destination $excelTemplatePath -Force
                    Write-Log "Mappe.xltx wurde erfolgreich angepasst." "INFO"
                } catch {
                    Write-Log "Fehler beim Anpassen der Mappe.xltx: $($_.Exception.Message)" "ERROR"
                }
            }

            try {
                $newWorkbook = $excel.Workbooks.Add()
                [void](Set-ExcelWorkbookTheme -Workbook $newWorkbook -ThemePath $selectedOfficeThemePath)

                $normalStyle = $newWorkbook.Styles.Item("Normal")
                $normalStyle.Font.Name = $FontName
                $normalStyle.Font.Size = $FontSize

                $worksheet = $newWorkbook.Worksheets.Item(1)
                $worksheet.Cells.Font.Name = $FontName
                $worksheet.Cells.Font.Size = $FontSize

                $alternativeTemplatePath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Excel\XLSTART\book.xltx')
                Remove-Item $alternativeTemplatePath -Force -ErrorAction SilentlyContinue
                $newWorkbook.SaveAs($alternativeTemplatePath, 54)
                $newWorkbook.Close($false)
                Write-Log "Alternative Excel-Vorlage (book.xltx) erstellt." "INFO"
            } catch {
                Write-Log "Fehler beim Erstellen der alternativen Vorlage: $($_.Exception.Message)" "WARN"
            }

            Write-Log "Excel-Customizer erfolgreich abgeschlossen." "INFO"
        } catch {
            Write-Log "Fehler in Set-ExcelCustomizer: $($_.Exception.Message)" "ERROR"
        } finally {
            if ($excel) {
                try {
                    $excel.DisplayAlerts = $true
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                } catch {
                    Write-Log "Fehler beim Schließen von Excel: $($_.Exception.Message)" "WARN"
                }
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    # Aufruf der Funktionen NACH der Schriftart- und Schriftgrößen-Auswahl
    Write-Host -ForegroundColor Cyan "⏳ Schritt 2/7: Word-Vorlagen werden angepasst..."
    Write-Host -ForegroundColor Cyan " "
    if ($selectedOfficeThemePath) {
        Set-WordCustomizer -FontName $ActualFontName -FontSize $FontSizeWord
        Write-Host -foregroundcolor Green "  ✓ Schritt 2/7 abgeschlossen: Word-Vorlage und Corporate Design wurden übernommen."
    } else {
        Write-Log "Schritt 2/7: Kein Corporate-Design verfügbar; die vorbereitete Normal.dotm wurde ohne Theme-Einbettung übernommen." "WARN"
        Write-Host -ForegroundColor Cyan "  ⚠ Schritt 2/7 abgeschlossen mit Hinweis: Corporate Design konnte nicht eingebettet werden."
    }

    Write-Host -ForegroundColor Cyan "⏳ Schritt 3/7: Word-Lernsituationen werden angepasst..."
    Write-Host -ForegroundColor Cyan " "
    Set-WordLernsituationenCustomizer -FontName $ActualFontName -FontSize $FontSizeWord
    Write-Host -foregroundcolor Green "  ✓ Schritt 3/7 abgeschlossen: Word-Lernsituationen wurden verarbeitet."

    Write-Host -ForegroundColor Cyan "⏳ Schritt 4/7: Excel-Vorlagen werden angepasst..."
    Write-Host -ForegroundColor Cyan " "
    if ($selectedOfficeThemePath) {
        Set-ExcelCustomizer -FontName $ActualFontName -FontSize $FontSizeExcel
        Set-OutlookTemplateTheme -ThemePath $selectedOfficeThemePath -FontName $ActualFontName -FontSize $FontSizeWord
        Write-Host -foregroundcolor Green "  ✓ Schritt 4/7 abgeschlossen: Excel-, Outlook-Vorlagen und Corporate Design wurden übernommen."
    } else {
        Write-Log "Schritt 4/7: Kein Corporate-Design verfügbar; Mappe.xltx und NormalEmail.dotm wurden ohne Theme-Einbettung übernommen." "WARN"
        Write-Host -ForegroundColor Cyan "  ⚠ Schritt 4/7 abgeschlossen mit Hinweis: Corporate Design konnte nicht eingebettet werden."
    }

    Write-Host -ForegroundColor Cyan "⏳ Schritt 5/7: Registry-Einstellungen werden gesetzt..."
    Write-Host -ForegroundColor Cyan " "
    Set-WordAutoCorrectRegistry
    Write-Host -foregroundcolor Green "  ✓ Schritt 5/7 abgeschlossen: Registry-Einstellungen wurden verarbeitet."

    Write-Host -ForegroundColor Cyan "⏳ Schritt 6/7: Windows-Einstellungen werden angepasst..."
    Write-Host -ForegroundColor Cyan " "

    # Windows 11 - Taskleisteneinstellungen anpassen
    function Set-TaskbarSettings {
        param (
            [ValidateSet("Left", "Center")]
            [string]$Alignment = "Left",
            [ValidateSet("Hidden", "Icon", "Box")]
            [string]$Search = "Icon"
        )

        $taskbarRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        $searchRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        $taskbarValueName = "TaskbarAl"
        $searchValueName = "SearchboxTaskbarMode"

        function Set-RegistryDwordSafe {
            param(
                [Parameter(Mandatory = $true)][string]$Path,
                [Parameter(Mandatory = $true)][string]$Name,
                [Parameter(Mandatory = $true)][int]$Value
            )

            try {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWord -Force -ErrorAction Stop
                Write-Log "Registry gesetzt: $Path -> $Name = $Value" "INFO"
                return $true
            } catch {
                Write-Log "Registry-Wert konnte nicht gesetzt werden ($Path -> $Name): $($_.Exception.Message)" "WARN"
                return $false
            }
        }

        try {
            if (-not (Test-Path $taskbarRegistryPath)) {
                New-Item -Path $taskbarRegistryPath -Force | Out-Null
            }
            if (-not (Test-Path $searchRegistryPath)) {
                New-Item -Path $searchRegistryPath -Force | Out-Null
            }

            if ($Alignment -eq "Left") {
                $taskbarValue = 0
            } else {
                $taskbarValue = 1
            }

            switch ($Search) {
                "Hidden" { $searchValue = 0 }
                "Icon" { $searchValue = 1 }
                "Box" { $searchValue = 2 }
            }

            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name $taskbarValueName -Value $taskbarValue)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "TaskbarDa" -Value 0)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_Layout" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_TrackProgs" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_TrackDocs" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_ShowDocuments" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_ShowDownloads" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_ShowNetwork" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_ShowFileExplorer" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_ShowSettings" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name "Start_ShowPowerButton" -Value 1)
            [void](Set-RegistryDwordSafe -Path $taskbarRegistryPath -Name $searchValueName -Value $searchValue)
            [void](Set-RegistryDwordSafe -Path $searchRegistryPath -Name $searchValueName -Value $searchValue)
            [void](Set-RegistryDwordSafe -Path $searchRegistryPath -Name "SearchboxTaskbarModeCache" -Value $searchValue)

            Write-Log "Taskbar-Einstellungen erfolgreich gesetzt: Alignment=$Alignment, Search=$Search" "INFO"

        } catch {
            Write-Log "Fehler beim Setzen der Taskbar-Einstellungen: $($_.Exception.Message)" "ERROR"
        }
    }

    $taskbarAlignment = if ($Params.TaskbarAlignment -eq 'Left') { 'Left' } else { 'Center' }
    Set-TaskbarSettings -Alignment $taskbarAlignment -Search "Icon"

    Write-Host -foregroundcolor Green "  ✓ Schritt 6/7 abgeschlossen: Windows-Einstellungen wurden verarbeitet."

    Write-Host "   "
    Write-Host -ForegroundColor Cyan "⏳ Schritt 7/7: Verknüpfung für 'Kontaktdaten DBK.xlsx' wird erstellt..."
    Write-Host -ForegroundColor Cyan " "

    # --- Shortcut für Kontaktdaten DBK.xlsx im Benutzer-Ordner erstellen ---
    $step7Succeeded = $false
    try {
        $sourceFile = Join-Path $BackupTargetPath 'Duisdorfer BüroKonzept KG\Datenquellen\Kontaktdaten DBK.xlsx'
        if (-not (Test-Path $sourceFile)) {
            $fallbackSource = Join-Path $AssetRoot 'Datei-Vorlagen\Duisdorfer BüroKonzept KG\Datenquellen\Kontaktdaten DBK.xlsx'
            if (Test-Path $fallbackSource) {
                Write-Log "Kontaktdaten DBK.xlsx nicht im benutzerspezifischen Vorlagenpfad gefunden. Fallback auf Skriptpfad wird verwendet." "WARN"
                $sourceFile = $fallbackSource
            } else {
                throw "Kontaktdaten DBK.xlsx wurde weder im benutzerspezifischen Vorlagenpfad noch im Skriptpfad gefunden."
            }
        }
        $documentsPath = [Environment]::GetFolderPath('MyDocuments')
        $targetDir = Join-Path $documentsPath 'Meine Datenquellen'
        $shortcutPath = Join-Path $targetDir 'Kontaktdaten DBK.lnk'

        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        $wshShell = New-Object -ComObject WScript.Shell
        $shortcut = $wshShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $sourceFile
        $shortcut.WorkingDirectory = Split-Path $sourceFile
        $shortcut.WindowStyle = 1
        $shortcut.Description = 'Kontaktdaten Duisdorfer BüroKonzept KG'
        $shortcut.Save()
        $step7Succeeded = $true
        Write-Host -ForegroundColor Green "  Verknüpfung zu 'Kontaktdaten DBK.xlsx' wurde in '$targetDir' erstellt."
    } catch {
        Write-Host -ForegroundColor Red "Fehler beim Erstellen der Verknüpfung: $($_.Exception.Message)"
        Write-Log "Fehler beim Erstellen der Verknüpfung: $($_.Exception.Message)" "ERROR"
    }

    if ($step7Succeeded) {
        Write-Host -foregroundcolor Green "  ✓ Schritt 7/7 abgeschlossen: Verknüpfung wurde erstellt."
    } else {
        Write-Host -ForegroundColor Cyan "  ⚠ Schritt 7/7 abgeschlossen mit Hinweis: Verknüpfung konnte nicht erstellt werden."
    }

    Write-Log "=== Invoke-PCKonfiguratorPipeline abgeschlossen ===" "INFO"

    return [pscustomobject]@{
        Success = $true
        BackupTargetPath = $BackupTargetPath
    }
}

# ============================================================================
# Invoke-PCKonfiguratorFinalize
# ----------------------------------------------------------------------------
# Wird NACH Bestätigung auf der Abschluss-Seite des Assistenten ausgeführt:
# optionaler Explorer-Neustart sowie Aufräumarbeiten (Praktikum-Ordner).
# Ebenfalls im Hintergrund über eine Runspace ausgeführt, da ein Explorer-
# Neustart laut Original bis zu 30-60 Sekunden dauern kann.
# ============================================================================
function Invoke-PCKonfiguratorFinalize {
    param(
        [Parameter(Mandatory = $true)]$Params
    )

    $logDir = $Params.LogDir
    $logFile = $Params.LogFile
    $UiQueue = $Params.UiQueue
    $BackupTargetPath = $Params.BackupTargetPath
    $RestartExplorer = [bool]$Params.RestartExplorer

    function Write-Host {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
            [object[]]$Object,
            [ConsoleColor]$ForegroundColor,
            [ConsoleColor]$BackgroundColor,
            [string]$Separator,
            [switch]$NoNewline
        )
        try {
            $text = ($Object -join ' ')
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $UiQueue.Enqueue($text)
            }
        } catch {
            $null = $_.Exception.Message
        }
    }

    function Write-Log {
        param(
            [Parameter(Mandatory=$true)][string]$Message,
            [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
        )
        try {
            if (-not (Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $entry = "$timestamp [$Level] - $Message"
            Add-Content -Path $logFile -Value $entry -Encoding UTF8
        } catch {
            Write-Host -ForegroundColor Cyan "LOGFALLBACK [$Level] $Message"
        }
        try {
            $UiQueue.Enqueue("[$Level] $Message")
        } catch {
            $null = $_.Exception.Message
        }
    }

    function Restart-ExplorerIfRunning {
        $explorerRunning = Get-Process -Name "explorer" -ErrorAction SilentlyContinue

        if ($explorerRunning) {
            Write-Log "Explorer-Neustart wird vorbereitet - sichere geöffnete Fenster" "INFO"

            $openWindows = @()

            try {
                $shell = New-Object -ComObject Shell.Application
                foreach ($window in $shell.Windows()) {
                    if ($window.Name -match "Windows Explorer|File Explorer|Explorer" -or $window.FullName -match "explorer\.exe") {
                        try {
                            $currentPath = $null

                            if ($window.LocationURL) {
                                $currentPath = $window.LocationURL -replace "file:///", "" -replace "/", "\"
                                $currentPath = $currentPath -replace "%20", " "
                                $currentPath = $currentPath -replace "%C3%A4", "ä"
                                $currentPath = $currentPath -replace "%C3%B6", "ö"
                                $currentPath = $currentPath -replace "%C3%BC", "ü"
                                $currentPath = $currentPath -replace "%C3%9F", "ß"
                            }

                            if (-not $currentPath -and $window.Document -and $window.Document.Folder) {
                                $currentPath = $window.Document.Folder.Self.Path
                            }

                            if (-not $currentPath -and $window.LocationName) {
                                $locationName = $window.LocationName
                                if ($locationName -notmatch "Dieser PC|This PC|Computer|Arbeitsplatz") {
                                    $currentPath = $locationName
                                }
                            }

                            if ($currentPath -and $currentPath -ne "" -and (Test-Path $currentPath -ErrorAction SilentlyContinue)) {
                                if ($openWindows -notcontains $currentPath) {
                                    $openWindows += $currentPath
                                    Write-Log "Gespeichert: Explorer-Fenster für Pfad '$currentPath'" "INFO"
                                }
                            }
                        } catch {
                            Write-Log "Fehler beim Ermitteln des Pfads für ein Explorer-Fenster: $($_.Exception.Message)" "WARN"
                        }
                    }
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            } catch {
                Write-Log "Fehler beim Speichern der Explorer-Fenster: $($_.Exception.Message)" "WARN"
            }

            Write-Log "Stoppe Windows-Explorer ($(($openWindows).Count) Fenster gespeichert)" "INFO"

            Write-Host -ForegroundColor Cyan "Stoppe Datei-Explorer..."
            Stop-Process -Name "explorer" -Force

            Start-Sleep -Seconds 1

            Write-Host -ForegroundColor Cyan "Starte Datei-Explorer neu..."
            Start-Process "explorer.exe" -WindowStyle Hidden

            Start-Sleep -Seconds 2

            try {
                $newWindows = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
                if ($newWindows) {
                    Write-Host -ForegroundColor Cyan "Schließe automatisch geöffnete Fenster..."
                    Start-Sleep -Milliseconds 500

                    $shell = New-Object -ComObject Shell.Application
                    $windowsToClose = @()

                    foreach ($window in $shell.Windows()) {
                        if ($window.Name -match "Windows Explorer|File Explorer|Explorer" -or $window.FullName -match "explorer\.exe") {
                            $windowsToClose += $window
                        }
                    }

                    foreach ($window in $windowsToClose) {
                        try {
                            $window.Quit()
                        } catch {
                            $null = $_.Exception.Message
                        }
                    }

                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
                    Start-Sleep -Milliseconds 500
                }
            } catch {
                Write-Log "Fehler beim Schließen automatischer Explorer-Fenster: $($_.Exception.Message)" "WARN"
            }

            if ($openWindows.Count -gt 0) {
                Write-Log "Stelle $($openWindows.Count) Explorer-Fenster wieder her" "INFO"
                Write-Host -ForegroundColor Cyan "Stelle zuvor geöffnete Fenster des Datei-Explorers wieder her..."

                foreach ($path in $openWindows) {
                    try {
                        if (Test-Path $path -ErrorAction SilentlyContinue) {
                            Start-Process "explorer.exe" -ArgumentList "`"$path`""
                            Write-Log "Wiederhergestellt: Explorer-Fenster für '$path'" "INFO"
                            Start-Sleep -Milliseconds 200
                        } else {
                            Write-Log "Pfad nicht mehr verfügbar: '$path'" "WARN"
                        }
                    } catch {
                        Write-Log "Fehler beim Wiederherstellen von '$path': $($_.Exception.Message)" "ERROR"
                    }
                }

                Write-Host -foregroundcolor Green "Wiederherstellung der Fenster des Datei-Explorers abgeschlossen"
                Write-Log "Explorer-Fenster-Wiederherstellung abgeschlossen" "INFO"
            } else {
                Write-Log "Keine Explorer-Fenster zum Wiederherstellen gefunden" "INFO"
            }

        } else {
            Write-Log "Der Windows-Explorer ist aktuell nicht aktiv - kein Neustart erforderlich." "INFO"
        }
    }

    function Remove-PraktikumOrdnerBenutzer {
        $benutzerVorlagenPfad = $BackupTargetPath
        $praktikumPfad = Join-Path $benutzerVorlagenPfad 'Praktikum'
        if (Test-Path $praktikumPfad) {
            try {
                Remove-Item -Path $praktikumPfad -Recurse -Force -ErrorAction Stop
                Write-Log "Ordner 'Praktikum' im Benutzer-Vorlagenverzeichnis wurde entfernt." "INFO"
            } catch {
                Write-Log "Fehler beim Entfernen des Ordners 'Praktikum': $($_.Exception.Message)" "WARN"
            }
        }
    }

    if ($RestartExplorer) {
        Write-Host -ForegroundColor Cyan "Starte den Datei-Explorer neu ..."
        Restart-ExplorerIfRunning
    } else {
        Write-Host -ForegroundColor Cyan "Explorer-Neustart übersprungen. Änderungen werden beim nächsten Windows-Neustart aktiv."
        Write-Log "Explorer-Neustart vom Benutzer übersprungen" "INFO"
    }

    Remove-PraktikumOrdnerBenutzer

    Write-Host "   "
    Write-Host -ForegroundColor Green "Der PC-Konfigurator hat Ihren Rechner konfiguriert."
    Write-Log "=== Invoke-PCKonfiguratorFinalize abgeschlossen ===" "INFO"

    return [pscustomobject]@{
        Success = $true
    }
}

# ============================================================================
# Windows-Presentation-Foundation-Einrichtungsassistent (Wizard)
# ----------------------------------------------------------------------------
# XAML wird bewusst als Here-String direkt im Skript eingebettet (nicht aus
# einer externen .xaml-Datei geladen), da eine mit ps2exe kompilierte EXE
# keine zusätzlichen Dateien zur Laufzeit nachladen kann.
# ============================================================================
[xml]$xamlDefinition = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PC-Konfigurator-GUI" Height="620" Width="820"
        WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="PC-Konfigurator-GUI" FontSize="22" FontWeight="Bold"/>
            <TextBlock x:Name="txtStepIndicator" Text="Schritt 1 von 6: Begrüßung" FontSize="12" Foreground="#555555" Margin="0,2,0,0"/>
        </StackPanel>

        <Grid Grid.Row="1">

            <!-- Seite A: Begrüßung -->
            <StackPanel x:Name="PanelWelcome" Visibility="Visible">
                <TextBlock TextWrapping="Wrap" Margin="0,0,0,12" FontSize="13">
Willkommen beim PC-Konfigurator-GUI-Einrichtungsassistenten.
Dieser Assistent richtet Ihre Office-Vorlagen, Schriftarten, Ihr Corporate Design
und weitere Windows-/Office-Einstellungen automatisch für Sie ein.
                </TextBlock>
                <TextBlock TextWrapping="Wrap" Margin="0,0,0,12" FontWeight="Bold">
Bevor Sie fortfahren, speichern Sie bitte alle geöffneten Office-Dateien.
                </TextBlock>
                <CheckBox x:Name="chkOfficeClosed" Content="Haben Sie alle Office-Dateien gespeichert und die Office-Programme Excel, Word und Outlook geschlossen?" Margin="0,0,0,20" FontSize="13"/>
                <Button x:Name="btnWelcomeNext" Content="Weiter" Width="140" Height="32" HorizontalAlignment="Right" IsEnabled="False"/>
            </StackPanel>

            <!-- Seite B: Zielauswahl -->
            <StackPanel x:Name="PanelTarget" Visibility="Collapsed">
                <TextBlock Text="Wählen Sie den Zielort für die Datei-Vorlagen:" FontWeight="Bold" Margin="0,0,0,12" FontSize="13"/>
                <RadioButton x:Name="radDrive" GroupName="Target" Content="Auf einem bestimmten Laufwerk (z. B. Netzlaufwerk)" IsChecked="True" Margin="0,0,0,6"/>
                <StackPanel Orientation="Horizontal" Margin="24,0,0,12">
                    <TextBlock Text="Laufwerksbuchstabe:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBox x:Name="txtDriveLetter" Width="40" MaxLength="1" Text="Z"/>
                </StackPanel>
                <RadioButton x:Name="radDocuments" GroupName="Target" Content="Im Dokumente-Ordner des angemeldeten Benutzers" Margin="0,0,0,20"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnTargetBack" Content="Zurück" Width="140" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnTargetNext" Content="Weiter" Width="140" Height="32"/>
                </StackPanel>
            </StackPanel>

            <!-- Seite C: Corporate Design -->
            <StackPanel x:Name="PanelDesign" Visibility="Collapsed">
                <TextBlock Text="Wählen Sie das gewünschte Corporate Design:" FontWeight="Bold" Margin="0,0,0,12" FontSize="13"/>
                <RadioButton x:Name="radDesign1" GroupName="Design" Content="INN-tegrativ" IsChecked="True" Margin="0,0,0,4"/>
                <TextBlock Text="Design-Vorlage für die Marke INN-tegrativ." Margin="24,0,0,12" Foreground="#555555" TextWrapping="Wrap"/>
                <RadioButton x:Name="radDesign2" GroupName="Design" Content="Duisdorfer BüroKonzept (DBK)" Margin="0,0,0,4"/>
                <TextBlock Text="Design-Vorlage für die Marke Duisdorfer BüroKonzept KG." Margin="24,0,0,12" Foreground="#555555" TextWrapping="Wrap"/>
                <RadioButton x:Name="radDesign3" GroupName="Design" Content="Careli" Margin="0,0,0,4"/>
                <TextBlock Text="Design-Vorlage für die Marke Careli." Margin="24,0,0,20" Foreground="#555555" TextWrapping="Wrap"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnDesignBack" Content="Zurück" Width="140" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnDesignNext" Content="Weiter" Width="140" Height="32"/>
                </StackPanel>
            </StackPanel>

            <!-- Seite D: Schriftart und Schriftgrößen -->
            <StackPanel x:Name="PanelFont" Visibility="Collapsed">
                <TextBlock Text="Schriftart und Schriftgrößen" FontWeight="Bold" Margin="0,0,0,12" FontSize="13"/>
                <TextBlock TextWrapping="Wrap" Margin="0,0,0,12">
Standardmäßig werden Aptos (Word/Outlook: 11 Punkt, Excel: 10 Punkt) verwendet.
Aktivieren Sie das Kontrollkästchen, um eine individuelle Auswahl zu treffen.
                </TextBlock>
                <CheckBox x:Name="chkIndividualFonts" Content="Schriftart und Schriftgrößen individuell wählen" IsChecked="False" Margin="0,0,0,16"/>

                <StackPanel x:Name="PanelFontDetails" IsEnabled="False" Margin="24,0,0,0">
                    <TextBlock Text="Schriftart:" Margin="0,0,0,4"/>
                    <ComboBox x:Name="cmbFontName" Width="220" HorizontalAlignment="Left" Margin="0,0,0,12" SelectedIndex="0">
                        <ComboBoxItem Content="Aptos"/>
                        <ComboBoxItem Content="Aptos Narrow"/>
                        <ComboBoxItem Content="Arial"/>
                        <ComboBoxItem Content="Calibri"/>
                        <ComboBoxItem Content="Futura"/>
                        <ComboBoxItem Content="PT Sans"/>
                        <ComboBoxItem Content="Roboto"/>
                        <ComboBoxItem Content="Segoe UI"/>
                    </ComboBox>

                    <TextBlock Text="Schriftgröße Word/Outlook:" Margin="0,0,0,4"/>
                    <ComboBox x:Name="cmbFontSizeWord" Width="120" HorizontalAlignment="Left" Margin="0,0,0,12" SelectedIndex="1">
                        <ComboBoxItem Content="10"/>
                        <ComboBoxItem Content="11"/>
                        <ComboBoxItem Content="12"/>
                    </ComboBox>

                    <TextBlock Text="Schriftgröße Excel:" Margin="0,0,0,4"/>
                    <ComboBox x:Name="cmbFontSizeExcel" Width="120" HorizontalAlignment="Left" Margin="0,0,0,12" SelectedIndex="0">
                        <ComboBoxItem Content="10"/>
                        <ComboBoxItem Content="11"/>
                        <ComboBoxItem Content="12"/>
                    </ComboBox>
                </StackPanel>

                <TextBlock Text="Taskleisten-Ausrichtung" FontWeight="Bold" Margin="0,8,0,6"/>
                <TextBlock Text="Wählen Sie, wie die Symbole auf der Windows-Taskleiste ausgerichtet werden sollen." TextWrapping="Wrap" Margin="0,0,0,6" Foreground="#555555"/>
                <RadioButton x:Name="radTaskbarCenter" GroupName="TaskbarAlignment" Content="Zentriert (Windows-Standard)" IsChecked="True" Margin="0,0,0,4"/>
                <RadioButton x:Name="radTaskbarLeft" GroupName="TaskbarAlignment" Content="Linksbündig" Margin="0,0,0,8"/>

                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
                    <Button x:Name="btnFontBack" Content="Zurück" Width="140" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnFontNext" Content="Weiter" Width="140" Height="32"/>
                </StackPanel>
            </StackPanel>

            <!-- Seite E: Ausführung / Fortschritt -->
            <Grid x:Name="PanelExecution" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Text="Konfiguration wird ausgeführt" FontWeight="Bold" Margin="0,0,0,8" FontSize="13"/>
                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
                    <Button x:Name="btnStartPipeline" Content="Start" Width="140" Height="32" Margin="0,0,10,0"/>
                    <TextBlock x:Name="lblExecutionStatus" Text="Bereit zum Start." VerticalAlignment="Center"/>
                </StackPanel>
                <Border Grid.Row="2" BorderBrush="#CCCCCC" BorderThickness="1" Margin="0,0,0,8">
                    <ScrollViewer x:Name="scrollLog" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                        <TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                                 FontFamily="Consolas" FontSize="11" BorderThickness="0" Background="#FAFAFA"/>
                    </ScrollViewer>
                </Border>
                <ProgressBar x:Name="progBar" Grid.Row="3" Height="22" Minimum="0" Maximum="7" Value="0"/>
            </Grid>

            <!-- Seite F: Abschluss -->
            <StackPanel x:Name="PanelFinish" Visibility="Collapsed">
                <TextBlock Text="Konfiguration abgeschlossen" FontWeight="Bold" Margin="0,0,0,12" FontSize="14" Foreground="#1a7d1a"/>
                <TextBlock TextWrapping="Wrap" Margin="0,0,0,12">
Die Konfiguration Ihres Rechners wurde erfolgreich durchgeführt.
Damit alle Änderungen im Windows-Explorer sichtbar werden (z. B. Schnellzugriffe),
wird ein Neustart des Datei-Explorers empfohlen.
                </TextBlock>
                <CheckBox x:Name="chkRestartExplorer" Content="Windows-Explorer jetzt neu starten" IsChecked="True" Margin="0,0,0,20"/>
                <Button x:Name="btnFinish" Content="Fertig" Width="140" Height="32" HorizontalAlignment="Right"/>
            </StackPanel>

        </Grid>
    </Grid>
</Window>
'@

# ----------------------------------------------------------------------------
# Optionale Referenz-Datei (rein dokumentarisch): das eigentliche Fenster wird
# ausschliesslich aus dem obigen Here-String geladen, NICHT aus dieser Datei.
# ----------------------------------------------------------------------------
try {
    $xamlRefPath = Join-Path $script:ScriptRoot 'PC-Konfigurator-GUI.xaml'
    if (-not (Test-Path $xamlRefPath)) {
        $xamlDefinition.OuterXml | Out-File -FilePath $xamlRefPath -Encoding utf8 -ErrorAction SilentlyContinue
    }
} catch {
    $null = $_.Exception.Message
}

$reader = New-Object System.Xml.XmlNodeReader $xamlDefinition
$window = [Windows.Markup.XamlReader]::Load($reader)

# --- Steuerelemente referenzieren ---
$ctrl = @{}
foreach ($name in @(
        'txtStepIndicator',
        'PanelWelcome', 'chkOfficeClosed', 'btnWelcomeNext',
        'PanelTarget', 'radDrive', 'txtDriveLetter', 'radDocuments', 'btnTargetBack', 'btnTargetNext',
        'PanelDesign', 'radDesign1', 'radDesign2', 'radDesign3', 'btnDesignBack', 'btnDesignNext',
        'PanelFont', 'chkIndividualFonts', 'PanelFontDetails', 'cmbFontName', 'cmbFontSizeWord', 'cmbFontSizeExcel', 'radTaskbarCenter', 'radTaskbarLeft', 'btnFontBack', 'btnFontNext',
        'PanelExecution', 'btnStartPipeline', 'lblExecutionStatus', 'txtLog', 'scrollLog', 'progBar',
        'PanelFinish', 'chkRestartExplorer', 'btnFinish'
    )) {
    $ctrl[$name] = $window.FindName($name)
}

$script:AllPages = @('PanelWelcome', 'PanelTarget', 'PanelDesign', 'PanelFont', 'PanelExecution', 'PanelFinish')
$script:PageTitles = @{
    'PanelWelcome'   = 'Schritt 1 von 6: Begrüßung'
    'PanelTarget'    = 'Schritt 2 von 6: Zielauswahl'
    'PanelDesign'    = 'Schritt 3 von 6: Corporate Design'
    'PanelFont'      = 'Schritt 4 von 6: Schriftart und Schriftgrößen'
    'PanelExecution' = 'Schritt 5 von 6: Ausführung'
    'PanelFinish'    = 'Schritt 6 von 6: Abschluss'
}

function Show-WizardPage {
    param([Parameter(Mandatory = $true)][string]$PageName)
    foreach ($page in $script:AllPages) {
        $ctrl[$page].Visibility = if ($page -eq $PageName) { 'Visible' } else { 'Collapsed' }
    }
    $ctrl['txtStepIndicator'].Text = $script:PageTitles[$PageName]
}

# --- Seite A: Begrüßung ---
$ctrl['chkOfficeClosed'].Add_Click({
        $ctrl['btnWelcomeNext'].IsEnabled = [bool]$ctrl['chkOfficeClosed'].IsChecked
    })

$ctrl['btnWelcomeNext'].Add_Click({
        if (-not $script:TestUIMode) {
            Write-Log "Benutzer hat Office-Programme-Schließung bestätigt." "INFO"
            $sysReqOk = Test-SystemRequirements
            if (-not $sysReqOk) {
                [System.Windows.MessageBox]::Show(
                    "Die Systemanforderungen sind nicht erfüllt (erforderlich: Windows 10 Build 18362+ oder Windows 11). Der Assistent wird beendet.",
                    "Systemanforderungen nicht erfüllt", 'OK', 'Error') | Out-Null
                $window.Close()
                return
            }
        }
        Show-WizardPage 'PanelTarget'
    })

# --- Seite B: Zielauswahl ---
$ctrl['radDrive'].Add_Click({ $ctrl['txtDriveLetter'].IsEnabled = $true })
$ctrl['radDocuments'].Add_Click({ $ctrl['txtDriveLetter'].IsEnabled = $false })

$ctrl['btnTargetBack'].Add_Click({ Show-WizardPage 'PanelWelcome' })
$ctrl['btnTargetNext'].Add_Click({
        if ([bool]$ctrl['radDrive'].IsChecked) {
            $driveLetterValue = ("$($ctrl['txtDriveLetter'].Text)").Trim().TrimEnd(':')
            if ([string]::IsNullOrWhiteSpace($driveLetterValue) -or $driveLetterValue.Length -ne 1) {
                [System.Windows.MessageBox]::Show("Bitte geben Sie genau einen gültigen Laufwerksbuchstaben ein.", "Ungültige Eingabe", 'OK', 'Warning') | Out-Null
                return
            }
        }
        Show-WizardPage 'PanelDesign'
    })

# --- Seite C: Corporate Design ---
$ctrl['btnDesignBack'].Add_Click({ Show-WizardPage 'PanelTarget' })
$ctrl['btnDesignNext'].Add_Click({ Show-WizardPage 'PanelFont' })

# --- Seite D: Schriftart und Schriftgrößen ---
$ctrl['chkIndividualFonts'].Add_Click({
        $ctrl['PanelFontDetails'].IsEnabled = [bool]$ctrl['chkIndividualFonts'].IsChecked
    })
$ctrl['btnFontBack'].Add_Click({ Show-WizardPage 'PanelDesign' })
$ctrl['btnFontNext'].Add_Click({ Show-WizardPage 'PanelExecution' })

# --- Seite E: Ausführung / Fortschritt (Hintergrund-Runspace + DispatcherTimer) ---
$script:PipelinePowerShell = $null
$script:PipelineHandle = $null
$script:PipelineRunspace = $null
$script:PipelineResult = $null
$script:UiTimer = $null
$script:FinalizePowerShell = $null
$script:FinalizeHandle = $null
$script:FinalizeRunspace = $null
$script:FinalizeTimer = $null
$script:TestUiTimer = $null
$script:IsShuttingDown = $false

function Stop-PCKonfiguratorBackgroundWork {
    if ($script:IsShuttingDown) { return }
    $script:IsShuttingDown = $true

    foreach ($timerName in @('UiTimer', 'FinalizeTimer', 'TestUiTimer')) {
        $timer = Get-Variable -Name $timerName -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        if ($timer) {
            try { $timer.Stop() } catch { $null = $_.Exception.Message }
            Set-Variable -Name $timerName -Scope Script -Value $null
        }
    }

    foreach ($work in @(
            @{ PowerShell = 'PipelinePowerShell'; Handle = 'PipelineHandle'; Runspace = 'PipelineRunspace' },
            @{ PowerShell = 'FinalizePowerShell'; Handle = 'FinalizeHandle'; Runspace = 'FinalizeRunspace' }
        )) {
        $powerShell = Get-Variable -Name $work.PowerShell -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        if ($powerShell) {
            try {
                if ($powerShell.InvocationStateInfo.State -eq 'Running') {
                    $powerShell.Stop()
                }
            } catch { $null = $_.Exception.Message }
            try { $powerShell.Dispose() } catch { $null = $_.Exception.Message }
            Set-Variable -Name $work.PowerShell -Scope Script -Value $null
        }

        $runspace = Get-Variable -Name $work.Runspace -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        if ($runspace) {
            try { $runspace.Close() } catch { $null = $_.Exception.Message }
            try { $runspace.Dispose() } catch { $null = $_.Exception.Message }
            Set-Variable -Name $work.Runspace -Scope Script -Value $null
        }
        Set-Variable -Name $work.Handle -Scope Script -Value $null
    }
}

function Write-LogLine {
    param([string]$Line)
    if ([string]::IsNullOrEmpty($Line)) { return }
    $ctrl['txtLog'].AppendText("$Line`r`n")
    $ctrl['scrollLog'].ScrollToEnd()
    if ($Line -match 'Schritt\s+(\d)/7') {
        $stepNumber = [int]$Matches[1]
        $ctrl['progBar'].Value = $stepNumber
    }
}

function Start-BackgroundPipeline {
    $selectedDesignValue = if ([bool]$ctrl['radDesign1'].IsChecked) { 'INN-tegrativ' }
    elseif ([bool]$ctrl['radDesign2'].IsChecked) { 'DBK' }
    else { 'Careli' }

    $selectedFontName = if ($ctrl['cmbFontName'].SelectedItem) { $ctrl['cmbFontName'].SelectedItem.Content.ToString() } else { 'Aptos' }
    $selectedFontSizeWord = if ($ctrl['cmbFontSizeWord'].SelectedItem) { [int]$ctrl['cmbFontSizeWord'].SelectedItem.Content.ToString() } else { 11 }
    $selectedFontSizeExcel = if ($ctrl['cmbFontSizeExcel'].SelectedItem) { [int]$ctrl['cmbFontSizeExcel'].SelectedItem.Content.ToString() } else { 10 }
    $individualFonts = [bool]$ctrl['chkIndividualFonts'].IsChecked
    $taskbarAlignment = if ([bool]$ctrl['radTaskbarLeft'].IsChecked) { 'Left' } else { 'Center' }
    if (-not $individualFonts) {
        $selectedFontName = 'Aptos'
        $selectedFontSizeWord = 11
        $selectedFontSizeExcel = 10
    }

    $pipelineParams = @{
        LogDir          = $script:logDir
        LogFile         = $script:logFile
        RobocopyLogDir  = $script:robocopyLogDir
        ScriptRoot      = $script:ScriptRoot
        UiQueue         = $script:UiQueue
        DryRun          = [bool]$DryRun
        UseDocuments    = [bool]$ctrl['radDocuments'].IsChecked
        DriveLetter     = ("$($ctrl['txtDriveLetter'].Text)").Trim().TrimEnd(':')
        SelectedDesign  = $selectedDesignValue
        IndividualFonts = $individualFonts
        FontName        = $selectedFontName
        FontSizeWord    = $selectedFontSizeWord
        FontSizeExcel   = $selectedFontSizeExcel
        TaskbarAlignment = $taskbarAlignment
    }

    $ctrl['btnStartPipeline'].IsEnabled = $false
    $ctrl['lblExecutionStatus'].Text = 'Konfiguration läuft...'
    Write-Log "=== Wizard-Eingaben übernommen: Ziel=$(if ($pipelineParams.UseDocuments) { 'Dokumente' } else { "Laufwerk $($pipelineParams.DriveLetter)" }), Design=$($pipelineParams.SelectedDesign), IndividualFonts=$($pipelineParams.IndividualFonts), Taskleiste=$($pipelineParams.TaskbarAlignment) ===" "INFO"

    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $script:PipelineRunspace = [runspacefactory]::CreateRunspace($initialSessionState)
    $script:PipelineRunspace.ApartmentState = 'STA'
    $script:PipelineRunspace.ThreadOptions = 'ReuseThread'
    $script:PipelineRunspace.Open()

    $script:PipelinePowerShell = [powershell]::Create()
    $script:PipelinePowerShell.Runspace = $script:PipelineRunspace

    $pipelineFunctionText = (Get-Item function:Invoke-PCKonfiguratorPipeline).Definition
    [void]$script:PipelinePowerShell.AddScript("function Invoke-PCKonfiguratorPipeline { $pipelineFunctionText }")
    [void]$script:PipelinePowerShell.AddScript('param($Params) Invoke-PCKonfiguratorPipeline -Params $Params')
    [void]$script:PipelinePowerShell.AddArgument($pipelineParams)

    $script:PipelineHandle = $script:PipelinePowerShell.BeginInvoke()

    $script:UiTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:UiTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:UiTimer.Add_Tick({
            $lineText = $null
            while ($script:UiQueue.TryDequeue([ref]$lineText)) {
                Write-LogLine -Line $lineText
            }

            if ($script:PipelineHandle.IsCompleted) {
                $script:UiTimer.Stop()
                try {
                    $pipelineOutput = $script:PipelinePowerShell.EndInvoke($script:PipelineHandle)
                    $script:PipelineResult = $pipelineOutput | Select-Object -Last 1

                    if ($script:PipelinePowerShell.Streams.Error.Count -gt 0) {
                        foreach ($errorRecord in $script:PipelinePowerShell.Streams.Error) {
                            Write-LogLine -Line "[FEHLER] $($errorRecord.ToString())"
                        }
                        $ctrl['lblExecutionStatus'].Text = 'Konfiguration mit Fehlern beendet. Details siehe Log.'
                    } else {
                        $script:BackupTargetPath = if ($script:PipelineResult) { $script:PipelineResult.BackupTargetPath } else { $null }
                        $ctrl['lblExecutionStatus'].Text = 'Konfiguration abgeschlossen.'
                        $ctrl['progBar'].Value = 7
                        Show-WizardPage 'PanelFinish'
                    }
                } catch {
                    Write-LogLine -Line "[FEHLER] Pipeline-Ausführung fehlgeschlagen: $($_.Exception.Message)"
                    $ctrl['lblExecutionStatus'].Text = 'Konfiguration fehlgeschlagen.'
                } finally {
                    if ($script:PipelinePowerShell) {
                        $script:PipelinePowerShell.Dispose()
                        $script:PipelinePowerShell = $null
                    }
                    if ($script:PipelineRunspace) {
                        $script:PipelineRunspace.Close()
                        $script:PipelineRunspace.Dispose()
                        $script:PipelineRunspace = $null
                    }
                    $script:PipelineHandle = $null
                }
            }
        })
    $script:UiTimer.Start()
}

$ctrl['btnStartPipeline'].Add_Click({
        if ($script:TestUIMode) {
            return
        }
        Start-BackgroundPipeline
    })

# --- Seite F: Abschluss ---
function Close-PCKonfiguratorApplication {
    Stop-PCKonfiguratorBackgroundWork
    if ($window) {
        $window.Close()
    }
    $application = [System.Windows.Application]::Current
    if ($application) {
        $application.Shutdown()
    }
    [Environment]::Exit(0)
}

$window.Add_Closed({
    Stop-PCKonfiguratorBackgroundWork
    $application = [System.Windows.Application]::Current
    if ($application) {
        $application.Shutdown()
    }
})

$ctrl['btnFinish'].Add_Click({
        if ($script:TestUIMode) {
            Close-PCKonfiguratorApplication
            return
        }

        $ctrl['btnFinish'].IsEnabled = $false

        $finalizeParams = @{
            LogDir           = $script:logDir
            LogFile          = $script:logFile
            UiQueue          = $script:UiQueue
            BackupTargetPath = $script:BackupTargetPath
            RestartExplorer  = [bool]$ctrl['chkRestartExplorer'].IsChecked
        }

        $finalizeInitialState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $script:FinalizeRunspace = [runspacefactory]::CreateRunspace($finalizeInitialState)
        $script:FinalizeRunspace.ApartmentState = 'STA'
        $script:FinalizeRunspace.ThreadOptions = 'ReuseThread'
        $script:FinalizeRunspace.Open()

        $script:FinalizePowerShell = [powershell]::Create()
        $script:FinalizePowerShell.Runspace = $script:FinalizeRunspace
        $finalizeFunctionText = (Get-Item function:Invoke-PCKonfiguratorFinalize).Definition
        [void]$script:FinalizePowerShell.AddScript("function Invoke-PCKonfiguratorFinalize { $finalizeFunctionText }")
        [void]$script:FinalizePowerShell.AddScript('param($Params) Invoke-PCKonfiguratorFinalize -Params $Params')
        [void]$script:FinalizePowerShell.AddArgument($finalizeParams)

        $script:FinalizeHandle = $script:FinalizePowerShell.BeginInvoke()

        $script:FinalizeTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:FinalizeTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $script:FinalizeTimer.Add_Tick({
                $lineText = $null
                while ($script:UiQueue.TryDequeue([ref]$lineText)) {
                    Write-LogLine -Line $lineText
                }
                if ($script:FinalizeHandle -and $script:FinalizeHandle.IsCompleted) {
                    $script:FinalizeTimer.Stop()
                    try {
                        $null = $script:FinalizePowerShell.EndInvoke($script:FinalizeHandle)
                    } catch {
                        Write-LogLine -Line "[FEHLER] Abschlussarbeiten fehlgeschlagen: $($_.Exception.Message)"
                    } finally {
                        if ($script:FinalizePowerShell) {
                            $script:FinalizePowerShell.Dispose()
                            $script:FinalizePowerShell = $null
                        }
                        if ($script:FinalizeRunspace) {
                            $script:FinalizeRunspace.Close()
                            $script:FinalizeRunspace.Dispose()
                            $script:FinalizeRunspace = $null
                        }
                        $script:FinalizeHandle = $null
                        $script:FinalizeTimer = $null
                    }
                    [System.Windows.MessageBox]::Show("Der PC-Konfigurator hat Ihren Rechner erfolgreich konfiguriert.", "Fertig", 'OK', 'Information') | Out-Null
                    Close-PCKonfiguratorApplication
                }
            })
        $script:FinalizeTimer.Start()
    })

# ----------------------------------------------------------------------------
# Programmstart
# ----------------------------------------------------------------------------
$script:TestUIMode = [bool]$TestUI

if (-not $script:TestUIMode) {
    Clear-OldLogs
    Write-Log "=== PC-Konfigurator-GUI gestartet (Version 1.0.0.0) ===" "INFO"
}

Show-WizardPage 'PanelWelcome'

if ($script:TestUIMode) {
    # Reiner Rendertest: Fenster anzeigen und nach 2 Sekunden automatisch schließen,
    # OHNE jegliche echte Registry-/Datei-Änderung vorzunehmen.
    $script:TestUiTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:TestUiTimer.Interval = [TimeSpan]::FromSeconds(2)
    $script:TestUiTimer.Add_Tick({
            $script:TestUiTimer.Stop()
            $script:TestUiTimer = $null
            $window.Close()
        })
    $script:TestUiTimer.Start()
}

$window.Add_ContentRendered({
        # Nach dem vorbereitenden Laufzeitfenster startet die Anwendung in
        # einem neuen Prozess aus LocalAppData. Das Hauptfenster wird hier
        # ausdrücklich wiederhergestellt und aktiviert, damit es sichtbar im
        # Vordergrund erscheint.
        $window.WindowState = 'Normal'
        $window.Topmost = $true
        $window.Activate() | Out-Null
        $window.Focus() | Out-Null
        $window.Topmost = $false
    })

[void]$window.ShowDialog()
exit 0
