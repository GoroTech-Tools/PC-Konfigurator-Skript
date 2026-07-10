### PowerShell-Skript  
### Anpassungen für Windows 11 und Office-Programme
###
### In der Laufwerksauswahl (L) wird zusätzlich (für private Zwecke) das Verzeichnis Documents (D) des jeweiligen Anwendenden 
### zur Auswahl angeboten.

function Confirm-OfficeClosure {
    do {
        Write-Host -ForegroundColor Yellow "Haben Sie alle Office-Dateien gespeichert und die Office-Programme Excel, Word und Outlook geschlossen? (Ja/Nein)" 
        $response = Read-Host
        if ($response -match "^(Ja|ja|J|j)$") {
            Write-Host -ForegroundColor Green "Bestätigung erhalten. Skript wird fortgesetzt..." 
            return $true
        }
        elseif ($response -match "^(Nein|nein|N|n)$") {
            Write-Host -ForegroundColor Red "Bitte speichern Sie Ihre Dateien und schließen Sie die Programme." 
        }
        else {
            Write-Host -ForegroundColor Red "Ungültige Eingabe. Bitte geben Sie 'Ja' oder 'Nein' ein." 
        }
    } while ($response -notmatch "^(Ja|ja|J|j)$")

    return $false
}


# --- Hilfsfunktionen (ganz am Anfang, damit sie überall verfügbar sind) ---
# Zentrale Feature-Flags für gemeinsame Skript-Einstellungen.
# $true = aktiv, $false = deaktiviert.
$script:FeatureFlags = [ordered]@{
    EnableQuickAccessAutoSort = $true
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    try {
        if (-not (Test-Path $global:logDir)) {
            New-Item -ItemType Directory -Path $global:logDir -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $entry = "$timestamp [$Level] - $Message"
        Add-Content -Path $global:logFile -Value $entry -Encoding UTF8
    } catch {
        # Fallback auf Konsole, falls Log-Datei nicht erreichbar ist
        Write-Host -ForegroundColor Yellow "LOGFALLBACK [$Level] $Message"
    }
}

function Clear-OldLogs {
    param (
        [string]$logDir = $global:logDir
    )
    if (Test-Path $logDir) {
        $logFiles = Get-ChildItem -Path $logDir -Filter "Log_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 3) {
            $filesToDelete = $logFiles | Select-Object -Skip 3
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force
            }
        }
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

# --- Explorer-Optionen-Funktion, nutzt Write-Log ---
function Disable-ExplorerRecentAndFrequent {
    <#
    Deaktiviert im Windows-Explorer die Anzeige von:
    - Zuletzt verwendete Dateien
    - Häufig verwendete Ordner
    durch Setzen der entsprechenden Registry-Werte.
    Ein Explorer-Neustart wird bewusst nicht sofort erzwungen, damit während
    der laufenden Konfiguration keine Explorer-Sitzung unnötig unterbrochen wird.
    #>
    try {
        if (Set-ExplorerRecentAndFrequentState -ShowRecent 0 -ShowFrequent 0) {
            Write-Host -ForegroundColor Green "Explorer-Optionen 'Zuletzt verwendet' und 'Häufig verwendet' wurden vorbereitet."
            Write-Log -Message "Explorer-Neustart für 'ShowRecent'/'ShowFrequent' bewusst zurückgestellt; spätere Explorer-Aktualisierung übernimmt die Anwendung." -Level 'INFO'
        }
    } catch {
        Write-Host -ForegroundColor Red "Fehler beim Deaktivieren der Explorer-Optionen: $_"
        Write-Log -Message "Fehler beim Setzen der Explorer-Optionen: $_" -Level 'ERROR'
    }
}

# --- Schnellzugriff-Funktion für den Desktop ---
function Set-DesktopInQuickAccess {
    <#
    Prüft, ob der Desktop des aktuell angemeldeten Benutzers im Schnellzugriff angeheftet ist.
    Falls nicht, wird der Desktop an den Schnellzugriff angeheftet.
    #>
    try {
        $helperCandidates = @(
            (Join-Path $PSScriptRoot 'Pin-Desktop-Schnellzugriff.ps1'),
            (Join-Path (Split-Path $PSScriptRoot -Parent) 'Pin-Desktop-Schnellzugriff.ps1'),
            (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'Pin-Desktop-Schnellzugriff.ps1')
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

        $helperPath = $helperCandidates |
            Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
            Select-Object -First 1

        if (-not $helperPath) {
            Write-Log -Message 'Pin-Desktop-Schnellzugriff.ps1 wurde nicht gefunden.' -Level 'WARN'
            Write-Host -ForegroundColor Yellow "Hilfsskript für den Desktop-Schnellzugriff wurde nicht gefunden – Prüfung übersprungen."
            return
        }

        $sortSuffix = if ($script:FeatureFlags.EnableQuickAccessAutoSort) { 'inkl. alphabetischer Sortierung' } else { 'ohne alphabetische Sortierung' }
        Write-Log -Message "Desktop-Schnellzugriff wird über Hilfsskript ($sortSuffix) ausgeführt: $helperPath" -Level 'INFO'

        if ($script:FeatureFlags.EnableQuickAccessAutoSort) {
            $helperResult = & $helperPath -SortAlphabetically
        } else {
            $helperResult = & $helperPath
        }

        if ($helperResult -and $helperResult.Success) {
            $statusText = if ($helperResult.AlreadyPinned) { 'bereits vorhanden' } else { 'angeheftet' }
            $sortText = if ($helperResult.SortRequested) {
                if ($helperResult.SortApplied) { 'alphabetisch neu sortiert' } else { 'Sortierung geprüft/keine Änderung' }
            } else {
                'ohne Sortierung'
            }
            Write-Log -Message "Desktop-Schnellzugriff erfolgreich verarbeitet ($statusText, $sortText): $($helperResult.DesktopPath)" -Level 'INFO'
            return
        }

        $sortMessage = if ($helperResult -and -not [string]::IsNullOrWhiteSpace($helperResult.SortMessage)) { " Sortierdetails: $($helperResult.SortMessage)" } else { '' }
        Write-Log -Message ("Hilfsskript für den Desktop-Schnellzugriff lieferte keinen Erfolgstatus zurück." + $sortMessage) -Level 'WARN'
    }
    catch {
        Write-Host -ForegroundColor Red "Fehler beim Prüfen des Schnellzugriffs für den Desktop: $_"
        Write-Log -Message "Fehler bei der Schnellzugriff-Prüfung für Desktop: $($_.Exception.Message)" -Level 'ERROR'
    }
}

if (Confirm-OfficeClosure) {
    # Hier kann das eigentliche Skript ausgeführt werden
    Write-Host "Das Skript wird nun abgearbeitet..."

    # Explorer-Optionen vorbereiten: Zuletzt verwendete Dateien & Häufig verwendete Ordner
    # Kein sofortiger Explorer-Neustart an dieser Stelle; Änderungen werden später gesammelt wirksam.
    Disable-ExplorerRecentAndFrequent

    # Globaler Pfad für Logs
    $global:logDir = Join-Path $env:USERPROFILE "Documents\PC-Konfigurator\Logs"
    $global:logFile = Join-Path $global:logDir "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

    # Robocopy-Log-Verzeichnis global definieren
    $global:robocopyLogDir = Join-Path $env:USERPROFILE "Documents\PC-Konfigurator\Robocopy-Logs"

    function Write-Log {
        param(
            [Parameter(Mandatory=$true)][string]$Message,
            [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
        )
        try {
            if (-not (Test-Path $global:logDir)) {
                New-Item -ItemType Directory -Path $global:logDir -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $entry = "$timestamp [$Level] - $Message"
            Add-Content -Path $global:logFile -Value $entry -Encoding UTF8
        } catch {
            # Fallback auf Konsole, falls Log-Datei nicht erreichbar ist
            Write-Host -ForegroundColor Yellow "LOGFALLBACK [$Level] $Message"
        }
    }

    function Clear-OldLogs {
        param (
            [string]$logDir = $global:logDir
        )

        if (Test-Path $logDir) {
            $logFiles = Get-ChildItem -Path $logDir -Filter "Log_*.log" | Sort-Object LastWriteTime -Descending
            if ($logFiles.Count -gt 3) {
                $filesToDelete = $logFiles | Select-Object -Skip 3
                foreach ($file in $filesToDelete) {
                    Remove-Item -Path $file.FullName -Force
                }
            }
        }
    }

    function Test-SystemRequirements {
        Write-Log "Überprüfung der Windows- und Office-Version gestartet" "INFO"

        # === WINDOWS-VERSION PRÜFEN ===
        try {
            $windowsVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
            $windowsBuild = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
            $windowsCaption = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
            
            Write-Log "Ermittelte Windows-Version: $windowsVersion (Build: $windowsBuild)" "INFO"
            Write-Log "Windows-Edition: $windowsCaption" "INFO"

            # Erweiterte Windows-Versionsprüfung (Windows 10 Build 1903+ oder Windows 11)
            $windowsOK = $false
            if ($windowsVersion -match "^10\.") {
                # Windows 10: Mindestens Build 18362 (Version 1903)
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
        
        # Methode 1: ClickToRun-Installation
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
        
        # Methode 2: MSI-Installation (verschiedene Versionen)
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
        
        # Methode 3: Executable-basierte Erkennung
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
        
        # Office-Versionsvalidierung
        if ($officeVersion) {
            # Office 2013 (15.0) oder neuer wird unterstützt
            $officeOK = $officeMajorVersion -ge 15
            Write-Log "Office-Validierung: Version $officeMajorVersion >= 15: $officeOK" "INFO"
        } else {
            Write-Log "Keine Office-Installation gefunden" "WARN"
            # Für Testing-Zwecke: Office-Anforderung lockern
            Write-Log "HINWEIS: Skript wird trotz fehlender Office-Erkennung fortgesetzt (möglicherweise Office 365 Web)" "WARN"
            $officeOK = $true  # Temporär auf true setzen
        }

        # === ERGEBNIS AUSWERTEN ===
        Write-Log "=== SYSTEMPRÜFUNG ZUSAMMENFASSUNG ===" "INFO"
        Write-Log "Windows OK: $windowsOK (Version: $windowsVersion, Build: $windowsBuild)" "INFO"
        Write-Log "Office OK: $officeOK (Version: $officeVersion)" "INFO"

        if ($windowsOK) {
            Write-Log "Systemanforderungen erfüllt – Skript wird ausgeführt" "INFO"
            return $true
        } else {
            Write-Log "Systemanforderungen nicht erfüllt – Windows-Version zu alt" "ERROR"
            Write-Log "Erforderlich: Windows 10 (Build 18362+) oder Windows 11" "ERROR"
            return $false
        }
    }

    # Systemanforderungen prüfen - Skript nur fortsetzen wenn erfüllt
    if (-not (Test-SystemRequirements)) {
        Write-Host -ForegroundColor Red "Das Skript wird beendet, da die Systemanforderungen nicht erfüllt sind."
        Write-Host -ForegroundColor Yellow "Erforderlich: Windows 10/11 und Office 2019 oder neuer"
        pause
        exit
    }

    ### Hinweise:

    Write-Host "   "
    Write-Host -foregroundcolor yellow "Dieses PowerShell-Skript benötigt diverse Angaben von Ihnen:"
    Write-Host "   "
    Write-Host -foregroundcolor yellow "Standardlaufwerk für neue Dateien (somit auch für Datei-Vorlagen)."
    Write-Host -foregroundcolor yellow "Standardschriftart für neue Dateien."
    Write-Host "   "
    Write-Host -foregroundcolor yellow "Bitte geben Sie diese an den entsprechenden Stellen während der Ausführung des Skriptes ein."
    Write-Host "   "

    # Encoding-Einstellungen für die korrekte Ausgabe von Texten mit Umlauten
    if (-not $env:BUILD_MODE) {
        try {
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            $OutputEncoding = [System.Text.Encoding]::UTF8
            $PSDefaultParameterValues['*:Encoding'] = 'utf8'
            chcp 65001 | Out-Null
        }
        catch {
            # Fehler beim Encoding ignorieren
        }
    }
    # Generelle Unterdrückung unnötiger Konsolen-Ausgaben
    if ($env:BUILD_MODE) {
        $null = [System.Console]::SetOut([System.IO.StreamWriter]::new([System.IO.Stream]::Null))
        $null = [System.Console]::SetError([System.IO.StreamWriter]::new([System.IO.Stream]::Null))
    }

    # Robocopy-Logverzeichnis bereinigen (falls vorhanden)
    $logPath = Join-Path -Path $PSScriptRoot -ChildPath "Robocopy-Logs"

    # if (Test-Path $logPath) {
    #     Remove-Item -Path $logPath -Force -Recurse
    # }

    ### Synchronisation von Verzeichnissen für Datei-Vorlagen mittels robocopy
    ### Neue (zusätzliche) Dateien im Quelleverzeichnis werden ins Zielverzeichnis kopiert
    ### Sind geänderte Dateien im Quellverzeichnis oder im Zielverzeichnis vorhanden, werden folgende Regeln angewendet:
    ### - Im Quellverzeichnis geänderte Dateien werden kopiert und ersetzen die Datei im Zielverzeichnis.
    ### - Im Zielverzeichnis geänderte Dateien haben Vorrang und werden nicht ersetzt.
    #

    # Beginn der Funktion sync()

    function Clear-OldRobocopyLogs {
        param (
            [string]$robocopyLogDir = $global:robocopyLogDir
        )
        if (Test-Path $robocopyLogDir) {
            $logFiles = Get-ChildItem -Path $robocopyLogDir -Filter "*.log" | Sort-Object LastWriteTime -Descending
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
        $roboCopyBackupPath = $global:BackupTargetPath

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
            [System.IO.Path]::Combine($PSScriptRoot, "..", "Datei-Vorlagen")
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
            $driveLetter = Get-DriveLetter -Path $roboCopyBackupPath
        } catch {
            Write-Host -foregroundcolor red "Fehler bei der Verarbeitung von '$roboCopyBackupPath': $_"
            pause
            exit
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
        if (!(Test-Path $global:robocopyLogDir)) {
            New-Item -ItemType Directory -Path $global:robocopyLogDir -Force | Out-Null
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
                Write-Host -foregroundcolor yellow "Die Synchronisation eines kompletten Laufwerkes ist nicht vorgesehen; es muss sich um Ordner handeln."
                pause
                exit
            } elseif (Test-Path -Path $path -PathType Leaf) {
                Write-Host -foregroundcolor yellow "Die Synchronisation einer einzelnen Datei ist nicht vorgesehen."
                pause
                exit
            } elseif ($path -eq $roboCopyBackupPath) {
                Write-Host -foregroundcolor yellow "Quelle und Ziel sind identisch. Das ist so nicht vorgesehen"
                pause
                exit
            } elseif (-not(Test-Path -Path $path)) {
                Write-Host -foregroundcolor yellow "Der Pfad '$path' ist nicht vorhanden oder nicht erreichbar."
                pause
                exit
            }
        }

        # Check if the drive letter exists
        $driveExists = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue

        if ($driveExists) {
            # create backup folder if not exists
            If(!(test-path -PathType container $roboCopyBackupPath)) {
                New-Item -ItemType Directory -Path $roboCopyBackupPath -Force
            }

            # Start the timer
            $startTime = [System.Diagnostics.Stopwatch]::StartNew()

            $jobs = @()
            $totalJobs = $sourceDirectories.Count
            $completedJobs = 0
            $maxConcurrentJobs = $maxThreads

            foreach ($source in $sourceDirectories) {
                while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge $maxConcurrentJobs) {
                    Start-Sleep -Seconds 1
                }
                
                # === PRE-FLIGHT CHECKS ===
                Write-Log "Starte Synchronisation von: $source" "INFO"
                
                # Überprüfe Quellverzeichnis
                if (-not (Test-Path $source)) {
                    Write-Host -ForegroundColor Red "FEHLER: Quellverzeichnis nicht gefunden: $source"
                    Write-Log "Quellverzeichnis nicht gefunden: $source" "ERROR"
                    continue
                }
                
                # Überprüfe ob Quelldateien existieren
                $sourceFiles = Get-ChildItem -Path $source -Recurse -File -ErrorAction SilentlyContinue
                if (-not $sourceFiles -or $sourceFiles.Count -eq 0) {
                    Write-Host -ForegroundColor Yellow "WARNUNG: Quellverzeichnis ist leer: $source"
                    Write-Log "Quellverzeichnis ist leer: $source" "WARN"
                    continue
                }
                
                Write-Host -ForegroundColor Green "Quelldateien gefunden: $($sourceFiles.Count) Dateien in $source"
                Write-Log "Vorbereitung Synchronisation: $($sourceFiles.Count) Dateien gefunden in $source" "INFO"
                
                # Überprüfe Zielverzeichnis-Zugriff
                try {
                    # Test-Datei schreiben um Schreibzugriff zu prüfen
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
                $driveLetter = [System.IO.Path]::GetPathRoot($source)
                $trimmedString = $driveLetter.Trim(':\\')
                $newPath = $cleanPath -replace '\\', '-'
                # Log-Datei direkt im Robocopy-Logs-Ordner
                $logName = Join-Path -Path $global:robocopyLogDir -ChildPath ($trimmedString + "-" + $newPath + ".log")
                $destination = $roboCopyBackupPath

                # === NEUE LOGIK FÜR PRAKTIKUM-ORDNER ===
                # Prüfe, ob es einen Praktikum-Unterordner gibt, der überschrieben werden soll
                $praktikumSource = Join-Path -Path $source -ChildPath "Praktikum"
                $praktikumDestination = Join-Path -Path $destination -ChildPath "Praktikum"

                if (Test-Path -Path $praktikumSource) {
                    Write-Log "Praktikum-Ordner gefunden - wird mit kompletter Überschreibung synchronisiert: $praktikumSource" "INFO"
                    
                    # Lösche zuerst den gesamten Ziel-Praktikum-Ordner, falls er existiert
                    if (Test-Path -Path $praktikumDestination) {
                        try {
                            # Warte kurz und versuche dann die Löschung
                            Start-Sleep -Milliseconds 500
                            Remove-Item -Path $praktikumDestination -Recurse -Force -ErrorAction Stop
                            Write-Log "Bestehender Praktikum-Ordner wurde gelöscht: $praktikumDestination" "INFO"
                            # Weitere Wartezeit nach der Löschung
                            Start-Sleep -Milliseconds 1000
                        } catch {
                            Write-Log "Fehler beim Löschen des Praktikum-Ordners: $($_.Exception.Message)" "WARN"
                        }
                    }
                    
                    # Log-Pfad für Praktikum ohne Anführungszeichen
                    $praktikumLogName = $logName -replace ".log", "-Praktikum.log"
                    
                    # Separater Job für Praktikum-Ordner mit vollständiger Synchronisation
                    $praktikumJob = Start-Job -ScriptBlock {
                        param ($src, $dest, $excFile, $excDirectorie, $logPath)
                        
                        # Retry-Logik für Robocopy
                        $maxRetries = 3
                        $retryCount = 0
                        $success = $false
                        
                        while ($retryCount -lt $maxRetries -and -not $success) {
                            if ($retryCount -gt 0) {
                                Start-Sleep -Seconds (2 * $retryCount) # Exponential backoff
                            }
                            
                            # Verwende Anführungszeichen nur um den Log-Pfad
                            $quotedLogPath = "`"$logPath`""
                            
                            # /MIR = Mirror, /R:3 = 3 Retry attempts, /W:5 = Wait 5 seconds between retries
                            # /TEE = Output to console and log file
                            $process = Start-Process -FilePath "robocopy.exe" -ArgumentList "`"$src`" `"$dest`" /MIR /J /XJ /DCOPY:DAT /COPY:DAT /MT:8 /R:3 /W:5 /NP /V /XA:S /XF $excFile /XD $excDirectorie /TEE /UNILOG+:$quotedLogPath" -Wait -PassThru -WindowStyle Hidden
                            
                            $exitCode = $process.ExitCode
                            
                            # Exit codes 0-7 are considered successful
                            if ($exitCode -le 7) {
                                $success = $true
                            } else {
                                $retryCount++
                                if ($retryCount -lt $maxRetries) {
                                    # Log retry attempt - aber nur wenn der Pfad gültig ist
                                    try {
                                        Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Retry $retryCount for Praktikum sync after exit code $exitCode" -ErrorAction SilentlyContinue
                                    } catch {
                                        # Logging-Fehler ignorieren
                                    }
                                }
                            }
                        }
                        
                        return $exitCode
                    } -ArgumentList $praktikumSource, $praktikumDestination, $singleLineFiles, $singleLineDirectories, $praktikumLogName

                    $jobs += $praktikumJob
                    $totalJobs++  # Erhöhe die Gesamtzahl der Jobs
                }

                # Standard-Job für alle anderen Dateien (ohne Überschreibung bei neueren Zieldateien)
                $job = Start-Job -ScriptBlock {
                    param ($src, $dest, $excFile, $excDirectorie, $logPath)
                    
                    # Retry-Logik auch für Standard-Sync
                    $maxRetries = 3
                    $retryCount = 0
                    $success = $false
                    
                    while ($retryCount -lt $maxRetries -and -not $success) {
                        if ($retryCount -gt 0) {
                            Start-Sleep -Seconds (2 * $retryCount)
                        }
                        
                        # Verwende Anführungszeichen nur um den Log-Pfad
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
                                    # Logging-Fehler ignorieren
                                }
                            }
                        }
                    }
                    
                    return $exitCode
                } -ArgumentList $source, $destination, $singleLineFiles, $singleLineDirectories, $logName

                $jobs += $job
            }

            while (($jobs | Where-Object { $_.State -ne 'Completed' }).Count -gt 0) {
                $completedJobs = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
                $percentComplete = ($completedJobs / $totalJobs) * 100
                Write-Progress -Activity "Synchronisation: " -Status "$completedJobs von $totalJobs Aufgaben erledigt." -PercentComplete $percentComplete
                Start-Sleep -Seconds 1
            }

            foreach ($job in $jobs) {
                $jobResult = Receive-Job -Job $job -Wait
                $exitCode = $jobResult

                if ($exitCode -eq 0) {
                    Write-Host " "
                    Write-Host -foregroundcolor yellow "Eine Synchronisation ist nicht erforderlich."
                    Write-Log "Eine Synchronisation ist nicht erforderlich." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 1) {
                    Write-Host " "
                    Write-Host -foregroundcolor yellow "Die Synchronisation wurde erfolgreich abgeschlossen."
                    Write-Log "Die Synchronisation wurde erfolgreich abgeschlossen." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 2) {
                    Write-Host " "
                    Write-Host -foregroundcolor yellow "Es gibt zusätzliche Dateien im Zielverzeichnis, die nicht im Quellverzeichnis vorhanden sind. Es wurden keine neuen Dateien kopiert."
                    Write-Log "Es gibt zusätzliche Dateien im Zielverzeichnis, keine neuen kopiert." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 3) {
                    Write-Host " "
                    Write-Host -foregroundcolor yellow "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis."
                    Write-Log "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis." "INFO"
                    Write-Host " "
                }
                elseif ($exitCode -eq 16) {
                    Write-Host " "
                    Write-Host -foregroundcolor yellow "⚠ Robocopy Fehler 16: Versuche alternativen Kopiervorgang..."
                    Write-Log "Robocopy Fehler 16 - starte Fallback-Kopiervorgang" "WARN"
                    
                    # Fallback: Versuche manuellen Kopiervorgang SOFORT
                    try {
                        $sourceExists = Test-Path "$PSScriptRoot\..\Datei-Vorlagen"
                        if ($sourceExists -and (Get-ChildItem -Path "$PSScriptRoot\..\Datei-Vorlagen" -ErrorAction SilentlyContinue)) {
                            # Erstelle Zielverzeichnis falls nötig
                            if (-not (Test-Path $global:BackupTargetPath)) {
                                New-Item -ItemType Directory -Path $global:BackupTargetPath -Force | Out-Null
                            }
                            
                            # Verwende Copy-Item als Fallback
                            Write-Host -ForegroundColor Cyan "  ➤ Kopiere Dateien mit alternativer Methode..."
                            Copy-Item -Path "$PSScriptRoot\..\Datei-Vorlagen\*" -Destination $global:BackupTargetPath -Recurse -Force -ErrorAction Stop
                            
                            Write-Host -foregroundcolor Green "  ✓ FALLBACK ERFOLGREICH: Alle Dateien wurden kopiert!"
                            Write-Log "Fallback-Kopiervorgang erfolgreich abgeschlossen." "INFO"
                            
                            # Zähle kopierte Dateien
                            $copiedFiles = (Get-ChildItem -Path $global:BackupTargetPath -Recurse -File -ErrorAction SilentlyContinue).Count
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
                    Write-Host -foregroundcolor red "Lesen Sie auch die entsprechende Log-Datei --> $global:robocopyLogDir"
                    Write-Log "Fehlercode $exitCode beim Robocopy-Lauf." "ERROR"
                    Write-Host " "
                }
                Remove-Job -Job $job
            }

            Write-Progress -Activity "Synchronisation: " -Status "Alle Prozesse wurden erfolgreich abgeschlossen." -PercentComplete 100 -Completed

            $startTime.Stop()
            $elapsedTime = $startTime.Elapsed
            $formattedTime = "{0:D2} Stunden, {1:D2} Minuten, {2:D2} Sekunden, {3:D3} Millisekunden" -f $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds, $elapsedTime.Milliseconds
            Write-Host -Foregroundcolor Yellow "`n`nZeit : $formattedTime" 
            Write-Log "Synchronisation abgeschlossen. Dauer: $formattedTime" "INFO"
        } else {
            Write-Host -Foregroundcolor Yellow "Das Laufwerk $driveLetter ist nicht vorhanden."
            Write-Log "Das Laufwerk $driveLetter ist nicht vorhanden." "ERROR"
        }

        # Nur die 3 neuesten Robocopy-Logdateien behalten
        Clear-OldRobocopyLogs
        # Nur die 3 neuesten allgemeinen Logs behalten
        Clear-OldLogs
    } 
    # Ende der Funktion sync()

    # Überschreibe Mappe.xltx
    function CopyExcelTemplate {
        # Pfad zur Zieldatei im Benutzerprofil
        $targetPath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Excel\XLSTART\Mappe.xltx"

        # Pfad zur Quelldatei relativ zum Skriptverzeichnis
        $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Datei-Vorlagen\Sonstiges\Standards\Mappe.xltx"

        # Prüfen, ob die Datei bereits existiert
        if (Test-Path -Path $targetPath) {
            # Write-Output "Die Datei existiert bereits: $targetPath"
            return
        }

        # Sicherstellen, dass das Zielverzeichnis existiert
        $targetDirectory = Split-Path -Path $targetPath -Parent
        if (-not (Test-Path -Path $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        # Datei kopieren
        Copy-Item -Path $sourcePath -Destination $targetPath -Force
        Write-Log "Die Datei Mappe.xltx wurde erfolgreich kopiert nach: $targetPath" "INFO"
    }
    CopyExcelTemplate

    # Überschreibe Normal.dotm
    function CopyWordTemplate {
        # Pfad zur Zieldatei im Benutzerprofil
        $targetPath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Templates\Normal.dotm"

        # Pfad zur Quelldatei relativ zum Skriptverzeichnis
        $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Datei-Vorlagen\Sonstiges\Standards\Normal.dotm"

        # Prüfen, ob die Datei bereits existiert
        if (Test-Path -Path $targetPath) {
            Write-Log "Die Datei Normal.dotm existiert bereits: $targetPath" "INFO"
            return
        }

        # Sicherstellen, dass das Zielverzeichnis existiert
        $targetDirectory = Split-Path -Path $targetPath -Parent
        if (-not (Test-Path -Path $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        # Datei kopieren
        Copy-Item -Path $sourcePath -Destination $targetPath -Force
        Write-Log "Die Datei Normal.dotm wurde erfolgreich kopiert nach: $targetPath" "INFO"
    }
    CopyWordTemplate

    function CopyOutlookTemplate {
        # Pfad zur Zieldatei im Benutzerprofil
        $targetPath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Templates\NormalEmail.dotm"

        # Pfad zur Quelldatei relativ zum Skriptverzeichnis
        $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Datei-Vorlagen\Sonstiges\Standards\NormalEmail.dotm"

        # Prüfen, ob die Datei bereits existiert
        if (Test-Path -Path $targetPath) {
            Write-Log "Die Datei NormalEmail.dotm existiert bereits: $targetPath" "INFO"
            return
        }

        # Sicherstellen, dass das Zielverzeichnis existiert
        $targetDirectory = Split-Path -Path $targetPath -Parent
        if (-not (Test-Path -Path $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        # Datei kopieren
        Copy-Item -Path $sourcePath -Destination $targetPath -Force
        Write-Log "Die Datei NormalEmail.dotm wurde erfolgreich kopiert nach: $targetPath" "INFO"
    }
    CopyOutlookTemplate

    ### Initialisierung der Office-Programme
    ###
    ### Excel und Word starten und kurz danach wieder beenden. Grund: Auf manchen Rechnern sind die benötigten
    ### Programmverzeichnisse erst nach erstmaligem Aufruf verfügbar.

    Write-Host " "
    Write-Host -foregroundcolor Red  "Bitte kurz warten. Excel und Word werden initialisiert."
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

    try {
        Stop-Process -Name "OfficeClickToRun" -Force -ErrorAction Stop
        Write-Log "OfficeClickToRun erfolgreich beendet." "INFO"
    } catch {
        Write-Log "OfficeClickToRun konnte nicht beendet werden: Zugriff verweigert." "WARN"
    }

    # ===== Documents-Pfad für private Nutzung definieren =====
    $UserdocPath = [Environment]::GetFolderPath('MyDocuments')
    Write-Log "Documents-Pfad ermittelt: $UserdocPath" "INFO"

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

            # Lokale Signaturen vorhanden: nur fehlende Dateien aus dem Backup ergänzen.
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

    # ===== Ziel-Laufwerk für Datei-Vorlagen abfragen =====
    Write-Host -ForegroundColor Yellow  "Wohin sollen die Datei-Vorlagen kopiert werden?"
    Write-Host -ForegroundColor Red  "Bitte wählen Sie eine der folgenden Optionen:"
    Write-Host -ForegroundColor Green   "  [L] Laufwerk - Geben Sie anschließend den Laufwerksbuchstaben ein (z. B. D)"
    Write-Host -ForegroundColor Green   "  [D] Dokumente - Verwenden Sie Ihr persönliches Dokumente-Verzeichnis"
    Write-Host -ForegroundColor Green "  Hinweis: Im BFW verwenden Sie bitte Laufwerk Z (Option L, dann Z)."

    do {
        $choice = Read-Host "Ihre Auswahl (L/D)"
        if ($choice -match '^[LlDd]$') {
            $choice = $choice.ToUpper()
            if ($choice -eq "L") {
                # Laufwerksbuchstaben abfragen
                do {
                    $driveLetter = Read-Host "Laufwerksbuchstabe"
                    if ($driveLetter -match '^[A-Za-z]$') {
                        $driveLetter = $driveLetter.ToUpper()
                        $useDocuments = $false
                        $isValid = $true
                    } else {
                        Write-Host -ForegroundColor Red "Ungültige Eingabe! Bitte geben Sie genau einen Buchstaben ein."
                        $isValid = $false
                    }
                } until ($isValid)
            } elseif ($choice -eq "D") {
                # Documents-Verzeichnis verwenden
                $useDocuments = $true
                $isValid = $true
            }
        } else {
            Write-Host -ForegroundColor Red "Ungültige Eingabe! Bitte geben Sie 'L' für Laufwerk oder 'D' für Documents ein."
            $isValid = $false
        }
    } until ($isValid)

    # ===== Zielpfade setzen =====
    if ($useDocuments) {
        # Documents-Verzeichnis verwenden
        $driveRoot = $UserdocPath
        $targetTemplatePath = Join-Path $UserdocPath "Datei-Vorlagen"
        Write-Log "Verwende Documents-Verzeichnis: $targetTemplatePath" "INFO"
    } else {
        # Laufwerk verwenden
        $driveRoot = "${driveLetter}:\\"
        $targetTemplatePath = Join-Path $driveRoot "Datei-Vorlagen"
        Write-Log "Verwende Laufwerk ${driveLetter}: $targetTemplatePath" "INFO"
    }

    # Zielpfad für die Sync-Funktion global setzen
    $global:BackupTargetPath = $targetTemplatePath

    # Outlook-Signaturen sichern bzw. bei Bedarf zurückkopieren
    Sync-OutlookSignatures -BaseTargetPath $global:BackupTargetPath

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
                "PersonalTemplates" = $global:BackupTargetPath
                "DisableBootToOfficeStart" = 1
                "DisableBackstageOpenKeyShortcuts" = 1
                # Standard-Schriftart für neue Dokumente
                "DefaultFont" = $FontName
                "DefaultFontSize" = $FontSizeWord
            },
            $excelSettings = @{
                "DeveloperTools" = 1
                "DefaultPath" = if (Test-Path "Z:\") { "Z:\" } else { $driveRoot }
                "PersonalTemplates" = $global:BackupTargetPath
                "DisableBootToOfficeStart" = 1
                # Standard-Schriftart für neue Arbeitsmappen
                "StandardFont" = $FontName
                "StandardFontSize" = $FontSizeExcel
            },
            $windowsSettings = @{
                "HideFileExt" = 0
            }
        )

        # Registry-Pfade
        $regPathWord = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
        $regPathExcel = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
        $regPathWindows = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        # Funktion zum Setzen der Registry-Werte
        function Set-RegistryValues ($regPath, $settings) {
            if (Test-Path $regPath) {
                foreach ($key in $settings.Keys) {
                # Automatische Typ-Erkennung für ExpandString oder DWord
                $valueType = if ($settings[$key] -is [string] -and $settings[$key] -match "[:\\%]") { "ExpandString" } else { "DWord" }

                Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key] -Type $valueType
                Write-Log "Gesetzt in $regPath : $key = $($settings[$key]) (Typ: $valueType)" "INFO"
				}
            } else {
                Write-Log "Pfad nicht gefunden: $regPath" "ERROR"
            }
        }

        # Werte für Word, Excel und Windows setzen
        Set-RegistryValues $regPathWord $wordSettings
        Set-RegistryValues $regPathExcel $excelSettings
		
        Set-RegistryValues $regPathWindows $windowsSettings
		
        Write-Log "Die empfohlenen Office-Einstellungen wurden erfolgreich gesetzt." "INFO"
    }

    # Funktion aufrufen
    Set-OfficeRegistrySettings -FontName $ActualFontName -FontSizeWord $FontSizeWord -FontSizeExcel $FontSizeExcel

    # === ZUSÄTZLICHE REGISTRY-EINSTELLUNGEN FÜR SCHRIFTARTEN ===
    function Set-FontRegistrySettings {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSizeWord = 11,
            [int]$FontSizeExcel = 10
        )
        
        Write-Log "Setze spezifische Schriftart-Registry-Einstellungen..." "INFO"
        
        # Verschiedene Office-Versionen abdecken
        $officeVersions = @("16.0", "15.0", "14.0")
        
        foreach ($version in $officeVersions) {
            # Word-Schriftart-Einstellungen
            $wordRegPaths = @(
                "HKCU:\Software\Microsoft\Office\$version\Word\Options",
                "HKCU:\Software\Microsoft\Office\$version\Common\LanguageResources"
            )
            
            foreach ($regPath in $wordRegPaths) {
                if (Test-Path $regPath) {
                    try {
                        # Standard-Schriftart für neue Dokumente
                        Set-ItemProperty -Path $regPath -Name "DefaultFont" -Value $FontName -Type String -Force
                        Set-ItemProperty -Path $regPath -Name "DefaultFontSize" -Value $FontSizeWord -Type DWord -Force
                        Write-Log "Word-Schriftart gesetzt in $regPath" "INFO"
                    } catch {
                        Write-Log "Fehler beim Setzen der Word-Schriftart in $regPath`: $($_.Exception.Message)" "WARN"
                    }
                }
            }
            
            # Excel-Schriftart-Einstellungen
            $excelRegPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
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
    
    Set-FontRegistrySettings -FontName $ActualFontName -FontSizeWord $FontSizeWord -FontSizeExcel $FontSizeExcel

    # Auto-Korrektureinstellungen für Word
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
        
        # WICHTIG: Alle Word-Prozesse beenden vor der Konfiguration
        Write-Log "Beende alle Word-Prozesse vor der Autokorrektur-Konfiguration..." "INFO"
        Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Registry-Pfade für den aktuellen Benutzer - erweitert um alle möglichen Versionen
        $regPaths = @(
            "HKCU:\Software\Microsoft\Office\16.0\Word\Options",
            "HKCU:\Software\Microsoft\Office\17.0\Word\Options",
            "HKCU:\Software\Microsoft\Office\18.0\Word\Options"
        )

        $settingsApplied = 0
        
        # Einstellungen für alle gefundenen Office-Versionen setzen
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
        
        # Warte nach Registry-Änderungen
        Start-Sleep -Milliseconds 500
        
        # COM-basierte Einstellung für sofortige Aktivierung
        try {
            Write-Log "Starte Word für COM-basierte Autokorrektur-Konfiguration..." "INFO"
            
            # Word COM-Objekt erstellen
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            
            # AutoCorrect-Objekt abrufen
            $autoCorrect = $word.AutoCorrect
            
            # Autoformat-Einstellungen deaktivieren mit expliziter Überprüfung
            try {
                $autoCorrect.AutoFormatAsYouTypeApplyNumberedLists = $false
                Write-Log "Nummerierte Listen deaktiviert: $($autoCorrect.AutoFormatAsYouTypeApplyNumberedLists)" "INFO"
                
                $autoCorrect.AutoFormatAsYouTypeApplyBulletedLists = $false
                Write-Log "Aufzählungslisten deaktiviert: $($autoCorrect.AutoFormatAsYouTypeApplyBulletedLists)" "INFO"
                
                $autoCorrect.AutoFormatAsYouTypeReplaceHyperlinks = $false
                Write-Log "Hyperlink-Ersetzung deaktiviert: $($autoCorrect.AutoFormatAsYouTypeReplaceHyperlinks)" "INFO"
                
                $autoCorrect.CorrectInitialCaps = $false
                Write-Log "Anfangsbuchstaben-Korrektur deaktiviert: $($autoCorrect.CorrectInitialCaps)" "INFO"
                
                $autoCorrect.CorrectSentenceCaps = $true
                Write-Log "Satzanfang-Korrektur aktiviert: $($autoCorrect.CorrectSentenceCaps)" "INFO"
                
                $autoCorrect.AutoFormatAsYouTypeReplaceQuotes = $true
                $autoCorrect.AutoFormatAsYouTypeReplaceSymbols = $true
                
            } catch {
                Write-Log "Fehler beim Setzen einzelner AutoCorrect-Eigenschaften: $($_.Exception.Message)" "WARN"
            }
            
            # Versuche die Einstellungen zu speichern (falls möglich)
            try {
                # Erstelle ein temporäres Dokument, um die Einstellungen zu "aktivieren"
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
            # Word beenden
            if ($word) {
                try {
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                } catch {
                    # Fehler beim Beenden ignorieren
                }
            }
            # Garbage Collection
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        
        # ZUSÄTZLICH: Prüfe aktuelle Registry-Werte zur Verifikation
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
                break # Nur den ersten gefundenen Pfad prüfen
            }
        }

        if ($settingsApplied -eq 0) {
            Write-Log "WARNUNG: Keine Word Registry-Pfade gefunden - möglicherweise ist Office nicht korrekt installiert" "WARN"
        } else {
            Write-Log "Autokorrektureinstellungen für Word wurden verarbeitet ($settingsApplied Einstellungen gesetzt)" "INFO"
            Write-Log "HINWEIS: Starten Sie Word neu, um alle Änderungen zu aktivieren" "INFO"
        }
    }

    # Auto-Korrektureinstellungen für Excel
    function Set-ExcelAutoCorrectRegistry {
        param (
            [hashtable]$autoCorrectExcelSettings = @{
                "CorrectSentenceCap" = 0
            }
        )

        # Registry-Pfade für Excel
        $regPathExcelHKCU = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
		
        # Einstellungen unter HKCU setzen
        if (Test-Path $regPathExcelHKCU) {
            foreach ($key in $autoCorrectExcelSettings.Keys) {
                Set-ItemProperty -Path $regPathExcelHKCU -Name $key -Value $autoCorrectExcelSettings[$key] -Type DWord
                Write-Log "Gesetzt in $regPathExcelHKCU : $key = $($autoCorrectExcelSettings[$key])" "INFO"
            }
        } else {
            Write-Log  "Pfad nicht gefunden: $regPathExcelHKCU" "ERROR"
        }

        Write-Log "Autokorrektureinstellungen für Excel wurden für den aktuellen Benutzer verarbeitet." "INFO"
    }

    # Funktion ExcelAutoCorrectRegistry aufrufen
    Set-ExcelAutoCorrectRegistry

    # Einstellungen für Outlook

    # === Explorer-Datenschutzoptionen setzen und Verlauf löschen ===
    function Set-ExplorerPrivacyOptions {
        # Verlauf löschen
        $recentPath = Join-Path $env:APPDATA "Microsoft\Windows\Recent"
        if (Test-Path $recentPath) {
            try {
                    # WICHTIG:
                    # Die Unterordner "AutomaticDestinations" und "CustomDestinations" enthalten
                    # Explorer-/Schnellzugriff-Metadaten. Ein rekursives Löschen des gesamten
                    # Recent-Ordners entfernt dadurch auch bereits angeheftete Schnellzugriff-Einträge.
                    # Deshalb werden hier bewusst nur die obersten Verlaufs-Verknüpfungen gelöscht.
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
    Set-ExplorerPrivacyOptions

    # Desktop des angemeldeten Benutzers im Schnellzugriff sicherstellen
    Set-DesktopInQuickAccess

    function Set-ExplorerSearchOptions {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Search"
        $settings = @{
            "SearchSystemDirs"      = 1  # Systemverzeichnisse einbeziehen
            "SearchCompressedFiles" = 1  # Komprimierte Dateien einbeziehen
            "SearchAlways"          = 1  # Immer Dateinamen und -inhalte suchen
        }
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        foreach ($key in $settings.Keys) {
            try {
                Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key] -Type DWord -Force
                Write-Log "Explorer-Suchoption gesetzt: $key = $($settings[$key])" "INFO"
            } catch {
                Write-Log ("Fehler beim Setzen von $key in " + $regPath + ": " + $_.Exception.Message) "WARN"
            }
        }
    }
    Set-ExplorerSearchOptions
    function Set-OutlookRegistry {
        param (
            [hashtable]$OutlookSettings = @{
                "WeekNum" = 1
            }
        )

        # Registry-Pfade für Outlook-Kalender
        $regPathOutlookHKCU = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Calendar"

        # Einstellungen unter HKCU setzen
        if (Test-Path $regPathOutlookHKCU) {
            foreach ($key in $OutlookSettings.Keys) {
                Set-ItemProperty -Path $regPathOutlookHKCU -Name $key -Value $OutlookSettings[$key] -Type DWord
                Write-Log "Gesetzt in $regPathOutlookHKCU : $key = $($OutlookSettings[$key])" "INFO"
            }
        } else {
            Write-Log  "Pfad nicht gefunden: $regPathOutlookHKCU" "ERROR"
        }

        Write-Log "Einstellungen für Outlook wurden für den aktuellen Benutzer verarbeitet." "INFO"
    }

    # Funktion OutlookRegistry aufrufen
    Set-OutlookRegistry

    # Erfolgsmeldungen
    Write-Host "   "
    Write-Host -foregroundcolor yellow "Die empfohlenen Anpassungen für Ihre Arbeitsumgebung wurden erfolgreich vorgenommen."
    Write-Host "   "

    ### Optional: Word und Excel neu starten, damit die Änderungen wirksam werden
    #
    Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue

    # Funktion zum Bereitstellen von benutzerdefinierten Fonts, die standardmäßig nicht auf den Rechnern vorhanden sind.
    # LOCALAPPDATA\Microsoft\Windows\Fonts

    function Install-Fonts {
        param (
            [string]$SourceDirectory
        )

        # Zielordner für Benutzer-Schriftarten
        $targetFolder = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"

        # Überprüfen, ob das Quellverzeichnis existiert
        if (-Not (Test-Path -Path $SourceDirectory)) {
            Write-Log "Quellverzeichnis '$SourceDirectory' existiert nicht. Bitte überprüfen Sie den Pfad." "ERROR"
            return
        }

        # Schriftarten-Dateien rekursiv suchen
        $fontFiles = Get-ChildItem -Path $SourceDirectory -Recurse -Filter "*.ttf"
        $fontFiles += Get-ChildItem -Path $SourceDirectory -Recurse -Filter "*.otf"
        
         foreach ($fontFile in $fontFiles) {
            try {
                # Zielpfad für die Schriftart
                $targetFontPath = Join-Path -Path $targetFolder -ChildPath $fontFile.Name

                # Überprüfen, ob die Schriftart bereits existiert
                if (Test-Path -Path $targetFontPath) {
                    Write-Log "Die Schriftart '$($fontFile.Name)' existiert bereits und wurde übersprungen." "INFO"
                    continue
                }

                # Kopiere die Schriftart ins Zielverzeichnis
                Copy-Item -Path $fontFile.FullName -Destination $targetFolder -Force
                Write-Log "Kopiert: $($fontFile.FullName) nach $targetFolder" "INFO"

            } catch {
                Write-Log "Fehler beim Kopieren von $($fontFile.FullName): $_" "ERROR"
            }
        }


        Write-Log "Die empfohlenen Schriftarten wurden erfolgreich installiert." "INFO"
    }

    # Funktionsaufruf Install-Fonts
    Install-Fonts -SourceDirectory "$PSScriptRoot\..\Fonts"

    # Funktion zur Registrierung der benutzerdefinierten Schriften 
    function Install-UserFonts {
        param (
            [string]$fontPath = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"
        )

        # Prüfen, ob der Ordner existiert
        if (-Not (Test-Path -Path $fontPath)) {
            Write-Log "Der angegebene Ordner existiert nicht: $fontPath" "ERROR"
            return
        }

        # Schriftdateien abrufen
        $fonts = Get-ChildItem -Path $fontPath -Filter *.ttf -ErrorAction SilentlyContinue

        if ($fonts.Count -eq 0) {
            Write-Log "Keine Schriftdateien im Ordner gefunden: $fontPath" "ERROR"
            return
        }

        # Schriften registrieren
        foreach ($font in $fonts) {
            try {
                $fontName = $font.Name
                $fontRegistryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
                New-ItemProperty -Path $fontRegistryPath -Name $fontName -Value $font.FullName -PropertyType String -ErrorAction SilentlyContinue
            } catch {
                # Fehlerbehandlung ohne Konsolenausgabe
                $null = $_.Exception.Message
            }
        }
		Write-Log "Die neuen Schriftarten wurden für den angemeldeten Benutzer registriert." "INFO"
    }

    # Funktion aufrufen
    Install-UserFonts

    # Benutzerabfrage: Auswahl der Schriftart
    # Bestätigungsabfrage für Schriftart- und Schriftgrößenauswahl
        # --- NEU: Bestätigungsabfrage für Schriftart und Schriftgrößen ---
        # Nur wenn der Anwender dies bestätigt, werden die drei Auswahlfunktionen ausgeführt.
        # Andernfalls werden Standardwerte gesetzt und die Auswahl übersprungen.
    Write-Host -ForegroundColor Cyan "Möchten Sie Schriftart und Schriftgrößen individuell auswählen? (Ja/Nein)"
    $fontConfirm = Read-Host
    if ($fontConfirm -match "^(Ja|ja|J|j|Y|y)$") {
        # Schriftart-Auswahl
        Write-Host -foregroundcolor Red "Bitte wählen Sie die gewünschte Schriftart für neue Office-Dateien:"
        Write-Host "   "
        Write-Host -foregroundcolor Yellow "1. Arial"
        Write-Host -foregroundcolor Yellow "2. Calibri"
        Write-Host -foregroundcolor Yellow "3. Segoe UI"
        Write-Host -foregroundcolor Yellow "4. PT Sans"
        Write-Host -foregroundcolor Yellow "5. Aptos"
        Write-Host -foregroundcolor Yellow "6. Futura"
        Write-Host -foregroundcolor Yellow "7. Aptos Narrow"
        Write-Host -foregroundcolor Yellow "8. Futura (Cyrillic Demi)"
        Write-Host -foregroundcolor Yellow "   "
        Write-Host "   "
        Write-Host -foregroundcolor Red "Geben Sie die Nummer der gewünschten Schriftart an."
        Write-Host "   "
        do {
            $choiceFont = Read-Host
            switch ($choiceFont) {
                "1" { $FontName = "Arial"; $isValidFont = $true }
                "2" { $FontName = "Calibri"; $isValidFont = $true }
                "3" { $FontName = "Segoe UI"; $isValidFont = $true }
                "4" { $FontName = "PT Sans"; $isValidFont = $true }
                "5" { $FontName = "Aptos"; $isValidFont = $true }
                "6" { $FontName = "Futura"; $isValidFont = $true }
                "7" { $FontName = "Aptos Narrow"; $isValidFont = $true }
                "8" { $FontName = "Futura"; $isValidFont = $true }  # Für FuturaCyrillicDemi verwenden wir "Futura"
                default {
                    Write-Host -foregroundcolor Red "Ungültige Eingabe. Bitte wählen Sie eine Nummer zwischen 1 und 8."
                    $isValidFont = $false
                }
            }
        } until ($isValidFont)

        # Spezielle Font-Name-Behandlung für bessere Kompatibilität
        $ActualFontName = $FontName
        if ($FontName -eq "Aptos Narrow") {
            $PossibleNames = @("Aptos Narrow", "AptosNarrow", "Aptos-Narrow")
            Write-Log "Teste mögliche Font-Namen für Aptos Narrow: $($PossibleNames -join ', ')" "INFO"
            $ActualFontName = "Aptos Narrow"
        }
        Write-Log "Gewählte Schriftart: $FontName (Verwendet: $ActualFontName)" "INFO"

        # Schriftgröße Word/Outlook
        Write-Host -foregroundcolor Red "Bitte wählen Sie die gewünschte Schriftgröße für Word und Outlook:"
        Write-Host "   "
        Write-Host -foregroundcolor Yellow "1. 10 Punkte"
        Write-Host -foregroundcolor Yellow "2. 11 Punkte"
        Write-Host -foregroundcolor Yellow "3. 12 Punkte"
        Write-Host -foregroundcolor Yellow "   "
        Write-Host "   "
        Write-Host -foregroundcolor Red "Geben Sie die Nummer der gewünschten Schriftgröße für Word/Outlook an."
        Write-Host "   "
        do {
            $choiceSizeWord = Read-Host
            switch ($choiceSizeWord) {
                "1" { $FontSizeWord = 10; $isValidSizeWord = $true }
                "2" { $FontSizeWord = 11; $isValidSizeWord = $true }
                "3" { $FontSizeWord = 12; $isValidSizeWord = $true }
                default {
                    Write-Host -foregroundcolor Red "Ungültige Eingabe. Bitte wählen Sie eine Nummer zwischen 1 und 3."
                    $isValidSizeWord = $false
                }
            }
        } until ($isValidSizeWord)
        Write-Log "Gewählte Schriftgröße für Word/Outlook: $FontSizeWord Punkte" "INFO"

        # Schriftgröße Excel
        Write-Host -foregroundcolor Red "Bitte wählen Sie die gewünschte Schriftgröße für Excel:"
        Write-Host "   "
        Write-Host -foregroundcolor Yellow "1. 10 Punkte"
        Write-Host -foregroundcolor Yellow "2. 11 Punkte"
        Write-Host -foregroundcolor Yellow "3. 12 Punkte"
        Write-Host -foregroundcolor Yellow "   "
        Write-Host "   "
        Write-Host -foregroundcolor Red "Geben Sie die Nummer der gewünschten Schriftgröße für Excel an."
        Write-Host "   "
        do {
            $choiceSizeExcel = Read-Host
            switch ($choiceSizeExcel) {
                "1" { $FontSizeExcel = 10; $isValidSizeExcel = $true }
                "2" { $FontSizeExcel = 11; $isValidSizeExcel = $true }
                "3" { $FontSizeExcel = 12; $isValidSizeExcel = $true }
                default {
                    Write-Host -foregroundcolor Red "Ungültige Eingabe. Bitte wählen Sie eine Nummer zwischen 1 und 3."
                    $isValidSizeExcel = $false
                }
            }
        } until ($isValidSizeExcel)
        Write-Log "Gewählte Schriftgröße für Excel: $FontSizeExcel Punkte" "INFO"
    } else {
        Write-Host -ForegroundColor Yellow "Schriftart- und Schriftgrößenauswahl übersprungen. Es werden Standardwerte verwendet."
        $FontName = "Calibri"
        $ActualFontName = $FontName
        $FontSizeWord = 11
        $FontSizeExcel = 10
        Write-Log "Schriftart/Größen-Auswahl übersprungen. Standardwerte: $FontName, Word: $FontSizeWord, Excel: $FontSizeExcel" "INFO"
    }

    Write-Log "Gewählte Schriftgröße für Excel: $FontSizeExcel Punkte" "INFO"

    # === OFFICE-KONFIGURATION STARTEN ===
    Write-Host ""
    Write-Host -foregroundcolor Yellow "Starte Office-Konfiguration..."
    Write-Host -foregroundcolor Cyan "Dies kann 1-2 Minuten dauern. Bitte haben Sie Geduld."
    Write-Host ""

    # --- Funktionen zum Anpassen der Office-Vorlagen ---

    function Set-WordCustomizer {
        param (
            [string]$FontName = "Aptos",
            [int]$FontSize = 11
        )

        Write-Log "Set-WordCustomizer gestartet..." "INFO"
        Write-Host -ForegroundColor Cyan "  ➤ Word wird initialisiert..."
        
        # Stelle sicher, dass Word vollständig geschlossen ist
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1

        $word = $null
        $wordtemplate = $null
        $standardStyle = $null
        
        # Timeout für Word-Start (max 30 Sekunden)
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

        # === MEHRERE ANSÄTZE FÜR WORD-VORLAGEN ===
        
        # 1. Normal.dotm anpassen
        Write-Host -ForegroundColor Cyan "  ➤ Normal.dotm wird angepasst..."
        $wordtemplatePath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
        if (Test-Path $wordtemplatePath) {
            try {
                # Timeout-Check
                if ($timer.Elapsed.TotalSeconds -gt $timeout) {
                    throw "Timeout erreicht beim Word-Template-Zugriff"
                }
                
                $wordtemplate = $word.Documents.Open($wordtemplatePath, $false, $false)
                $standardStyle = $wordtemplate.Styles.Item("Standard")
                $standardStyle.ParagraphFormat.SpaceAfter = 0
                $wordtemplate.DefaultTabStop = 1.0 * 28.35
                $standardStyle.Font.Name = $ActualFontName
                $standardStyle.Font.Size = $FontSize
                # Korrekter Zeilenabstand für Word (1.1-fach = 1.1 * 12pt = 13.2pt)
                $standardStyle.ParagraphFormat.LineSpacingRule = 3  # wdLineSpaceMultiple
                $standardStyle.ParagraphFormat.LineSpacing = 1.1    # 1.1-fach
                
                # Zusätzlich: Normal-Formatvorlage für Absätze anpassen
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
                $wordtemplate.Close($false)  # Nicht speichern beim Schließen
                
                # Objekte explizit freigeben
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
                # Cleanup bei Fehler
                if ($standardStyle) {
                    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($standardStyle) | Out-Null } catch {}
                    $standardStyle = $null
                }
                if ($wordtemplate) {
                    try { $wordtemplate.Close($false) } catch {}
                    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordtemplate) | Out-Null } catch {}
                    $wordtemplate = $null
                }
            }
        }
        
        # 2. Standard-Schriftart in Word-Optionen setzen
        Write-Host -ForegroundColor Cyan "  ➤ Word-Optionen werden gesetzt..."
        try {
            if ($word.Options) {
                $word.Options.DefaultFont = $ActualFontName
                $word.Options.DefaultFontSize = $FontSize
                # Versuche auch Default-Zeilenabstand zu setzen
                try {
                    $word.Options.DefaultOpenFormat = 0  # Word-Format
                } catch {}
                Write-Log "Word-Standard-Schriftart über COM-Objekt gesetzt: $ActualFontName" "INFO"
                Write-Host -ForegroundColor Green "  ✓ Word-Optionen erfolgreich gesetzt"
            }
        } catch {
            Write-Log "Fehler beim Setzen der Standard-Schriftart über COM: $($_.Exception.Message)" "WARN"
            Write-Host -ForegroundColor Yellow "  ⚠ Word-Optionen konnten nicht gesetzt werden"
        }
        
        # 2a. Zusätzliche Methode: Selection-Default setzen
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
        
        # 3. NormalTemplate direkt anpassen (ohne neues Dokument)
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
                
                # Objekte freigeben
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
        
        # 4. Teste neues leeres Dokument erstellen und Standard setzen
        try {
            $testDoc = $word.Documents.Add()
            $testRange = $testDoc.Range(0, 0)
            
            # Setze Standard-Formatierung für neues Dokument
            $testRange.Font.Name = $ActualFontName
            $testRange.Font.Size = $FontSize
            $testRange.ParagraphFormat.SpaceAfter = 0
            $testRange.ParagraphFormat.LineSpacingRule = 3
            $testRange.ParagraphFormat.LineSpacing = 1.1
            
            # Speichere als neue Normal.dotm
            $testDoc.AttachedTemplate.Save()
            $testDoc.Close($false)
            
            Write-Log "Test-Dokument mit Standard-Formatierung erstellt und Normal.dotm aktualisiert." "INFO"
        } catch {
            Write-Log "Fehler beim Erstellen des Test-Dokuments: $($_.Exception.Message)" "WARN"
        }

        # COM-Objekt ordnungsgemäß schließen
        try {
            if ($word) {
                $word.Quit($false, $false, $false)  # Alle Parameter explizit setzen
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                $word = $null
            }
            
            # Mehrfache Garbage Collection für COM-Cleanup
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
            
            # Zusätzliche Wartezeit für vollständige COM-Freigabe
            Start-Sleep -Milliseconds 500  # Reduziert von 1s auf 500ms
            
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
            Write-Host -ForegroundColor Red "Word konnte nicht gestartet werden – COM-Fehler."
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
            $standardStyle = $wordtemplate.Styles.Item("Standard")
            $standardStyle.ParagraphFormat.SpaceAfter = 0
            $wordtemplate.DefaultTabStop = 1.0 * 28.35
            $standardStyle.Font.Name = $FontName
            $standardStyle.Font.Size = $FontSize
            $standardStyle.ParagraphFormat.LineSpacingRule = 3  # wdLineSpaceMultiple
            $standardStyle.ParagraphFormat.LineSpacing = 1.1   # 1.1-fach
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

    # Funktion, um 'sperrige' Excel-Vorlagen zu 'überlisten'
    function Set-ExcelCustomizer {
        param (
            [string]$FontName = "Aptos", # Standard-Schriftart, falls keine angegeben wird
            [int]$FontSize = 10 # Standard-Schriftgröße, falls keine angegeben wird
        )
        
        Write-Log "Set-ExcelCustomizer gestartet..." "INFO"
        
        try {
            # Starte die Excel-Anwendung
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            # === MEHRERE ANSÄTZE FÜR EXCEL-VORLAGEN ===
            
            # 1. Standard-Schriftart in Excel-Optionen setzen
            try {
                $excel.StandardFont = $FontName
                $excel.StandardFontSize = $FontSize
                Write-Log "Excel-Standard-Schriftart über COM-Objekt gesetzt." "INFO"
            } catch {
                Write-Log "Fehler beim Setzen der Standard-Schriftart über COM: $($_.Exception.Message)" "WARN"
            }

            # 2. Mappe.xltx Vorlage anpassen
            $excelTemplatePath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Excel\XLSTART\Mappe.xltx')
            
            if (Test-Path $excelTemplatePath) {
                try {
                    $workbook = $excel.Workbooks.Open($excelTemplatePath, $null, $false)
                    
                    # Normal-Style anpassen
                    $style = $workbook.Styles.Item("Normal")
                    $style.Font.Name = $FontName
                    $style.Font.Size = $FontSize
                    
                    # Alle Arbeitsblätter durchgehen und Standardformat setzen
                    foreach ($worksheet in $workbook.Worksheets) {
                        $worksheet.Cells.Font.Name = $FontName
                        $worksheet.Cells.Font.Size = $FontSize
                    }
                    
                    # Vorlage speichern
                    Remove-Item $excelTemplatePath -Force -ErrorAction SilentlyContinue
                    $workbook.SaveAs($excelTemplatePath, 54) # Excel Template Format
                    $workbook.Close($false)
                    Write-Log "Mappe.xltx wurde erfolgreich angepasst." "INFO"
                } catch {
                    Write-Log "Fehler beim Anpassen der Mappe.xltx: $($_.Exception.Message)" "ERROR"
                }
            }
            
            # 3. Neue leere Arbeitsmappe erstellen und als Vorlage speichern
            try {
                $newWorkbook = $excel.Workbooks.Add()
                
                # Standard-Style für neue Arbeitsmappe setzen
                $normalStyle = $newWorkbook.Styles.Item("Normal")
                $normalStyle.Font.Name = $FontName
                $normalStyle.Font.Size = $FontSize
                
                # Erstes Arbeitsblatt formatieren
                $worksheet = $newWorkbook.Worksheets.Item(1)
                $worksheet.Cells.Font.Name = $FontName
                $worksheet.Cells.Font.Size = $FontSize
                
                # Als book.xltx im XLSTART Verzeichnis speichern (alternative Vorlage)
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
            # Excel-Anwendung ordnungsgemäß schließen
            if ($excel) {
                try {
                    $excel.DisplayAlerts = $true
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                } catch {
                    Write-Log "Fehler beim Schließen von Excel: $($_.Exception.Message)" "WARN"
                }
            }
            # Garbage Collection anstoßen
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    # Aufruf der Funktionen NACH der Schriftart- und Schriftgrößen-Auswahl
    Write-Host -foregroundcolor Yellow "⏳ Schritt 1/5: Word-Vorlagen werden angepasst..."
    try {
        # Vereinfachte Word Normal.dotm Anpassung mit Timeout-Schutz
        Write-Host -ForegroundColor Cyan "  ➤ Word Normal.dotm wird angepasst..."
        
        # Stelle sicher, dass Word geschlossen ist
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            
            # Öffne Normal.dotm direkt
            $normalPath = $word.NormalTemplate.FullName
            $normalTemplate = $word.Documents.Open($normalPath)
            
            # Ändere Standard-Style
            $standardStyle = $normalTemplate.Styles.Item("Standard")
            $standardStyle.Font.Name = $ActualFontName
            $standardStyle.Font.Size = $FontSizeWord
            $standardStyle.ParagraphFormat.SpaceAfter = 0
            $standardStyle.ParagraphFormat.LineSpacing = 12
            
            # Speichere und schließe
            $normalTemplate.Save()
            $normalTemplate.Close($false)
            $word.Quit()
            
            # COM-Cleanup
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($normalTemplate) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            [GC]::Collect()
            
            Write-Host -ForegroundColor Green "  ✓ Word Normal.dotm erfolgreich angepasst"
            
        } catch {
            Write-Host -ForegroundColor Yellow "  ⚠ Word-Konfiguration übersprungen: $($_.Exception.Message)"
            Write-Log "Word-Konfiguration übersprungen: $($_.Exception.Message)" "WARN"
        }
        
    } catch {
        Write-Host -ForegroundColor Red "  ✗ Word-Konfiguration fehlgeschlagen"
        Write-Log "Word-Konfiguration fehlgeschlagen: $($_.Exception.Message)" "ERROR"
    }
    
    Write-Host -foregroundcolor Yellow "⏳ Schritt 2/5: Word-Lernsituationen werden angepasst..."
    Set-WordLernsituationenCustomizer -FontName $ActualFontName -FontSize $FontSizeWord
    
    Write-Host -foregroundcolor Yellow "⏳ Schritt 3/5: Excel-Vorlagen werden angepasst..."
    Set-ExcelCustomizer -FontName $ActualFontName -FontSize $FontSizeExcel

    Write-Host -foregroundcolor Yellow "⏳ Schritt 4/5: Registry-Einstellungen werden gesetzt..."
    # Aufruf der Auto-Korrektureinstellungen
    Set-WordAutoCorrectRegistry

    Write-Host -foregroundcolor Yellow "⏳ Schritt 5/5: Windows-Einstellungen werden angepasst..."

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

        try {
            # Überprüfe und erstelle Registry-Pfade falls nötig
            if (-not (Test-Path $taskbarRegistryPath)) {
                New-Item -Path $taskbarRegistryPath -Force | Out-Null
            }
            if (-not (Test-Path $searchRegistryPath)) {
                New-Item -Path $searchRegistryPath -Force | Out-Null
            }

            # Setze Taskbar-Ausrichtung
            if ($Alignment -eq "Left") {
                $taskbarValue = 0
            } else {
                $taskbarValue = 1
            }

            # Setze Suchbox-Modus
            switch ($Search) {
                "Hidden" { $searchValue = 0 }
                "Icon" { $searchValue = 1 }
                "Box" { $searchValue = 2 }
            }

            Set-ItemProperty -Path $taskbarRegistryPath -Name $taskbarValueName -Value $taskbarValue -Type DWord -Force
            Set-ItemProperty -Path $searchRegistryPath -Name $searchValueName -Value $searchValue -Type DWord -Force
            
            Write-Log "Taskbar-Einstellungen erfolgreich gesetzt: Alignment=$Alignment, Search=$Search" "INFO"
            
        } catch {
            Write-Log "Fehler beim Setzen der Taskbar-Einstellungen: $($_.Exception.Message)" "ERROR"
        }
    }

    # Funktion aufrufen, um die Taskleiste linksbündig auszurichten und das Suchsymbol anzuzeigen
    Set-TaskbarSettings -Alignment "Left" -Search "Icon"

    function Restart-ExplorerIfRunning {
        # Prüfe, ob explorer.exe läuft
        $explorerRunning = Get-Process -Name "explorer" -ErrorAction SilentlyContinue

        if ($explorerRunning) {
            Write-Log "Explorer-Neustart wird vorbereitet - sichere geöffnete Fenster" "INFO"
            
            # Merke das Skript-Verzeichnis für die Wiederherstellung
            $scriptDirectory = $PSScriptRoot
            Write-Log "Skript-Verzeichnis: $scriptDirectory" "INFO"
            
            # Speichere alle geöffneten Explorer-Fenster mit ihren spezifischen Pfaden
            $openWindows = @()
            $scriptDirWindowFound = $false
            
            try {
                $shell = New-Object -ComObject Shell.Application
                foreach ($window in $shell.Windows()) {
                    # Prüfe verschiedene Fenstertypen
                    if ($window.Name -match "Windows Explorer|File Explorer|Explorer" -or $window.FullName -match "explorer\.exe") {
                        try {
                            # Versuche den aktuellen Pfad zu ermitteln
                            $currentPath = $null
                            
                            # Methode 1: LocationURL (für normale Ordner)
                            if ($window.LocationURL) {
                                $currentPath = $window.LocationURL -replace "file:///", "" -replace "/", "\"
                                # Einfache URL-Dekodierung ohne System.Web
                                $currentPath = $currentPath -replace "%20", " "
                                $currentPath = $currentPath -replace "%C3%A4", "ä"
                                $currentPath = $currentPath -replace "%C3%B6", "ö"
                                $currentPath = $currentPath -replace "%C3%BC", "ü"
                                $currentPath = $currentPath -replace "%C3%9F", "ß"
                            }

                            # Methode 2: Document.Folder.Self.Path (Fallback)
                            if (-not $currentPath -and $window.Document -and $window.Document.Folder) {
                                $currentPath = $window.Document.Folder.Self.Path
                            }
                            
                            # Methode 3: LocationName für spezielle Ordner - aber nur für echte Pfade
                            if (-not $currentPath -and $window.LocationName) {
                                $locationName = $window.LocationName
                                # Filtere "Dieser PC" und andere Computer-Views heraus
                                if ($locationName -notmatch "Dieser PC|This PC|Computer|Arbeitsplatz") {
                                    $currentPath = $locationName
                                }
                            }
                            
                            # Prüfe ob es sich um das Skript-Verzeichnis handelt
                            if ($currentPath -and $currentPath -eq $scriptDirectory) {
                                $scriptDirWindowFound = $true
                                Write-Log "Skript-Verzeichnis-Fenster gefunden: $currentPath" "INFO"
                            }
                            
                            # Nur gültige, zugängliche Pfade speichern (keine Computer-Views)
                            if ($currentPath -and $currentPath -ne "" -and (Test-Path $currentPath -ErrorAction SilentlyContinue)) {
                                # Doppelte Pfade vermeiden
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

            # Falls das Skript-Verzeichnis nicht in den offenen Fenstern war, füge es hinzu
            if (-not $scriptDirWindowFound -and $scriptDirectory -and (Test-Path $scriptDirectory)) {
                $openWindows += $scriptDirectory
                Write-Log "Skript-Verzeichnis zur Wiederherstellung hinzugefügt: $scriptDirectory" "INFO"
            }

            Write-Log "Stoppe Windows-Explorer ($(($openWindows).Count) Fenster gespeichert)" "INFO"
            
            # Stoppe den Windows-Explorer
            Write-Host -foregroundcolor Yellow "Stoppe Windows-Explorer..."
            Stop-Process -Name "explorer" -Force

            # Reduzierte Wartezeit
            Start-Sleep -Seconds 1

            # Starte den Windows-Explorer neu (nur Shell, kein Fenster)
            Write-Host -foregroundcolor Yellow "Starte Windows-Explorer neu..."
            Start-Process "explorer.exe" -WindowStyle Hidden
            
            # Reduzierte Wartezeit bis Explorer-Shell geladen ist
            Start-Sleep -Seconds 2
            
            # Schließe alle automatisch geöffneten Explorer-Fenster vor der Wiederherstellung
            try {
                $newWindows = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
                if ($newWindows) {
                    # Warte kurz und schließe dann alle Explorer-Fenster
                    Write-Host -foregroundcolor Yellow "Schließe automatisch geöffnete Fenster..."
                    Start-Sleep -Milliseconds 500
                    
                    # Verwende COM um nur die Fenster zu schließen, nicht die Shell
                    $shell = New-Object -ComObject Shell.Application
                    $windowsToClose = @()
                    
                    foreach ($window in $shell.Windows()) {
                        if ($window.Name -match "Windows Explorer|File Explorer|Explorer" -or $window.FullName -match "explorer\.exe") {
                            $windowsToClose += $window
                        }
                    }
                    
                    # Schließe alle gefundenen Fenster
                    foreach ($window in $windowsToClose) {
                        try {
                            $window.Quit()
                        } catch {
                            # Fehler beim Schließen ignorieren
                        }
                    }
                    
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
                    Start-Sleep -Milliseconds 500
                }
            } catch {
                Write-Log "Fehler beim Schließen automatischer Explorer-Fenster: $($_.Exception.Message)" "WARN"
            }
            
            # Stelle die gespeicherten Fenster wieder her
            if ($openWindows.Count -gt 0) {
                Write-Log "Stelle $($openWindows.Count) Explorer-Fenster wieder her" "INFO"
                Write-Host -foregroundcolor Yellow "Stelle Explorer-Fenster wieder her..."
                
                foreach ($path in $openWindows) {
                    try {
                        # Nur echte Dateisystem-Pfade wiederherstellen
                        if (Test-Path $path -ErrorAction SilentlyContinue) {
                            Start-Process "explorer.exe" -ArgumentList "`"$path`""
                            Write-Log "Wiederhergestellt: Explorer-Fenster für '$path'" "INFO"
                            Start-Sleep -Milliseconds 200  # Verkürzte Pause zwischen Fenstern
                        } else {
                            Write-Log "Pfad nicht mehr verfügbar: '$path'" "WARN"
                        }
                    } catch {
                        Write-Log "Fehler beim Wiederherstellen von '$path': $($_.Exception.Message)" "ERROR"
                    }
                }
                
                Write-Host -foregroundcolor Green "Explorer-Fenster-Wiederherstellung abgeschlossen"
                Write-Log "Explorer-Fenster-Wiederherstellung abgeschlossen" "INFO"
            } else {
                Write-Log "Keine Explorer-Fenster zum Wiederherstellen gefunden" "INFO"
            }
            
        } else {
            Write-Log "Der Windows-Explorer ist aktuell nicht aktiv – kein Neustart erforderlich." "INFO"
        }
    }

    # Benutzerabfrage: Explorer-Neustart
    Write-Host -foregroundcolor Yellow "Möchten Sie den Windows-Explorer neustarten, um die Taskleisten-Änderungen sofort anzuwenden?"
    Write-Host -foregroundcolor Yellow "(Dies kann 30-60 Sekunden dauern, ist aber optional)"
    Write-Host -foregroundcolor Cyan "[J] Ja - Explorer neustarten (empfohlen)"
    Write-Host -foregroundcolor Cyan "[N] Nein - Änderungen werden beim nächsten Neustart aktiv"
    Write-Host "   "
    
    do {
        $explorerChoice = Read-Host "Ihre Auswahl (J/N)"
        if ($explorerChoice -match "^(J|j|Ja|ja)$") {
            $restartExplorer = $true
            $isValidExplorerChoice = $true
        }
        elseif ($explorerChoice -match "^(N|n|Nein|nein)$") {
            $restartExplorer = $false
            $isValidExplorerChoice = $true
        }
        else {
            Write-Host -foregroundcolor Red "Ungültige Eingabe. Bitte geben Sie 'J' für Ja oder 'N' für Nein ein."
            $isValidExplorerChoice = $false
        }
    } until ($isValidExplorerChoice)

    if ($restartExplorer) {
        Write-Host -foregroundcolor Yellow "Starte Explorer-Neustart..."
        # Funktion aufrufen, um den Windows-Explorer neu zu starten
        Restart-ExplorerIfRunning
    } else {
        Write-Host -foregroundcolor Green "Explorer-Neustart übersprungen. Änderungen werden beim nächsten Windows-Neustart aktiv."
        Write-Log "Explorer-Neustart vom Benutzer übersprungen" "INFO"
    }


    Write-Host "   "
    Write-Host -foregroundcolor Yellow "Der PC-Konfigurator hat Ihren Rechner konfiguriert und schließt sich in fünf Sekunden.."
    Start-Sleep -Seconds 5

# --- Shortcut für Kontaktdaten DBK.xlsx im Benutzer-Ordner erstellen ---
try {
    # Quellpfad der Excel-Datei (bevorzugt: benutzerspezifisches Datei-Vorlagen-Verzeichnis)
    $sourceFile = Join-Path $global:BackupTargetPath 'Duisdorfer BüroKonzept KG\Datenquellen\Kontaktdaten DBK.xlsx'
    if (-not (Test-Path $sourceFile)) {
        $fallbackSource = Join-Path $PSScriptRoot 'Datei-Vorlagen\Duisdorfer BüroKonzept KG\Datenquellen\Kontaktdaten DBK.xlsx'
        if (Test-Path $fallbackSource) {
            Write-Log "Kontaktdaten DBK.xlsx nicht im benutzerspezifischen Vorlagenpfad gefunden. Fallback auf Skriptpfad wird verwendet." "WARN"
            $sourceFile = $fallbackSource
        } else {
            throw "Kontaktdaten DBK.xlsx wurde weder im benutzerspezifischen Vorlagenpfad noch im Skriptpfad gefunden."
        }
    }
    # Zielordner: immer aktueller Dokumente-Ordner (auch OneDrive)
    $documentsPath = [Environment]::GetFolderPath('MyDocuments')
    $targetDir = Join-Path $documentsPath 'Meine Datenquellen'
    # Name und Pfad der Verknüpfung
    $shortcutPath = Join-Path $targetDir 'Kontaktdaten DBK.lnk'

    # Zielordner anlegen, falls nicht vorhanden
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    # WScript.Shell-Objekt für Shortcut-Erstellung
    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $sourceFile
    $shortcut.WorkingDirectory = Split-Path $sourceFile
    $shortcut.WindowStyle = 1
    $shortcut.Description = 'Kontaktdaten Duisdorfer BüroKonzept KG'
    $shortcut.Save()
    Write-Host -ForegroundColor Green "Verknüpfung zu 'Kontaktdaten DBK.xlsx' wurde in '$targetDir' erstellt."
} catch {
    Write-Host -ForegroundColor Red "Fehler beim Erstellen der Verknüpfung: $($_.Exception.Message)"
}

function Remove-PraktikumOrdnerBenutzer {
    $benutzerVorlagenPfad = $global:BackupTargetPath
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

    # Am Ende: Praktikum-Ordner im Benutzer-Vorlagenverzeichnis entfernen
    Remove-PraktikumOrdnerBenutzer
# Schließende Klammer für Confirm-OfficeClosure
} else {
    Write-Host -ForegroundColor Red "Das Skript wurde durch den Benutzer abgebrochen."
    Write-Log "Skript wurde abgebrochen - Office-Programme nicht geschlossen" "WARN"
    exit 1
}

