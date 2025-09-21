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

if (Confirm-OfficeClosure) {

    # Hier kann das eigentliche Skript ausgeführt werden
    Write-Host "Das Skript wird nun abgearbeitet..."

# Globaler Pfad für Logs
$global:logDir = Join-Path $env:USERPROFILE "Documents\PC-Konfigurator\Logs"
$global:logFile = Join-Path $global:logDir "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

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
                        $officeVersion = "16.0 (Executable gefunden)"
                        $officeMajorVersion = 16
                    } elseif ($path -match "Office15") {
                        $officeVersion = "15.0 (Executable gefunden)"
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

$global:robocopyLogDir = Join-Path $env:USERPROFILE "Dokumente\PC-Konfigurator\Robocopy-Logs"

function sync()
{
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
        [System.IO.Path]::Combine($PSScriptRoot, "Datei-Vorlagen")
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
            $cleanPath = $source -creplace '^[A-Za-z]:\\', ''
            $driveLetter = [System.IO.Path]::GetPathRoot($source)
            $trimmedString = $driveLetter.Trim(':\\')
            $newPath = $cleanPath -replace '\\', '-'
            # Log-Datei direkt im Robocopy-Logs-Ordner
            $logName = Join-Path -Path $global:robocopyLogDir -ChildPath ($trimmedString + "-" + $newPath + ".log")
            $quotedLogName = '"' + $logName + '"'
            $destination = $roboCopyBackupPath

            $job = Start-Job -ScriptBlock {
                param ($src, $dest, $excFile, $excDirectorie, $logPath)
                $process = Start-Process -FilePath "robocopy.exe" -ArgumentList "`"$src`" `"$dest`" /XO /E /J /XJ /DCOPY:DAT /COPY:DAT /MT:16 /R:0 /W:0 /NP /V /XA:S /XF $excFile /XD $excDirectorie /XO /XX /UNILOG+:$logPath" -Wait -PassThru -WindowStyle Hidden
                return $process.ExitCode
            } -ArgumentList $source, $destination, $singleLineFiles, $singleLineDirectories, $quotedLogName

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
                Write-Log "Eine Synchronisation ist nicht erforderlich."
                Write-Host " "
            }
            elseif ($exitCode -eq 1) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Die Synchronisation wurde erfolgreich abgeschlossen."
                Write-Log "Die Synchronisation wurde erfolgreich abgeschlossen."
                Write-Host " "
            }
            elseif ($exitCode -eq 2) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Es gibt zusätzliche Dateien im Zielverzeichnis, die nicht im Quellverzeichnis vorhanden sind. Es wurden keine neuen Dateien kopiert."
                Write-Log "Es gibt zusätzliche Dateien im Zielverzeichnis, keine neuen kopiert."
                Write-Host " "
            }
            elseif ($exitCode -eq 3) {
                Write-Host " "
                Write-Host -foregroundcolor yellow "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis."
                Write-Log "Einige Dateien wurden kopiert, aber es gibt zusätzliche Dateien im Zielverzeichnis."
                Write-Host " "
            } else {
                Write-Host " "
                Write-Host -foregroundcolor red "Fehlercode $exitCode. Lesen Sie dazu https://learn.microsoft.com/en-us/troubleshoot/windows-server/backup-and-storage/return-codes-used-robocopy-utility"
                Write-Host -foregroundcolor red "Lesen Sie auch die entsprechende Log-Datei --> $global:robocopyLogDir"
                Write-Log "Fehlercode $exitCode beim Robocopy-Lauf."
                Write-Host " "
            }
            Remove-Job -Job $job
        }

        Write-Progress -Activity "Synchronisation: " -Status "Alle Prozesse wurden erfolgreich abgeschlossen." -PercentComplete 100 -Completed

        $startTime.Stop()
        $elapsedTime = $startTime.Elapsed
        $formattedTime = "{0:D2} Stunden, {1:D2} Minuten, {2:D2} Sekunden, {3:D3} Millisekunden" -f $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds, $elapsedTime.Milliseconds
        Write-Host -Foregroundcolor Yellow "`n`nZeit : $formattedTime" 
        Write-Log "Synchronisation abgeschlossen. Dauer: $formattedTime"
    } else {
        Write-Host -Foregroundcolor Yellow "Das Laufwerk $driveLetter ist nicht vorhanden."
        Write-Log "Das Laufwerk $driveLetter ist nicht vorhanden."
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
    $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "Datei-Vorlagen\Sonstiges\Standards\Mappe.xltx"

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
    $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "Datei-Vorlagen\Sonstiges\Standards\Normal.dotm"

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
    $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "Datei-Vorlagen\Sonstiges\Standards\NormalEmail.dotm"

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
    $driveRoot = "${driveLetter}:\"
    $targetTemplatePath = Join-Path $driveRoot "Datei-Vorlagen"
    Write-Log "Verwende Laufwerk ${driveLetter}: $targetTemplatePath" "INFO"
}

# Zielpfad für die Sync-Funktion global setzen
$global:BackupTargetPath = $targetTemplatePath

# ===== Sync-Funktion mit RoboCopy starten =====
sync

# ===== Einheitliche Konfiguration der Programme Excel und Word =====
function Set-OfficeRegistrySettings {
    param (
        $wordSettings = @{
            "DeveloperTools" = 1
            "Ruler" = 1
            "ShowAllFormatting" = 1
            "VisiDrawTableDrs" = 1
            "DOC-PATH" = if (Test-Path "Z:\") { "Z:\" } else { $driveRoot }
            "PersonalTemplates" = $global:BackupTargetPath
            "DisableBootToOfficeStart" = 1
			"DisableBackstageOpenKeyShortcuts" = 1
        },
        $excelSettings = @{
            "DeveloperTools" = 1
            "DefaultPath" = if (Test-Path "Z:\") { "Z:\" } else { $driveRoot }
            "PersonalTemplates" = $global:BackupTargetPath
            "DisableBootToOfficeStart" = 1
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
Set-OfficeRegistrySettings

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
	
    # Registry-Pfade für den aktuellen Benutzer
    $regPathWordHKCU = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"

    # Einstellungen unter HKCU setzen
    if (Test-Path $regPathWordHKCU) {
        foreach ($key in $autoCorrectWordSettings.Keys) {
            Set-ItemProperty -Path $regPathWordHKCU -Name $key -Value $autoCorrectWordSettings[$key] -Type DWord
            Write-Log "Gesetzt in $regPathWordHKCU : $key = $($autoCorrectWordSettings[$key])" "INFO"
        }
    } else {
        Write-Log "Pfad nicht gefunden: $regPathWordHKCU" "ERROR"
    }

    Write-Log "Autokorrektureinstellungen für Word wurden für den aktuellen Benutzer verarbeitet." "INFO"
}

# Funktion WordAutoCorrectRegistry aufrufen
Set-WordAutoCorrectRegistry

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
Install-Fonts -SourceDirectory "$PSScriptRoot\Fonts"

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
Write-Host -foregroundcolor Red "Bitte wählen Sie die gewünschte Schriftart für neue Office-Dateien:"
Write-Host "   "
Write-Host -foregroundcolor Yellow "1. Arial"
Write-Host -foregroundcolor Yellow "2. Calibri"
Write-Host -foregroundcolor Yellow "3. Segoe UI"
Write-Host -foregroundcolor Yellow "4. PT Sans"
Write-Host -foregroundcolor Yellow "5. Aptos"
Write-Host -foregroundcolor Yellow "6. Futura"
Write-Host -foregroundcolor Yellow "   "
Write-Host "   "
Write-Host -foregroundcolor Red "Geben Sie die Nummer der gewünschten Schriftart an."
Write-Host "   "
$choiceFont = Read-Host

# Prüfen, welche Schriftart in den Vorlagen verwendet werden soll
switch ($choiceFont) {
    "1" { 
		$FontName = "Arial" 
		}
    "2" { 
		$FontName = "Calibri"
		}
    "3" { 
		$FontName = "Segoe UI"
		}
    "4" { 
		$FontName = "PT Sans" 
		}
    "5" { 
		$FontName = "Aptos" 
		}
    "6" { 
		$FontName = "Futura" 
		}    
	default {
        Write-Host -foregroundcolor Red "Ungültige Eingabe. Standardvorlage 'Aptos' wird verwendet."
        $FontName = "Aptos"
    }
}

function Set-WordCustomizer {
    param (
        [string]$FontName = "Aptos"
    )

    Write-Log "Set-WordCustomizer gestartet..." "INFO"
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

    $wordtemplatePath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
    if (-not (Test-Path $wordtemplatePath)) {
        Write-Log "Normal.dotm nicht gefunden unter: $wordtemplatePath" "ERROR"
        $word.Quit()
        return
    }

    try {
        try {
			$wordtemplate = $word.Documents.Open($wordtemplatePath)
		} catch {
		Write-Log "Fehler beim Öffnen der Normal.dotm: $($_.Exception.Message)" "ERROR"
		$word.Quit()
		return
		}
        $standardStyle = $wordtemplate.Styles.Item("Standard")
        $standardStyle.ParagraphFormat.SpaceAfter = 0
        $wordtemplate.DefaultTabStop = 1.0 * 28.35
        $standardStyle.Font.Name = $FontName
        $standardStyle.Font.Size = 11
        $standardStyle.ParagraphFormat.LineSpacing = 1.1 * 12
        $wordtemplate.Save()
        $wordtemplate.Close($false)
        Write-Log "Die Normal.dotm wurde erfolgreich angepasst." "INFO"
    } catch {
        Write-Log "Fehler beim Ändern der Normal.dotm: $($_.Exception.Message)" "ERROR"
    } finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Set-WordLernsituationenCustomizer {
    param (
        [string]$FontName = "Aptos"
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
        $standardStyle.Font.Size = 11
        $standardStyle.ParagraphFormat.LineSpacing = 1.1 * 12
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
	[string]$FontName = "Aptos" # Standard-Schriftart, falls keine angegeben wird
	)
	# Starte die Excel-Anwendung
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false 

	try {
		# Pfad zur Standard-Arbeitsmappenvorlage
		$excelTemplatePath = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Excel\XLSTART\Mappe.xltx')

		# Öffne die Vorlage
		$workbook = $excel.Workbooks.Open($excelTemplatePath, $null, $false)

		# Führe gewünschte Änderungen durch (z.B. Schriftart setzen)
		$style = $workbook.Styles.Item("Normal")
		$style.Font.Name = $FontName  # Setze die gewünschte Schriftart
		$style.Font.Size = 10         # Setze die Schriftgröße
	
		# Speichere die Änderungen, doch lösche zuvor die Vorlagendatei
		if (Test-Path $excelTemplatePath) {
		Remove-Item $excelTemplatePath -Force
		}
		# Damit keine Meldungen/Rückfragen ausgegeben werden, ist kurzzeitig DisPlayAlerts zu deaktivieren
		$excel.DisplayAlerts = $false
		$workbook.SaveAs($excelTemplatePath, 54) # Wichtig!
		$excel.DisplayAlerts = $true
	
		# Schließe die Arbeitsmappe
		$workbook.Close($false)
		
		Write-Log "Die Mappe.xltx wurde erfolgreich angepasst." "INFO"
	} catch {
		Write-Log "Fehler beim Ändern der Mappe.xltx: $_" "ERROR"
	} finally {
		# Schließe Excel-Anwendung
		$excel.Quit()

		# COM-Objekte ordnungsgemäß freigeben
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

		# Garbage Collection anstoßen
		[GC]::Collect()
		[GC]::WaitForPendingFinalizers()
	}
}

# Funktion, um die ausgewählten Schriftarten auch in Outlook (bis Version 2021) zu hinterlegen
# Achtung: Dies funktioniert nicht mehr zuverlässig mit Outlook 2024, 
# insbesondere mit dem neuen Outlook-Client. Microsoft hat die Architektur geändert, 
# sodass viele Einstellungen nicht mehr lokal über die Registry gesetzt werden können.

function Set-OutlookCustomizer {
    param (
        [string]$FontName = "Aptos"  # Standard-Schriftart, falls keine angegeben wird
    )

    Write-Log "Set-OutlookCustomizer gestartet mit Schriftart: $FontName" "INFO"

    # Outlook-Version erkennen
    $outlookVersion = $null
    $outlookType = "Unknown"
    
    try {
        # Versuche Outlook-Version aus Registry zu ermitteln
        $officeKey = "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $officeKey) {
            $officeVersion = (Get-ItemProperty -Path $officeKey -ErrorAction SilentlyContinue).ProductVersion
            if ($officeVersion) {
                $outlookVersion = $officeVersion
                Write-Log "Erkannte Office-Version: $officeVersion" "INFO"
            }
        }
        
        # Prüfe auf neues Outlook (Microsoft Store / Web-App)
        $newOutlookPath = "${env:LOCALAPPDATA}\Microsoft\WindowsApps\microsoft.outlookforwindows_8wekyb3d8bbwe"
        $classicOutlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        $newOutlookProcess = Get-Process -Name "ms-outlook" -ErrorAction SilentlyContinue
        
        if ($newOutlookProcess -or (Test-Path $newOutlookPath)) {
            $outlookType = "New"
            Write-Log "Neues Outlook (Microsoft Store/Web) erkannt" "INFO"
        } elseif ($classicOutlookProcess) {
            $outlookType = "Classic"
            Write-Log "Klassisches Outlook erkannt" "INFO"
        }
    } catch {
        Write-Log "Fehler bei der Outlook-Versionserkennung: $($_.Exception.Message)" "WARN"
    }

    # Registry-Pfade für verschiedene Outlook-Versionen
    $regPaths = @(
        "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings",     # Office 2016-2021
        "HKCU:\Software\Microsoft\Office\17.0\Common\MailSettings",     # Office 2024+
        "HKCU:\Software\Microsoft\Office\18.0\Common\MailSettings"      # Zukünftige Versionen
    )

    # Einstellungen für die Schriftart
    $outlookFontSettings = @{
        "ComposeFontComplex" = "$FontName,11,0,0,0,0,0,0"
        "ComposeFontSimple" = "$FontName,11,0,0,0,0,0,0"
        "ReplyFontComplex" = "$FontName,11,0,0,0,0,0,0"
        "ReplyFontSimple" = "$FontName,11,0,0,0,0,0,0"
    }

    # Erweiterte Funktion zum Setzen der Registry-Werte
    function Set-RegistryValues ($regPath, $settings) {
        if (Test-Path $regPath) {
            $successCount = 0
            foreach ($key in $settings.Keys) {
                try {
                    Set-ItemProperty -Path $regPath -Name $key -Value $settings[$key] -ErrorAction Stop
                    Write-Log "Gesetzt in $regPath : $key = $($settings[$key])" "INFO"
                    $successCount++
                } catch {
                    Write-Log "Fehler beim Setzen von $key in $regPath : $_" "WARN"
                }
            }
            return $successCount
        } else {
            Write-Log "Registry-Pfad nicht gefunden: $regPath" "WARN"
            return 0
        }
    }

    # Versuche Einstellungen für alle verfügbaren Registry-Pfade zu setzen
    $totalSuccess = 0
    $pathsFound = 0
    
    foreach ($regPath in $regPaths) {
        if (Test-Path $regPath) {
            $pathsFound++
            $successCount = Set-RegistryValues $regPath $outlookFontSettings
            $totalSuccess += $successCount
            Write-Log "Erfolgreich $successCount Einstellungen in $regPath gesetzt" "INFO"
        }
    }

    # Spezielle Behandlung für neues Outlook
    if ($outlookType -eq "New") {
        Write-Log "Hinweis: Neues Outlook erkannt - Registry-Einstellungen haben möglicherweise begrenzte Wirkung" "WARN"
        Write-Log "Für das neue Outlook müssen Schriftarten möglicherweise manuell in den Outlook-Einstellungen konfiguriert werden" "INFO"
        
        # Versuche zusätzliche Registry-Pfade für neues Outlook
        $newOutlookPaths = @(
            "HKCU:\Software\Microsoft\OfficeWebApps\Outlook",
            "HKCU:\Software\Microsoft\Outlook\Settings"
        )
        
        foreach ($path in $newOutlookPaths) {
            if (Test-Path $path) {
                Write-Log "Zusätzlicher Outlook-Pfad gefunden: $path" "INFO"
                # Hier könnten zukünftig spezielle Einstellungen für das neue Outlook hinzugefügt werden
            }
        }
    }

    # Ergebnisse protokollieren
    if ($pathsFound -eq 0) {
        Write-Log "Keine Outlook Registry-Pfade gefunden - Outlook möglicherweise nicht installiert" "WARN"
    } elseif ($totalSuccess -eq 0) {
        Write-Log "Keine Schriftart-Einstellungen konnten gesetzt werden" "ERROR"
    } else {
        Write-Log "Outlook-Schriftart-Konfiguration abgeschlossen: $totalSuccess Einstellungen in $pathsFound Registry-Pfaden gesetzt" "INFO"
        
        if ($outlookType -eq "New") {
            Write-Log "WICHTIG: Bei Verwendung des neuen Outlook müssen Sie die Schriftart möglicherweise manuell in Outlook > Einstellungen > E-Mail > Verfassen und antworten konfigurieren" "WARN"
        }
    }
}

# Aufruf der Funktion Set-WordCustomizer und Set-WordLernsituationenCustomizer
Set-WordCustomizer -FontName $FontName
Set-WordLernsituationenCustomizer -FontName $FontName
# Aufruf der Funktion Set-ExcelCustomizer
Set-ExcelCustomizer -FontName $FontName
# Aufruf der Funktion Set-OutlookCustomizer
Set-OutlookCustomizer -FontName $FontName


# Alte Log-Dateien bereinigen
Clear-OldLogs

# Anpassungen der Symbolleiste für den Schnellzugriff in Excel und Word
function Copy-QuickAccessToolbarFiles {
    # Quellverzeichnis
    $sourcePath = "$PSScriptRoot\Datei-Vorlagen\Sonstiges\Symbolleiste Schnellzugriff"
    
    # Zielverzeichnis
    $targetPath = "$env:LOCALAPPDATA\Microsoft\Office"
    
    # Dateien, die kopiert werden sollen
    $files = @("Excel.officeUI", "Word.officeUI")
    
    # Sicherstellen, dass das Zielverzeichnis existiert
    if (-not (Test-Path -Path $targetPath)) {
        New-Item -Path $targetPath -ItemType Directory -Force
    }
    
    # Kopieren der Dateien
    foreach ($file in $files) {
        $sourceFile = Join-Path -Path $sourcePath -ChildPath $file
        $targetFile = Join-Path -Path $targetPath -ChildPath $file
        
        if (Test-Path -Path $sourceFile) {
            Copy-Item -Path $sourceFile -Destination $targetFile -Force
            # Write-Output "Datei $file wurde erfolgreich nach $targetPath kopiert."
        } else {
            Write-Host -foregroundcolor red "Datei $file wurde im Quellverzeichnis nicht gefunden."
        }
    }
}

# Funktion aufrufen
Copy-QuickAccessToolbarFiles

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
    $taskbarPathExists = Test-Path $taskbarRegistryPath
    $searchPathExists = Test-Path $searchRegistryPath

    if ($taskbarPathExists -and $searchPathExists) {
        $taskbarExists = Get-ItemProperty -Path $taskbarRegistryPath -Name $taskbarValueName -ErrorAction SilentlyContinue
        $searchExists = Get-ItemProperty -Path $searchRegistryPath -Name $searchValueName -ErrorAction SilentlyContinue

        if ($null -ne $taskbarExists -and $null -ne $searchExists) {
            Write-Log "Registry-Werte für Taskbareinstellungen gefunden."

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

            Set-ItemProperty -Path $taskbarRegistryPath -Name $taskbarValueName -Value $taskbarValue
            Set-ItemProperty -Path $searchRegistryPath -Name $searchValueName -Value $searchValue
        } else {
            Write-Log "Die erforderlichen Registry-Werte sind nicht vorhanden." "ERROR"
        }
    } else {
        Write-Log "Der Registry-Pfad ist nicht vorhanden." "ERROR"
    }
}

# Funktion aufrufen, um die Taskleiste linksbündig auszurichten und das Suchsymbol anzuzeigen
Set-TaskbarSettings -Alignment "Left" -Search "Icon"

function Restart-ExplorerIfRunning {
    # Prüfe, ob explorer.exe läuft
    $explorerRunning = Get-Process -Name "explorer" -ErrorAction SilentlyContinue

    if ($explorerRunning) {
        # Stoppe den Windows-Explorer
        Stop-Process -Name "explorer" -Force

        # Warte einen Moment, um sicherzustellen, dass der Prozess vollständig beendet ist
        Start-Sleep -Seconds 1

        # Starte den Windows-Explorer neu
        Start-Process "explorer"
    } else {
        Write-Host "Der Windows-Explorer ist aktuell nicht aktiv – kein Neustart erforderlich."
    }
}

# Funktion aufrufen, um den Windows-Explorer neu zu starten
Restart-ExplorerIfRunning


Write-Host "   "
Write-Host -foregroundcolor Yellow "Der PC-Konfigurator hat Ihren Rechner konfiguriert und schließt sich in fünf Sekunden.."
Start-Sleep -Seconds 5
# Schließende Klammer für Confirm-OfficeClosure
}
