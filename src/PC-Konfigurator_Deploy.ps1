<#
.SYNOPSIS
    Erstellt ein ZIP aus einem lokalen Ordner mit automatischer Versionierung (LOKAL - ohne SharePoint)

.DESCRIPTION
    Diese Version erstellt nur lokale ZIP-Dateien und pflegt ein lokales Changelog.
    Keine SharePoint-Abhängigkeiten erforderlich.
#>

param(
    [string]$SourceFolder = ".",
    [string]$BaseName = "PC-Konfigurator",
    [string]$OutputFolder = ".\Release",
    [string]$ArchivOrdner = "Archiv",
    [string]$ChangelogFile = "CHANGELOG.md",
    [string]$HinweisDatei = "INSTALLATIONSHINWEISE.html",
    [string]$Kommentar = "Routine-Update",
    [string]$VersionOverride,
    [switch]$SkipLint
)

if (-not $SkipLint) {
    $lintScriptPath = Join-Path $PSScriptRoot 'PC-Konfigurator_Lint.ps1'

    if (Test-Path -LiteralPath $lintScriptPath -PathType Leaf) {
        Write-Output "Starte Lint-Routine vor dem Release-Build..."
        & $lintScriptPath -ProjectRoot $PSScriptRoot -SkipPowerShellLint

        if ($LASTEXITCODE -ne 0) {
            throw 'Lint-Routine meldet Fehler. Release wird abgebrochen.'
        }
    }
    else {
        Write-Warning "Lint-Skript nicht gefunden: $lintScriptPath. Linting wird übersprungen."
    }
}

# Ausgabeordner erstellen
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Output "Ausgabeordner erstellt: $OutputFolder"
}

# Archivordner erstellen
$archivPath = Join-Path $OutputFolder $ArchivOrdner
if (-not (Test-Path $archivPath)) {
    New-Item -ItemType Directory -Path $archivPath -Force | Out-Null
    Write-Output "Archiv-Ordner erstellt: $archivPath"
}

# Das Archiv ist ausschließlich für alte Release-Pakete, Git-Platzhalter und
# begleitende Dokumentation vorgesehen. Fehlablagen werden vor jedem Release
# entfernt, damit Vorlagen, Fonts oder Fremdprojektdateien nicht anwachsen.
$archiveFilesToRemove = Get-ChildItem -Path $archivPath -Recurse -Force -File |
    Where-Object {
        $_.Name -notin @('.gitignore', '.gitkeep') -and
        $_.Name -notlike "$BaseName-v*.zip" -and
        $_.Extension -notin @('.md', '.html')
    }
foreach ($archiveFile in $archiveFilesToRemove) {
    Remove-Item -LiteralPath $archiveFile.FullName -Force
    Write-Output "Archiv-Fehlablage entfernt: $($archiveFile.FullName)"
}
Get-ChildItem -Path $archivPath -Recurse -Force -Directory |
    Sort-Object FullName -Descending |
    Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force) } |
    Remove-Item -Force

# Hoechste vorhandene Version ermitteln (Format: vMajor.Minor, kompatibel mit altem vMajor)
$releaseFiles = Get-ChildItem -Path $OutputFolder -Filter "$BaseName*.zip" -File
if ($VersionOverride) {
    $newVersion = $VersionOverride.Trim()
} else {
    $latestVersion = [pscustomobject]@{ Major = 0; Minor = 0 }

    foreach ($f in $releaseFiles) {
        if ($f.Name -match "^$([regex]::Escape($BaseName))-v(?<Major>\d+)(?:\.(?<Minor>\d+))?\.zip$") {
            $major = [int]$matches['Major']
            $minor = if ($matches['Minor']) { [int]$matches['Minor'] } else { 0 }

            if ($major -gt $latestVersion.Major -or ($major -eq $latestVersion.Major -and $minor -gt $latestVersion.Minor)) {
                $latestVersion = [pscustomobject]@{ Major = $major; Minor = $minor }
            }
        }
    }

    if ($latestVersion.Major -eq 0 -and $latestVersion.Minor -eq 0) {
        $newVersion = "2.0"
    } else {
        $newVersion = "$($latestVersion.Major).$($latestVersion.Minor + 1)"
    }
}

$ZipName = "$BaseName-v$newVersion.zip"
$zipPath = Join-Path $OutputFolder $ZipName
$releaseBuildDate = Get-Date -Format 'dd.MM.yyyy HH:mm'

# Hilfsdateien für bestehende Release-Unterordner synchron halten
$releaseSupportFiles = @(
    'Pin-Desktop-Schnellzugriff.ps1'
)

$releaseStagingFolders = @(
    (Join-Path $OutputFolder 'PC-Konfigurator-v11')
) | Where-Object { Test-Path $_ -PathType Container }

foreach ($releaseFolder in $releaseStagingFolders) {
    foreach ($supportFile in $releaseSupportFiles) {
        $sourcePath = Join-Path $SourceFolder $supportFile
        $targetPath = Join-Path $releaseFolder $supportFile

        if (-not (Test-Path $sourcePath -PathType Leaf)) {
            Write-Output "Hilfsdatei nicht gefunden, Release-Ordner wird nicht aktualisiert: $sourcePath"
            continue
        }

        Copy-Item -Path $sourcePath -Destination $targetPath -Force
        Write-Output "Release-Hilfsdatei synchronisiert: $targetPath"
    }
}

# Aeltere Versionen ins Archiv verschieben
foreach ($file in $releaseFiles) {
    $archivFile = Join-Path $archivPath $file.Name
    Move-Item -Path $file.FullName -Destination $archivFile -Force
    Write-Output "Archiviert: $($file.Name)"
}

# ZIP erstellen (ohne das Release-Verzeichnis selbst)
Write-Output "Erstelle ZIP-Datei..."
$excludePaths = @(
    "Release",
    ".git",
    ".vs",
    ".venv",
    "venv",
    ".markdownlint.json",
    ".editorconfig",
    "node_modules",
    "package.json",
    "package-lock.json",
    "PSScriptAnalyzerSettings.psd1",
    "PC-Konfigurator_Lint.ps1",
    "_Entwicklung",
    "docs\*.docx",
    "*.log",
    "*.tmp",
    "pc-konfigurator.7z",
    "pc-konfigurator_deploy.ps1",
    "PC-konfigurator_deploy_local.ps1",
	"Deploy.bat"
)

# Temporaeres Verzeichnis fuer bereinigten Inhalt
$tempDir = Join-Path $env:TEMP "PC-Konfigurator-Build"
if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Alle Dateien kopieren (ausser ausgeschlossene)
Get-ChildItem -Path $SourceFolder -Recurse | ForEach-Object {
    $relativePath = $_.FullName.Substring((Resolve-Path $SourceFolder).Path.Length + 1)
    $shouldExclude = $relativePath -like 'docs\*.docx' -or $relativePath -like 'docs/*.docx'

    foreach ($exclude in $excludePaths) {
        if ($shouldExclude) {
            break
        }

        if ($relativePath -like "*$exclude*" -or $_.Name -like $exclude) {
            $shouldExclude = $true
            break
        }
    }

    if (-not $shouldExclude -and $_.PSIsContainer -eq $false) {
        $destPath = Join-Path $tempDir $relativePath
        $destDir = Split-Path $destPath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Copy-Item -Path $_.FullName -Destination $destPath -Force
    }
}

# ZIP aus dem temporaeren Verzeichnis erstellen
Compress-Archive -Path "$tempDir\*" -DestinationPath $zipPath -Force

# Release-Version und Erstelldatum in den Starter einbetten.
$releaseStarterPath = Join-Path $tempDir 'PC-Konfigurator.bat'
if (Test-Path -LiteralPath $releaseStarterPath -PathType Leaf) {
    $starterContent = Get-Content -LiteralPath $releaseStarterPath -Raw
    $starterContent = $starterContent.Replace('__RELEASE_VERSION__', "v$newVersion")
    $starterContent = $starterContent.Replace('__RELEASE_DATE__', $releaseBuildDate)
    Set-Content -LiteralPath $releaseStarterPath -Value $starterContent -Encoding ASCII -NoNewline
    # ZIP nach der Einbettung neu erstellen.
    Compress-Archive -Path "$tempDir\*" -DestinationPath $zipPath -Force
}

# Temporaeres Verzeichnis bereinigen
Remove-Item -Path $tempDir -Recurse -Force

$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
Write-Output "ZIP erstellt: $ZipName ($zipSize MB)"

# Installationshinweise generieren
$hinweisPath = Join-Path $OutputFolder $HinweisDatei
$hinweis = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="utf-8">
    <title>PC-Konfigurator Installation</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 20px; }
        h2 { color: #0078d4; }
        ol { line-height: 1.6; }
        .warning { background: #fff3cd; border: 1px solid #ffecb5; padding: 10px; border-radius: 5px; }
        .success { background: #d1edd1; border: 1px solid #badbba; padding: 10px; border-radius: 5px; }
    </style>
</head>
<body>
    <h2>PC-Konfigurator Installation - Version $newVersion</h2>

    <div class="success">
        <strong>Aktuelle Datei:</strong> $ZipName<br>
        <strong>Groesse:</strong> $zipSize MB<br>
        <strong>Datum:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm')
    </div>

    <h3>Installationsschritte:</h3>
    <ol>
        <li><strong>Download:</strong> Laden Sie die Datei <code>$ZipName</code> herunter.</li>
        <li><strong>Entpacken:</strong> Rechtsklick auf die ZIP-Datei → "Alle extrahieren..."</li>
        <li><strong>Ordner oeffnen:</strong> Navigieren Sie zum entpackten Ordner.</li>
        <li><strong>Empfohlener Start:</strong> Rechtsklick auf <code>PC-Konfigurator.bat</code> → "Als Administrator ausfuehren"</li>
        <li><strong>Installation folgen:</strong> Waehlen Sie Zielpfad (L = Laufwerk, D = Documents) und folgen Sie den Anweisungen.</li>
    </ol>

    <div class="warning">
        <strong>Wichtige Hinweise:</strong>
        <ul>
            <li>Schliessen Sie alle Office-Programme vor der Installation</li>
            <li>Administrator-Rechte sind empfohlen (fuer vollstaendige Anwendung aller Aenderungen)</li>
            <li>Windows 10 (Build 18362+) oder Windows 11 erforderlich</li>
            <li>Office 2013 oder neuer muss installiert sein</li>
        </ul>
    </div>

    <h3>Was wird installiert?</h3>
    <ul>
        <li><strong>Office-Vorlagen:</strong> Word, Excel, Outlook Templates</li>
        <li><strong>Schriftarten:</strong> Aptos, Montserrat, Font Awesome, weitere</li>
        <li><strong>Registry-Einstellungen:</strong> Optimierte Office-Konfiguration</li>
        <li><strong>Synchronisation:</strong> Datei-Vorlagen und Fonts</li>
    </ul>

    <h3>Dokumentation</h3>
    <ul>
        <li><a href="docs/DOKUMENTATION_ANWENDER.md">Anwenderdokumentation</a></li>
        <li><a href="docs/DOKUMENTATION_TECHNIK.md">Technische Dokumentation</a></li>
        <li><a href="docs/Registry-Einstellungen.md">Registry-Übersicht</a></li>
        <li><a href="docs/PC-Konfigurator%20-%20Einstellungen.pdf">Detaillierte Einstellungsübersicht (PDF)</a></li>
    </ul>

    <p><strong>Support:</strong> Bei Problemen wenden Sie sich an Ihren IT-Administrator.</p>
</body>
</html>
"@

$hinweis | Out-File -FilePath $hinweisPath -Encoding UTF8
Write-Output "Installationshinweise erstellt: $HinweisDatei"

# CHANGELOG pflegen
$changelogPath = Join-Path $OutputFolder $ChangelogFile
$datum = Get-Date -Format "yyyy-MM-dd HH:mm"

# Pruefen, ob CHANGELOG bereits existiert
if (Test-Path $changelogPath) {
    $clContent = (Get-Content $changelogPath -Raw -Encoding UTF8).Trim()
} else {
    $clContent = "# Changelog fuer $BaseName`r`n`r`nAutomatisch generiertes Changelog fuer PC-Konfigurator Releases."
}

# Neuen Eintrag oben einfuegen
$newEntry = @"
## Version $newVersion - $datum

**Aenderungen:**

- $Kommentar
- Datei: $ZipName
- Groesse: $zipSize MB

**Systemanforderungen:**

- Windows 10 (Build 18362+) oder Windows 11
- Office 2013 oder neuer
- Administrator-Rechte empfohlen

---
"@

if ([string]::IsNullOrWhiteSpace($clContent)) {
    $clContent = $newEntry.TrimEnd()
}
else {
    $clContent = $newEntry.TrimEnd() + "`r`n`r`n" + $clContent
}

$clContent = ($clContent -replace "(\r?\n){4,}", "`r`n`r`n`r`n").TrimEnd() + "`r`n"

# Speichern
Set-Content -LiteralPath $changelogPath -Value $clContent -Encoding UTF8 -NoNewline
Write-Output "Changelog aktualisiert: $ChangelogFile"

# Zusammenfassung
Write-Output ""
Write-Output "=== RELEASE ZUSAMMENFASSUNG ==="
Write-Output "Version: $newVersion"
Write-Output "Datei: $zipPath"
Write-Output "Groesse: $zipSize MB"
Write-Output "Hinweise: $hinweisPath"
Write-Output "Changelog: $changelogPath"
Write-Output "Archivierte Dateien: $($releaseFiles.Count)"

if ($releaseFiles.Count -gt 0) {
    Write-Output ""
    Write-Output "Archivierte Versionen:"
    $releaseFiles | ForEach-Object { Write-Output "  - $($_.Name)" }
}

Write-Output ""
Write-Output "Release erfolgreich erstellt!"

# Optional: SharePoint Upload nach erfolgreichem lokalem Build
if ($SharePointUpload -and $SharePointSite -and $SharePointLibrary) {
    Write-Output "Lade zu SharePoint hoch..."
    # SharePoint-Upload-Code hier einfügen
}