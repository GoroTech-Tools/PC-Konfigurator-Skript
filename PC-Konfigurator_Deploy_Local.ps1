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
    [string]$Kommentar = "Routine-Update"
)

# Ausgabeordner erstellen
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Host "Ausgabeordner erstellt: $OutputFolder" -ForegroundColor Green
}

# Archivordner erstellen
$archivPath = Join-Path $OutputFolder $ArchivOrdner
if (-not (Test-Path $archivPath)) {
    New-Item -ItemType Directory -Path $archivPath -Force | Out-Null
    Write-Host "Archiv-Ordner erstellt: $archivPath" -ForegroundColor Green
}

# Hoechste vorhandene Version ermitteln
$files = Get-ChildItem -Path $OutputFolder -Filter "$BaseName*.zip" -File
$maxVersion = 0

foreach ($f in $files) {
    if ($f.Name -match "$BaseName-v(\d+)\.zip") {
        $ver = [int]$matches[1]
        if ($ver -gt $maxVersion) { $maxVersion = $ver }
    }
}

$newVersion = $maxVersion + 1
$ZipName = "$BaseName-v$newVersion.zip"
$zipPath = Join-Path $OutputFolder $ZipName

# Aeltere Versionen ins Archiv verschieben
foreach ($file in $files) {
    $archivFile = Join-Path $archivPath $file.Name
    Move-Item -Path $file.FullName -Destination $archivFile -Force
    Write-Host "Archiviert: $($file.Name)" -ForegroundColor Yellow
}

# ZIP erstellen (ohne das Release-Verzeichnis selbst)
Write-Host "Erstelle ZIP-Datei..." -ForegroundColor Blue
$excludePaths = @("Release", ".git", ".vs", "*.log", "*.tmp")

# Temporaeres Verzeichnis fuer bereinigten Inhalt
$tempDir = Join-Path $env:TEMP "PC-Konfigurator-Build"
if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Alle Dateien kopieren (ausser ausgeschlossene)
Get-ChildItem -Path $SourceFolder -Recurse | ForEach-Object {
    $relativePath = $_.FullName.Substring((Resolve-Path $SourceFolder).Path.Length + 1)
    $shouldExclude = $false
    
    foreach ($exclude in $excludePaths) {
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

# Temporaeres Verzeichnis bereinigen
Remove-Item -Path $tempDir -Recurse -Force

$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
Write-Host "ZIP erstellt: $ZipName ($zipSize MB)" -ForegroundColor Green

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
        <li><strong>Als Administrator starten:</strong> Rechtsklick auf <code>PC-Konfigurator.bat</code> → "Als Administrator ausfuehren"</li>
        <li><strong>Installation folgen:</strong> Waehlen Sie Zielpfad (L = Laufwerk, D = Documents) und folgen Sie den Anweisungen.</li>
    </ol>

    <div class="warning">
        <strong>Wichtige Hinweise:</strong>
        <ul>
            <li>Schliessen Sie alle Office-Programme vor der Installation</li>
            <li>Administrator-Rechte sind zwingend erforderlich</li>
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

    <p><strong>Support:</strong> Bei Problemen wenden Sie sich an Ihren IT-Administrator.</p>
</body>
</html>
"@

$hinweis | Out-File -FilePath $hinweisPath -Encoding UTF8
Write-Host "Installationshinweise erstellt: $HinweisDatei" -ForegroundColor Green

# CHANGELOG pflegen  
$changelogPath = Join-Path $OutputFolder $ChangelogFile
$datum = Get-Date -Format "yyyy-MM-dd HH:mm"

# Pruefen, ob CHANGELOG bereits existiert
if (Test-Path $changelogPath) {
    $clContent = Get-Content $changelogPath -Raw -Encoding UTF8
} else {
    $clContent = "# Changelog fuer $BaseName`r`n`r`nAutomatisch generiertes Changelog fuer PC-Konfigurator Releases.`r`n`r`n"
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
- Administrator-Rechte erforderlich

---

"@

$clContent = $newEntry + $clContent

# Speichern
$clContent | Out-File -FilePath $changelogPath -Encoding UTF8
Write-Host "Changelog aktualisiert: $ChangelogFile" -ForegroundColor Green

# Zusammenfassung
Write-Host "`n" -NoNewline
Write-Host "=== RELEASE ZUSAMMENFASSUNG ===" -ForegroundColor Cyan
Write-Host "Version: $newVersion" -ForegroundColor White
Write-Host "Datei: $zipPath" -ForegroundColor White  
Write-Host "Groesse: $zipSize MB" -ForegroundColor White
Write-Host "Hinweise: $hinweisPath" -ForegroundColor White
Write-Host "Changelog: $changelogPath" -ForegroundColor White
Write-Host "Archivierte Dateien: $($files.Count)" -ForegroundColor White

if ($files.Count -gt 0) {
    Write-Host "`nArchivierte Versionen:" -ForegroundColor Yellow
    $files | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
}

Write-Host "`nRelease erfolgreich erstellt!" -ForegroundColor Green