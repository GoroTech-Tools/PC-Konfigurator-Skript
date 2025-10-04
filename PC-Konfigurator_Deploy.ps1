<#
.SYNOPSIS
    Erstellt ein ZIP aus einem lokalen Ordner, vergi$hinweis = @"
# 7. Nutzerhinweis generieren
$hinweis = @"
<!DOCTYPE html>
<html lang="de">
<head><meta charset="utf-8"><title>Installationshinweis</title></head>
<body style="font-family:Segoe UI, sans-serif;">
    <h2>Installation des PC-Konfigurators</h2>
    <ol>
        <li>Laden Sie die Datei <b>$ZipName</b> herunter.</li>
        <li>Klicken Sie mit der rechten Maustaste auf die ZIP-Datei und waehlen Sie "Alle extrahieren...".</li>
        <li>Oeffnen Sie den entpackten Ordner.</li>
        <li>Starten Sie die Datei <b>PC-Konfigurator.bat</b> als Administrator durch Rechtsklick.</li>
        <li>Folgen Sie den Anweisungen im Installationsfenster.</li>
    </ol>
    <p style="color:red;">Hinweis: Bitte fuehren Sie die Installation nur mit Administrator-Rechten durch.</p>
</body>
</html>
"@html lang="de">
<head><meta charset="utf-8"><title>Installationshinweis</title></head>
<body style="font-family:Segoe UI, sans-serif;">
    <h2>Installation des PC-Konfigurators</h2>
    <ol>
        <li>Laden Sie die Datei <b>$ZipName</b> herunter.</li>
        <li>Klicken Sie mit der rechten Maustaste auf die ZIP-Datei und waehlen Sie "Alle extrahieren...".</li>
        <li>Oeffnen Sie den entpackten Ordner.</li>
        <li>Starten Sie die Datei <b>PC-Konfigurator.bat</b> als Administrator durch Rechtsklick.</li>
        <li>Folgen Sie den Anweisungen im Installationsfenster.</li>
    </ol>
    <p style="color:red;">Hinweis: Bitte fuehren Sie die Installation nur mit Administrator-Rechten durch.</p>
</body>
</html>
"@ne neue Versionsnummer,
    lädt es in eine SharePoint-Dokumentbibliothek hoch, archiviert ältere Versionen
    und pflegt ein Changelog (Markdown).

.Voraussetzungen
    - PowerShell 5+
    - Modul PnP.PowerShell (Install-Module PnP.PowerShell)
    - Zugriff auf die Ziel-Site und Bibliothek
#>

param(
        [string]$SourceFolder = ".",
        [string]$BaseName = "PC-Konfigurator",
        [string]$SiteUrl = "https://gorotech489.sharepoint.com/sites/zusammen/Freigegebene%20Dokumente/Forms/AllItems.aspx?id=%2Fsites%2Fzusammen%2FFreigegebene%20Dokumente%2FInformationstechnisches%20B%C3%BCromanagement&viewid=cc3007d7%2De6ef%2D4940%2D821d%2D9e0388742e15",
        [string]$Library = "Informationstechnisches Bueromanagement",
        [string]$HinweisDatei = "Hinweis.html",
        [string]$ArchivOrdner = "_Archiv",
        [string]$ChangelogFile = "CHANGELOG.md",
        [string]$Kommentar = "Routine-Update"
)

# 1. Verbindung zu SharePoint
Connect-PnPOnline -Url $SiteUrl -Interactive

# 2. Sicherstellen, dass Archiv-Ordner existiert
$archivFolder = Get-PnPFolder -Url "$Library/$ArchivOrdner" -ErrorAction SilentlyContinue
if (-not $archivFolder) {
        New-PnPFolder -Name $ArchivOrdner -Folder $Library | Out-Null
        Write-Host "Archiv-Ordner erstellt: $ArchivOrdner"
}

# 3. Hoechste vorhandene Version ermitteln
$files = Get-PnPFolderItem -FolderSiteRelativeUrl $Library -ItemType File |
        Where-Object { $_.Name -like "$BaseName*.zip" }

$maxVersion = 0
foreach ($f in $files) {
        if ($f.Name -match "$BaseName-v(\d+)\.zip") {
                $ver = [int]$matches[1]
                if ($ver -gt $maxVersion) { $maxVersion = $ver }
        }
}

$newVersion = $maxVersion + 1
$ZipName = "$BaseName-v$newVersion.zip"

# 4. Aeltere Versionen verschieben
foreach ($file in $files) {
        Move-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl `
                                 -TargetUrl "$Library/$ArchivOrdner/$($file.Name)" `
                                 -Force
        Write-Host "Archiviert: $($file.Name)"
}

# 5. ZIP erstellen
$zipPath = Join-Path $env:TEMP $ZipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $SourceFolder -DestinationPath $zipPath -Force
Write-Host "ZIP erstellt: $zipPath"

# 6. Neue ZIP hochladen
Add-PnPFile -Path $zipPath -Folder $Library -Overwrite
Write-Host "Neue ZIP ($ZipName) nach SharePoint hochgeladen."

# 7. Nutzerhinweis generieren
$hinweis = @"
<!DOCTYPE html>
<html lang="de">
<head><meta charset="utf-8"><title>Installationshinweis</title></head>
<body style="font-family:Segoe UI, sans-serif;">
    <h2>📥 Installation des Prüfungs-Setups</h2>
    <ol>
        <li>Laden Sie die Datei <b>$ZipName</b> herunter.</li>
        <li>Klicken Sie mit der rechten Maustaste auf die ZIP-Datei und wählen Sie „Alle extrahieren…“.</li>
        <li>Öffnen Sie den entpackten Ordner.</li>
        <li>Starten Sie die Datei <b>setup.exe</b> durch Doppelklick.</li>
        <li>Folgen Sie den Anweisungen im Installationsfenster.</li>
    </ol>
    <p style="color:red;">⚠️ Hinweis: Bitte führen Sie die Installation nur auf den vorgesehenen Prüfungsrechnern durch.</p>
</body>
</html>
"@

$hinweisPath = Join-Path $env:TEMP $HinweisDatei
$hinweis | Out-File -FilePath $hinweisPath -Encoding UTF8
Add-PnPFile -Path $hinweisPath -Folder $Library -Overwrite
Write-Host "Hinweisdatei nach SharePoint hochgeladen."

# 8. CHANGELOG pflegen
$changelogPath = Join-Path $env:TEMP $ChangelogFile
$datum = Get-Date -Format "yyyy-MM-dd HH:mm"

# Pruefen, ob CHANGELOG bereits existiert
try {
        $clItem = Get-PnPFile -Url "$Library/$ChangelogFile" -AsFile -Path $env:TEMP -ErrorAction Stop
        $clContent = Get-Content $changelogPath -Raw -Encoding UTF8
} catch {
        $clContent = "# Changelog fuer $BaseName`r`n`r`n"
}

# Neuen Eintrag oben einfuegen
$newEntry = "## Version $newVersion - $datum`r`n- $Kommentar`r`n- Datei: $ZipName`r`n`r`n"
$clContent = $newEntry + $clContent

# Speichern und hochladen
$clContent | Out-File -FilePath $changelogPath -Encoding UTF8
Add-PnPFile -Path $changelogPath -Folder $Library -Overwrite
Write-Host "Changelog aktualisiert."