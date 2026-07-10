[CmdletBinding()]
param(
    [switch]$SortAlphabetically,
    [switch]$SortDryRun
)

$ErrorActionPreference = 'Stop'

function Get-NormalizedPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        return [System.IO.Path]::GetFullPath($Path).TrimEnd('\\').ToLowerInvariant()
    }
    catch {
        return $Path.TrimEnd('\\').ToLowerInvariant()
    }
}

function Get-CleanVerbName {
    param(
        [Parameter(Mandatory)]
        [object]$Verb
    )

    try {
        return ($Verb.Name -replace '&', '').Trim()
    }
    catch {
        return ''
    }
}

function Get-ShellVerb {
    param(
        [Parameter(Mandatory)]
        [object]$Item,
        [string[]]$CanonicalVerbs = @(),
        [string]$DisplayNameRegex = ''
    )

    $verbs = @($Item.Verbs())

    foreach ($canonicalVerb in $CanonicalVerbs) {
        $found = $verbs | Where-Object { $_.Verb -eq $canonicalVerb } | Select-Object -First 1
        if ($found) {
            return $found
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($DisplayNameRegex)) {
        $found = $verbs |
            Where-Object {
                $name = Get-CleanVerbName -Verb $_
                $name -match $DisplayNameRegex
            } |
            Select-Object -First 1

        if ($found) {
            return $found
        }
    }

    return $null
}

function Invoke-PinToQuickAccessByPath {
    param(
        [Parameter(Mandatory)]
        [object]$Shell,
        [Parameter(Mandatory)]
        [string]$Path
    )

    $parentPath = Split-Path -Path $Path -Parent
    $leafName = Split-Path -Path $Path -Leaf

    $parentFolder = $Shell.Namespace($parentPath)
    if (-not $parentFolder) {
        throw "Elternordner konnte nicht über Shell.Application geöffnet werden: $parentPath"
    }

    $item = $parentFolder.ParseName($leafName)
    if (-not $item) {
        throw "Element konnte in der Shell nicht aufgelöst werden: $Path"
    }

    $pinVerb = Get-ShellVerb -Item $item -CanonicalVerbs @('pintohome') -DisplayNameRegex 'Schnellzugriff|Quick access|Anheften|Pin|Start'
    if (-not $pinVerb) {
        throw "Kein passender Anheften-Befehl für Schnellzugriff gefunden: $Path"
    }

    $pinVerb.DoIt()
}

function Get-QuickAccessFilesystemEntries {
    param(
        [Parameter(Mandatory)]
        [object]$Shell
    )

    $quickAccessNamespace = $Shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
    if (-not $quickAccessNamespace) {
        return @()
    }

    $entries = foreach ($qaItem in @($quickAccessNamespace.Items())) {
        $path = $null
        $name = $null

        try { $path = $qaItem.Path } catch {}
        try { $name = $qaItem.Name } catch {}

        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        if (-not (Test-Path -LiteralPath $path -PathType Container)) {
            continue
        }

        [PSCustomObject]@{
            Name = if ([string]::IsNullOrWhiteSpace($name)) { Split-Path -Path $path -Leaf } else { $name }
            Path = $path
            NormalizedPath = Get-NormalizedPath -Path $path
            Item = $qaItem
        }
    }

    return @($entries)
}

function Test-StringSequenceEqual {
    param(
        [string[]]$A,
        [string[]]$B
    )

    if ($A.Count -ne $B.Count) {
        return $false
    }

    for ($index = 0; $index -lt $A.Count; $index++) {
        if ($A[$index] -ne $B[$index]) {
            return $false
        }
    }

    return $true
}

function Reorder-QuickAccessAlphabetically {
    param(
        [Parameter(Mandatory)]
        [object]$Shell,
        [switch]$DryRun
    )

    $entries = Get-QuickAccessFilesystemEntries -Shell $Shell
    if (-not $entries -or $entries.Count -lt 2) {
        return [PSCustomObject]@{
            Success = $true
            Applied = $false
            AlreadySorted = $true
            Message = 'Zu wenig Schnellzugriff-Ordner für eine Sortierung vorhanden.'
            Count = @($entries).Count
        }
    }

    $currentOrder = @($entries | Select-Object -ExpandProperty NormalizedPath)
    $targetEntries = @($entries | Sort-Object -Property @{ Expression = { $_.Name.ToLowerInvariant() } }, @{ Expression = { $_.NormalizedPath } })
    $targetOrder = @($targetEntries | Select-Object -ExpandProperty NormalizedPath)

    if (Test-StringSequenceEqual -A $currentOrder -B $targetOrder) {
        return [PSCustomObject]@{
            Success = $true
            Applied = $false
            AlreadySorted = $true
            Message = 'Schnellzugriff ist bereits alphabetisch sortiert.'
            Count = $entries.Count
        }
    }

    if ($DryRun) {
        return [PSCustomObject]@{
            Success = $true
            Applied = $false
            AlreadySorted = $false
            Message = 'Dry-Run: Sortierung erkannt, aber nicht angewendet.'
            Count = $entries.Count
            CurrentOrder = @($entries | Select-Object Name, Path)
            TargetOrder = @($targetEntries | Select-Object Name, Path)
        }
    }

    $unpinnedNormalizedPaths = New-Object System.Collections.Generic.List[string]
    $skippedUnpinPaths = @()

    foreach ($entry in $entries) {
        $unpinVerb = Get-ShellVerb -Item $entry.Item -CanonicalVerbs @('unpinfromhome') -DisplayNameRegex 'Schnellzugriff|Quick access|Unpin|Lösen|Entfernen|Remove'
        if (-not $unpinVerb) {
            $skippedUnpinPaths += $entry.Path
            Write-Warning "Kein passender Lösen-Befehl für Schnellzugriff gefunden, Eintrag wird übersprungen: $($entry.Path)"
            continue
        }

        try {
            $unpinVerb.DoIt()
            $null = $unpinnedNormalizedPaths.Add($entry.NormalizedPath)
            Start-Sleep -Milliseconds 150
        } catch {
            $skippedUnpinPaths += $entry.Path
            Write-Warning "Lösen aus Schnellzugriff fehlgeschlagen, Eintrag wird übersprungen: $($entry.Path) - $($_.Exception.Message)"
        }
    }

    if ($unpinnedNormalizedPaths.Count -eq 0) {
        return [PSCustomObject]@{
            Success = $true
            Applied = $false
            AlreadySorted = $false
            Message = 'Sortierung übersprungen: Für bestehende Schnellzugriff-Einträge wurde kein kompatibler Lösen-Befehl gefunden.'
            Count = $entries.Count
            SkippedPaths = $skippedUnpinPaths
        }
    }

    foreach ($entry in $targetEntries) {
        if ($unpinnedNormalizedPaths -notcontains $entry.NormalizedPath) {
            continue
        }
        Invoke-PinToQuickAccessByPath -Shell $Shell -Path $entry.Path
        Start-Sleep -Milliseconds 150
    }

    $reloadedEntries = Get-QuickAccessFilesystemEntries -Shell $Shell
    $reloadedOrder = @($reloadedEntries | Select-Object -ExpandProperty NormalizedPath)
    $sortVerified = Test-StringSequenceEqual -A $reloadedOrder -B $targetOrder

    $resultMessage = if ($sortVerified) { 'Alphabetische Sortierung wurde angewendet.' } else { 'Sortierung wurde angewendet, konnte aber nicht eindeutig verifiziert werden.' }
    if ($skippedUnpinPaths.Count -gt 0) {
        $resultMessage += " ($($skippedUnpinPaths.Count) Eintrag/Einträge übersprungen)"
    }

    return [PSCustomObject]@{
        Success = $true
        Applied = $true
        AlreadySorted = $false
        Message = $resultMessage
        Count = $entries.Count
        SortVerified = $sortVerified
        SkippedPaths = $skippedUnpinPaths
    }
}

function Remove-StaleDesktopQuickAccessEntries {
    param(
        [Parameter(Mandatory)]
        [object]$Shell,
        [Parameter(Mandatory)]
        [string]$ExpectedDesktopPath
    )

    $quickAccessNamespace = $Shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
    if (-not $quickAccessNamespace) {
        return 0
    }

    $expectedNormalized = Get-NormalizedPath -Path $ExpectedDesktopPath
    $removedCount = 0

    foreach ($qaItem in @($quickAccessNamespace.Items())) {
        $itemName = $null
        $itemPath = $null

        try { $itemName = $qaItem.Name } catch {}
        try { $itemPath = $qaItem.Path } catch {}

        $isDesktopEntry = ($itemName -eq 'Desktop')
        $isExpectedDesktop = ((Get-NormalizedPath -Path $itemPath) -eq $expectedNormalized)
        if (-not $isDesktopEntry -or $isExpectedDesktop) {
            continue
        }

        $unpinVerb = Get-ShellVerb -Item $qaItem -CanonicalVerbs @('unpinfromhome') -DisplayNameRegex 'Schnellzugriff|Quick access|Unpin|Lösen|Entfernen|Remove'

        if (-not $unpinVerb) {
            Write-Warning "Veralteter Desktop-Eintrag erkannt, aber keine Lösen-Aktion gefunden: Name='$itemName', Path='$itemPath'"
            continue
        }

        $unpinVerb.DoIt()
        $removedCount++
        Write-Host "Veralteter Desktop-Eintrag aus dem Schnellzugriff entfernt: $itemPath" -ForegroundColor Yellow
        Start-Sleep -Milliseconds 200
    }

    return $removedCount
}

try {
    # Benutzerspezifischen Desktop ermitteln (auch bei OneDrive-Umleitung korrekt)
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)

    if ([string]::IsNullOrWhiteSpace($desktopPath) -or -not (Test-Path -LiteralPath $desktopPath -PathType Container)) {
        throw "Desktop-Ordner wurde nicht gefunden: $desktopPath"
    }

    $shell = New-Object -ComObject Shell.Application

    $removedCount = Remove-StaleDesktopQuickAccessEntries -Shell $shell -ExpectedDesktopPath $desktopPath
    if ($removedCount -gt 0) {
        Write-Host "$removedCount veraltete Desktop-Einträge wurden aus dem Schnellzugriff entfernt." -ForegroundColor Yellow
    }

    # Bereits im Schnellzugriff enthalten?
    $quickAccessNamespace = $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
    $desktopNormalized = Get-NormalizedPath -Path $desktopPath

    $alreadyPinned = $false
    if ($quickAccessNamespace) {
        foreach ($qaItem in $quickAccessNamespace.Items()) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($qaItem.Path)) {
                    if ((Get-NormalizedPath -Path $qaItem.Path) -eq $desktopNormalized) {
                        $alreadyPinned = $true
                        break
                    }
                }
            }
            catch {
                # Nicht jedes Element liefert einen Dateisystempfad -> ignorieren
            }
        }
    }

    if ($alreadyPinned -and -not $SortAlphabetically) {
        Write-Host "Der benutzerspezifische Desktop ist bereits im Schnellzugriff angeheftet." -ForegroundColor Green
        return [PSCustomObject]@{
            Success = $true
            AlreadyPinned = $true
            RemovedCount = $removedCount
            SortRequested = [bool]$SortAlphabetically
            DesktopPath = $desktopPath
        }
    }

    if (-not $alreadyPinned) {
        Invoke-PinToQuickAccessByPath -Shell $shell -Path $desktopPath
        Start-Sleep -Milliseconds 300
        Write-Host "Desktop wurde erfolgreich an den Schnellzugriff angeheftet." -ForegroundColor Green
    }

    $sortResult = $null
    if ($SortAlphabetically) {
        Write-Host "Alphabetische Schnellzugriff-Sortierung wird ausgeführt..." -ForegroundColor Cyan
        $sortResult = Reorder-QuickAccessAlphabetically -Shell $shell -DryRun:$SortDryRun

        if ($sortResult.Applied -and $sortResult.Success) {
            Write-Host $sortResult.Message -ForegroundColor Green
        }
        elseif ($sortResult.Applied -and -not $sortResult.Success) {
            Write-Warning $sortResult.Message
        }
        else {
            Write-Host $sortResult.Message -ForegroundColor Yellow
        }
    }

    return [PSCustomObject]@{
        Success = if ($sortResult) { [bool]$sortResult.Success } else { $true }
        AlreadyPinned = $alreadyPinned
        RemovedCount = $removedCount
        SortRequested = [bool]$SortAlphabetically
        SortDryRun = [bool]$SortDryRun
        SortApplied = if ($sortResult) { [bool]$sortResult.Applied } else { $false }
        SortMessage = if ($sortResult) { $sortResult.Message } else { $null }
        DesktopPath = $desktopPath
    }
}
catch {
    Write-Error "Fehler beim Anheften des Desktops: $($_.Exception.Message)"
    throw
}
