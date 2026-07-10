[CmdletBinding()]
param(
    [string]$ProjectRoot,
    [string]$AnalyzerSettingsPath,
    [switch]$SkipPowerShellLint,
    [switch]$SkipMarkdownLint,
    [bool]$FixTrailingWhitespace = $true
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } else {
        (Get-Location).Path
    }
}

if ([string]::IsNullOrWhiteSpace($AnalyzerSettingsPath)) {
    $AnalyzerSettingsPath = Join-Path $ProjectRoot 'PSScriptAnalyzerSettings.psd1'
}

function Get-MarkdownLintCommand {
    $localCandidates = @(
        (Join-Path $ProjectRoot 'node_modules\\.bin\\markdownlint-cli2.cmd'),
        (Join-Path $ProjectRoot 'node_modules\\.bin\\markdownlint-cli2')
    )

    foreach ($localCandidate in $localCandidates) {
        if (Test-Path -LiteralPath $localCandidate -PathType Leaf) {
            return $localCandidate
        }
    }

    $candidates = @('markdownlint-cli2')

    foreach ($candidate in $candidates) {
        $command = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($command) {
            return $command.Source
        }
    }

    return $null
}

function Initialize-PSScriptAnalyzer {
    try {
        Import-Module PSScriptAnalyzer -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning 'PSScriptAnalyzer nicht gefunden. Versuche automatische Installation (CurrentUser).'
    }

    try {
        Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module PSScriptAnalyzer -ErrorAction Stop
        Write-Output 'PSScriptAnalyzer wurde erfolgreich installiert.'
        return $true
    }
    catch {
        Write-Warning "PSScriptAnalyzer konnte nicht installiert werden: $($_.Exception.Message)"
        return $false
    }
}

function Get-ProjectPowerShellFiles {
    param(
        [string]$RootPath
    )

    Get-ChildItem -Path $RootPath -Recurse -File -Filter '*.ps1' |
        Where-Object {
            $_.FullName -notmatch '\\.venv\\' -and
            $_.FullName -notmatch '\\Release\\' -and
            $_.FullName -notmatch '\\_Entwicklung\\'
        } |
        Select-Object -ExpandProperty FullName
}

function Remove-TrailingWhitespace {
    param(
        [string[]]$FilePaths
    )

    $updatedFiles = @()
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)

    foreach ($filePath in $FilePaths) {
        $content = [System.IO.File]::ReadAllText($filePath)
        $cleaned = [System.Text.RegularExpressions.Regex]::Replace($content, '[\t ]+(?=\r?\n|$)', '')

        if ($cleaned -cne $content) {
            [System.IO.File]::WriteAllText($filePath, $cleaned, $utf8NoBom)
            $updatedFiles += $filePath
        }
    }

    return $updatedFiles
}

function Get-ProjectMarkdownFiles {
    param(
        [string]$RootPath
    )

    $searchRoot = $RootPath
    $parentPath = Split-Path -Path $RootPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parentPath)) {
        $parentReadme = Join-Path $parentPath 'README.md'
        if ((-not (Test-Path -LiteralPath (Join-Path $searchRoot 'README.md') -PathType Leaf)) -and (Test-Path -LiteralPath $parentReadme -PathType Leaf)) {
            $searchRoot = $parentPath
        }
    }

    $defaultMarkdownFiles = @(
        'README.md',
        'docs/DOKUMENTATION_ANWENDER.md',
        'docs/DOKUMENTATION_TECHNIK.md',
        'docs/Registry-Einstellungen.md',
        'Release/CHANGELOG.md'
    )

    $discoveredChangelogFiles = Get-ChildItem -Path $searchRoot -Recurse -File -Filter 'CHANGELOG.md' -ErrorAction SilentlyContinue |
        Where-Object {
            $_.FullName -notmatch '\\.git\\' -and
            $_.FullName -notmatch '\\node_modules\\' -and
            $_.FullName -notmatch '\\.venv\\'
        } |
        ForEach-Object {
            $relative = $_.FullName.Substring($searchRoot.Length).TrimStart('\\')
            $relative -replace '\\', '/'
        }

    return @($defaultMarkdownFiles + $discoveredChangelogFiles) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique |
        ForEach-Object {
            $relativePath = $_
            $rootCandidate = Join-Path $searchRoot ($relativePath -replace '/', '\\')
            if (Test-Path -LiteralPath $rootCandidate -PathType Leaf) {
                return $rootCandidate
            }

            $srcCandidate = Join-Path $RootPath ($relativePath -replace '/', '\\')
            if (Test-Path -LiteralPath $srcCandidate -PathType Leaf) {
                return $srcCandidate
            }

            return $null
        } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
}

function Test-MarkdownTableSeparatorLine {
    param(
        [string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $trimmed = $Line.Trim()
    return ($trimmed -match '^[\|:\-\s\t]+$' -and $trimmed -match '-')
}

function Test-MarkdownTableRowLine {
    param(
        [string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    $trimmed = $Line.Trim()
    if ($trimmed -notmatch '\|') {
        return $false
    }

    if ($trimmed -match '^\s*```' -or $trimmed -match '^\s*~~~') {
        return $false
    }

    return $true
}

function Get-MarkdownTableColumnCount {
    param(
        [string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return 0
    }

    $trimmed = $Line.Trim()
    if ($trimmed.StartsWith('|')) { $trimmed = $trimmed.Substring(1) }
    if ($trimmed.EndsWith('|')) { $trimmed = $trimmed.Substring(0, $trimmed.Length - 1) }

    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return 1
    }

    return ([System.Text.RegularExpressions.Regex]::Split($trimmed, '(?<!\\)\|')).Count
}

function Set-MarkdownTableColumnCount {
    param(
        [string]$Line,
        [int]$ExpectedColumnCount,
        [switch]$IsSeparator
    )

    if ($ExpectedColumnCount -le 0 -or [string]::IsNullOrWhiteSpace($Line)) {
        return $Line
    }

    $indent = ''
    $indentMatch = [System.Text.RegularExpressions.Regex]::Match($Line, '^\s*')
    if ($indentMatch.Success) {
        $indent = $indentMatch.Value
    }

    $trimmed = $Line.Trim()
    if ($trimmed.StartsWith('|')) { $trimmed = $trimmed.Substring(1) }
    if ($trimmed.EndsWith('|')) { $trimmed = $trimmed.Substring(0, $trimmed.Length - 1) }

    $cells = [System.Collections.Generic.List[string]]::new()
    foreach ($part in [System.Text.RegularExpressions.Regex]::Split($trimmed, '(?<!\\)\|')) {
        $cells.Add($part.Trim())
    }

    while ($cells.Count -lt $ExpectedColumnCount) {
        if ($IsSeparator) {
            $cells.Add('---')
        }
        else {
            $cells.Add('')
        }
    }

    if ($IsSeparator) {
        for ($i = 0; $i -lt $cells.Count; $i++) {
            if ([string]::IsNullOrWhiteSpace($cells[$i]) -or $cells[$i] -notmatch '^:?-{3,}:?$') {
                $cells[$i] = '---'
            }
        }
    }

    return ($indent + '| ' + ($cells -join ' | ') + ' |')
}

function Convert-MarkdownTables {
    param(
        [string]$Content
    )

    $lines = $Content -split "`r?`n", -1
    $inFence = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*```' -or $line -match '^\s*~~~') {
            $inFence = -not $inFence
            continue
        }

        if ($inFence) {
            continue
        }

        if (($i + 1) -ge $lines.Count) {
            continue
        }

        if (-not (Test-MarkdownTableRowLine -Line $lines[$i])) {
            continue
        }

        if (-not (Test-MarkdownTableSeparatorLine -Line $lines[$i + 1])) {
            continue
        }

        $expectedColumns = [Math]::Max(
            (Get-MarkdownTableColumnCount -Line $lines[$i]),
            (Get-MarkdownTableColumnCount -Line $lines[$i + 1])
        )

        $blockEnd = $i + 1
        for ($j = $i + 2; $j -lt $lines.Count; $j++) {
            if (-not (Test-MarkdownTableRowLine -Line $lines[$j])) {
                break
            }

            $blockEnd = $j
        }

        for ($j = $i; $j -le $blockEnd; $j++) {
            if ($j -eq ($i + 1)) {
                $lines[$j] = Set-MarkdownTableColumnCount -Line $lines[$j] -ExpectedColumnCount $expectedColumns -IsSeparator
            }
            else {
                $lines[$j] = Set-MarkdownTableColumnCount -Line $lines[$j] -ExpectedColumnCount $expectedColumns
            }
        }

        $i = $blockEnd
    }

    return ($lines -join "`r`n")
}

function Remove-MarkdownLintNoise {
    param(
        [string[]]$FilePaths
    )

    $updatedFiles = @()
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)

    foreach ($filePath in $FilePaths) {
        $content = [System.IO.File]::ReadAllText($filePath)

        $cleaned = [System.Text.RegularExpressions.Regex]::Replace($content, '[\t ]+(?=\r?\n|$)', '')
        $cleaned = Convert-MarkdownTables -Content $cleaned
        $cleaned = $cleaned.TrimEnd("`r", "`n") + "`r`n"

        if ($cleaned -cne $content) {
            [System.IO.File]::WriteAllText($filePath, $cleaned, $utf8NoBom)
            $updatedFiles += $filePath
        }
    }

    return $updatedFiles
}

Push-Location $ProjectRoot
try {
    $psFiles = Get-ProjectPowerShellFiles -RootPath $ProjectRoot

    if ($FixTrailingWhitespace -and $psFiles -and $psFiles.Count -gt 0) {
        Write-Output 'Bereinige Trailing Whitespace in PowerShell-Dateien...'
        $updatedFiles = Remove-TrailingWhitespace -FilePaths $psFiles
        if ($updatedFiles.Count -gt 0) {
            Write-Output "Trailing Whitespace entfernt in $($updatedFiles.Count) Datei(en)."
        }
        else {
            Write-Output 'Keine Trailing-Whitespace-Bereinigung erforderlich.'
        }
    }

    if (-not $SkipPowerShellLint) {
        if (-not (Initialize-PSScriptAnalyzer)) {
            Write-Warning 'PowerShell-Linting wird übersprungen, da PSScriptAnalyzer nicht verfügbar ist.'
            $SkipPowerShellLint = $true
        }
    }

    if (-not $SkipPowerShellLint) {
        if (-not (Test-Path -LiteralPath $AnalyzerSettingsPath -PathType Leaf)) {
            throw "PSScriptAnalyzer-Settings nicht gefunden: $AnalyzerSettingsPath"
        }

        Write-Output 'Starte PowerShell-Linting...'

            if (-not $psFiles -or $psFiles.Count -eq 0) {
                Write-Output 'PowerShell-Linting übersprungen: keine PS1-Dateien gefunden.'
            }

            $psResults = @()
            if ($psFiles -and $psFiles.Count -gt 0) {
                foreach ($psFile in $psFiles) {
                    $fileResults = Invoke-ScriptAnalyzer -Path $psFile -Settings $AnalyzerSettingsPath
                    if ($fileResults) {
                        $psResults += $fileResults
                    }
                }
            }

        if ($psResults) {
            $issueCount = @($psResults).Count
            Write-Output "PowerShell-Linting: $issueCount Befunde"
            $psResults | Format-Table -AutoSize RuleName, Severity, ScriptName, Line, Message
            throw 'PowerShell-Linting fehlgeschlagen.'
        }

        Write-Output 'PowerShell-Linting erfolgreich.'
    }

    if (-not $SkipMarkdownLint) {
        Write-Output 'Starte Markdown-Linting...'

        $markdownFiles = Get-ProjectMarkdownFiles -RootPath $ProjectRoot

        if (-not $markdownFiles -or $markdownFiles.Count -eq 0) {
            Write-Output 'Markdown-Linting übersprungen: keine Markdown-Dateien gefunden.'
            return
        }

        Write-Output 'Bereinige Markdown-Dateien (Whitespace, Tabellen, Abschluss-Umbruch)...'
        $updatedMarkdownFiles = Remove-MarkdownLintNoise -FilePaths $markdownFiles
        if ($updatedMarkdownFiles.Count -gt 0) {
            Write-Output "Markdown-Bereinigung angewendet in $($updatedMarkdownFiles.Count) Datei(en)."
        }
        else {
            Write-Output 'Keine Markdown-Bereinigung erforderlich.'
        }

        $markdownLintCommand = Get-MarkdownLintCommand

        if ($markdownLintCommand) {
            $argList = @('--config', '.markdownlint.json') + $markdownFiles
            & $markdownLintCommand @argList
            if ($LASTEXITCODE -ne 0) {
                throw 'Markdown-Linting fehlgeschlagen.'
            }

            Write-Output "Markdown-Linting erfolgreich für: $($markdownFiles -join ', ')"
        }
        else {
            Write-Warning 'Markdown-Linting übersprungen: markdownlint-cli2 nicht gefunden. Bitte `npm install` im Projekt ausführen.'
            Write-Output "Zu prüfende Dateien: $($markdownFiles -join ', ')"
        }
    }
}
finally {
    Pop-Location
}
