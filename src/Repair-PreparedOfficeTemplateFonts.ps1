<#
.SYNOPSIS
Synchronisiert Theme- und Schriftreferenzen vorbereiteter Office-Vorlagen mit
jener Schriftart, die im jeweiligen Dateinamen angegeben ist.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TemplateRoot = (Join-Path $PSScriptRoot '..\Datei-Vorlagen\Sonstiges\Standards')
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Update-ZipXmlEntry {
    param(
        [Parameter(Mandatory = $true)][string]$ArchivePath,
        [Parameter(Mandatory = $true)][string]$EntryName,
        [Parameter(Mandatory = $true)][scriptblock]$Transform
    )

    $archive = [System.IO.Compression.ZipFile]::Open($ArchivePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $archive.GetEntry($EntryName)
        if (-not $entry) {
            throw "XML-Eintrag fehlt: $EntryName"
        }

        $reader = [System.IO.StreamReader]::new($entry.Open())
        try {
            $content = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }

        $updatedContent = & $Transform $content
        if ($updatedContent -eq $content) {
            return $false
        }

        $entry.Delete()
        $newEntry = $archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = [System.IO.StreamWriter]::new($newEntry.Open(), [System.Text.UTF8Encoding]::new($false))
        try {
            $writer.Write($updatedContent)
        } finally {
            $writer.Dispose()
        }
        return $true
    } finally {
        $archive.Dispose()
    }
}

function Set-ThemeLatinFonts {
    param(
        [Parameter(Mandatory = $true)][string]$Xml,
        [Parameter(Mandatory = $true)][string]$FontName
    )

    [regex]::Replace(
        $Xml,
        '(<a:(?:majorFont|minorFont)>.*?<a:latin\b[^>]*\btypeface=")[^"]*(")',
        {
            param($match)
            "$($match.Groups[1].Value)$FontName$($match.Groups[2].Value)"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
}

function Set-WordFontReferences {
    param(
        [Parameter(Mandatory = $true)][string]$Xml,
        [Parameter(Mandatory = $true)][string]$FontName
    )

    $xmlWithoutThemeReferences = [regex]::Replace(
        $Xml,
        '\s+w:(?:asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme)="[^"]*"',
        ''
    )

    [regex]::Replace($xmlWithoutThemeReferences, '<w:rFonts\b[^>]*/>', {
            param($fontMatch)
            $tag = $fontMatch.Value
            foreach ($attribute in @('ascii', 'hAnsi', 'eastAsia', 'cs')) {
                $attributePattern = '\bw:' + [regex]::Escape($attribute) + '="[^"]*"'
                $replacement = 'w:' + $attribute + '="' + $FontName + '"'
                if ($tag -match $attributePattern) {
                    $tag = [regex]::Replace($tag, $attributePattern, $replacement)
                } else {
                    $tag = $tag -replace '(/>)$', (' w:' + $attribute + '="' + $FontName + '"$1')
                }
            }
            $tag
        })
}

function Set-ExcelFontReferences {
    param(
        [Parameter(Mandatory = $true)][string]$Xml,
        [Parameter(Mandatory = $true)][string]$FontName
    )

    $updatedXml = [regex]::Replace($Xml, '(<name\s+val=")[^"]*("/>)', {
            param($nameMatch)
            "$($nameMatch.Groups[1].Value)$FontName$($nameMatch.Groups[2].Value)"
        })
    $updatedXml = [regex]::Replace($updatedXml, '<scheme\s+val="(?:major|minor)"\s*/>', '')

    [regex]::Replace($updatedXml, '<xf\b(?=[^>]*\bxfId="0")[^>]*/>', {
            param($xfMatch)
            $xf = $xfMatch.Value
            if ($xf -match '\bfontId="\d+"') {
                $xf = [regex]::Replace($xf, '\bfontId="\d+"', 'fontId="0"')
            } else {
                $xf = $xf -replace '(/>)$', ' fontId="0"$1'
            }
            if ($xf -notmatch '\bapplyFont=') {
                $xf = $xf -replace '(/>)$', ' applyFont="1"$1'
            }
            $xf
        })
}

$templateFiles = Get-ChildItem -LiteralPath $TemplateRoot -File | Where-Object {
    $_.Name -match '^(?:Normal|NormalEmail|Mappe)-.+-1[0-2]\.(?:dotm|xltx)$'
}

if (-not $templateFiles) {
    throw "Keine vorbereiteten Vorlagen in '$TemplateRoot' gefunden."
}

foreach ($templateFile in $templateFiles) {
    $match = [regex]::Match($templateFile.Name, '^(?:Normal|NormalEmail|Mappe)-(?<FontName>.+)-1[0-2]\.(?:dotm|xltx)$')
    $fontName = $match.Groups['FontName'].Value

    if (-not $PSCmdlet.ShouldProcess($templateFile.FullName, "Schriftreferenzen auf '$fontName' setzen")) {
        continue
    }

    if ($templateFile.Extension -ieq '.dotm') {
        [void](Update-ZipXmlEntry -ArchivePath $templateFile.FullName -EntryName 'word/theme/theme1.xml' -Transform {
                param($xml)
                Set-ThemeLatinFonts -Xml $xml -FontName $fontName
            })
        foreach ($entryName in @('word/styles.xml', 'word/document.xml')) {
            [void](Update-ZipXmlEntry -ArchivePath $templateFile.FullName -EntryName $entryName -Transform {
                    param($xml)
                    Set-WordFontReferences -Xml $xml -FontName $fontName
                })
        }
    } else {
        [void](Update-ZipXmlEntry -ArchivePath $templateFile.FullName -EntryName 'xl/theme/theme1.xml' -Transform {
                param($xml)
                Set-ThemeLatinFonts -Xml $xml -FontName $fontName
            })
        [void](Update-ZipXmlEntry -ArchivePath $templateFile.FullName -EntryName 'xl/styles.xml' -Transform {
                param($xml)
                Set-ExcelFontReferences -Xml $xml -FontName $fontName
            })
    }

    Write-Output "Repariert: $($templateFile.Name) -> $fontName"
}
