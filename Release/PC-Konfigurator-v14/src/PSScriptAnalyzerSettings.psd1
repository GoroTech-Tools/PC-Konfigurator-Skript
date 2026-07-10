@{
    # Projektweite ScriptAnalyzer-Konfiguration
    # PSUseApprovedVerbs wird bewusst ausgeschlossen, da historisch gewachsene
    # Funktionsnamen im Projekt beibehalten werden sollen.
    ExcludeRules = @(
        'PSUseApprovedVerbs',
        # Interaktives Konsolenskript: Write-Host ist hier bewusst eingesetzt
        'PSAvoidUsingWriteHost',
        # Historische Funktionsnamen im Projekt sind teils plural
        'PSUseSingularNouns',
        # Für dieses Skript wird ShouldProcess nicht flächendeckend erzwungen
        'PSShouldProcess'
    )

    Severity = @(
        'Error',
        'Warning'
    )

    IncludeDefaultRules = $true
}
