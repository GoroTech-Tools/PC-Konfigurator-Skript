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
        'PSShouldProcess',
        # Interaktives Zustands-Skript: Scope-Variablen für Laufzeit-Kontext sind beabsichtigt
        'PSAvoidGlobalVars',
        # Für dieses Projekt nicht zielführend: ShouldProcess-Warnungen auf Set/Remove-Funktionen
        'PSUseShouldProcessForStateChangingFunctions',
        # Einzelne optionale/kompatible Parameter dürfen bewusst ungenutzt sein
        'PSReviewUnusedParameter',
        # Projekt nutzt bewusst den etablierten Logger-Namen Write-Log
        'PSAvoidOverwritingBuiltInCmdlets',
        # Projekt nutzt UTF-8 ohne BOM bewusst konsistent
        'PSUseBOMForUnicodeEncodedFile'
    )

    Severity = @(
        'Error',
        'Warning'
    )

    IncludeDefaultRules = $true
}
