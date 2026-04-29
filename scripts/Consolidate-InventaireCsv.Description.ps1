<#
.SYNOPSIS
    Fonctions de selection et de generation de descriptions a partir de templates.
#>

function Select-DescriptionType {
    param(
        [AllowNull()]
        [string]$TypeBoitier,

        [AllowNull()]
        [string]$Model,

        [AllowNull()]
        [string]$Brand,

        [AllowNull()]
        [string]$Motherboard
    )

    $typeBoitierKey = Normalize-Key $TypeBoitier
    $modelKey = Normalize-Key $Model
    $brandKey = Normalize-Key $Brand
    $motherboardKey = Normalize-Key $Motherboard
    $haystack = "{0} {1} {2} {3}" -f $typeBoitierKey, $modelKey, $brandKey, $motherboardKey

    if ($typeBoitierKey -match '\b(laptop|notebook|portable)\b' -or $modelKey -match '\b(latitude|thinkpad|probook|elitebook)\b') {
        return 'Portable'
    }

    if ($haystack -match '\b(dell\s*3040|optiplex\s*3050|optiplex\s*3046|thinkcentre\s*m900)\b') {
        return 'Mini'
    }

    if ($haystack -match '\b(micro|tiny|mini|usff|ultra small form factor)\b') {
        return 'Mini'
    }

    if ($haystack -match '\b(sff|small form factor|space[-\s]*saving)\b') {
        return 'SFF'
    }

    if ($haystack -match '\b(tower|tour|mt|mini tower)\b') {
        return 'Tour'
    }

    return 'Tour'
}

function Select-DescriptionTemplatePath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateMap,

        [AllowNull()]
        [string]$Brand,

        [AllowNull()]
        [string]$TypeBoitier,

        [AllowNull()]
        [string]$Model,

        [AllowNull()]
        [string]$Motherboard
    )

    $descriptionType = Select-DescriptionType -TypeBoitier $TypeBoitier -Model $Model -Brand $Brand -Motherboard $Motherboard
    $templateKey = switch ($descriptionType) {
        'Portable' { 'pcportable' }
        'Mini'     { 'mini' }
        'SFF'      { 'sff' }
        default    { 'tour' }
    }

    if ($TemplateMap.ContainsKey($templateKey)) {
        return $TemplateMap[$templateKey]
    }

    $fallbackKey = $TemplateMap.Keys |
        Where-Object { $_ -like "*$templateKey*" } |
        Sort-Object Length |
        Select-Object -First 1
    if (-not [string]::IsNullOrWhiteSpace($fallbackKey)) {
        return $TemplateMap[$fallbackKey]
    }

    if ($TemplateMap.ContainsKey('tour')) {
        return $TemplateMap['tour']
    }

    return $null
}

function Get-DescriptionFromTemplate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$TemplateText,

        [Parameter(Mandatory = $true)]
        [psobject]$ComputerData,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    if ([string]::IsNullOrWhiteSpace($TemplateText)) {
        return $MissingValue
    }

    $valueLookup = @{}
    foreach ($property in $ComputerData.PSObject.Properties) {
        $key = Convert-ToLookupKey $property.Name
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $valueLookup[$key] = [string]$property.Value
        }
    }

    $valueLookup.Remove((Convert-ToLookupKey 'TauxUsure')) | Out-Null

    $aliasToProperty = @{
        'marque'                  = 'NomMarqueOrdinateur'
        'nommarque'               = 'NomMarqueOrdinateur'
        'modele'                  = 'ModeleOrdinateur'
        'processeur'              = 'Processeur'
        'memoiretotale'           = 'MemoireTotale'
        'typeddrsupporte'         = 'TypeDDRSupporte'
        'typedisque'              = 'TypeDisque'
        'tailledisque'            = 'CapaciteSSD'
        'modelessd'               = 'ModeleSSD'
        'cartegraphique'          = 'CarteGraphique'
        'systeme'                 = 'SystemeOperateur'
        'typeboit'                = 'TypeBoitier'
        'typeboitier'             = 'TypeBoitier'
        'tailleecran'             = 'TailleEcran'
        'nombrecoeurs'            = 'NombreCoeurs'
        'nombreprocesseurslogiques' = 'NombreProcesseursLogiques'
        'modelecartemere'         = 'ModeleCarteMere'
        'modele carte mere'       = 'ModeleCarteMere'
        'modele de carte mere'    = 'ModeleCarteMere'
        'carte mere modele'       = 'ModeleCarteMere'
    }

    foreach ($alias in $aliasToProperty.Keys) {
        $propertyName = $aliasToProperty[$alias]
        $propertyKey = Convert-ToLookupKey $propertyName
        if ($valueLookup.ContainsKey($propertyKey) -and -not [string]::IsNullOrWhiteSpace($valueLookup[$propertyKey]) -and $valueLookup[$propertyKey] -ne $MissingValue) {
            $valueLookup[(Convert-ToLookupKey $alias)] = $valueLookup[$propertyKey]
        }
    }

    $missingSentinel = '__MISSING_VALUE__'
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($match)

        $token = $match.Groups['token'].Value
        $lookupKey = Convert-ToLookupKey $token
        if ($valueLookup.ContainsKey($lookupKey)) {
            $lookupValue = [string]$valueLookup[$lookupKey]
            if (-not [string]::IsNullOrWhiteSpace($lookupValue) -and $lookupValue -ne $MissingValue) {
                return $lookupValue
            }
        }

        return $missingSentinel
    }

    $result = [regex]::Replace($TemplateText, '\[(?<token>[^\[\]\r\n]+)\]', $evaluator)
    $result = [regex]::Replace($result, '^[^\r\n]*' + [regex]::Escape($missingSentinel) + '[^\r\n]*\r?\n?', '', [Text.RegularExpressions.RegexOptions]::Multiline)
    $result = [regex]::Replace($result, '^[^\r\n]*\[[^\]\r\n]+\][^\r\n]*\r?\n?', '', [Text.RegularExpressions.RegexOptions]::Multiline)
    $result = [regex]::Replace($result, '(\r\n|\r|\n){2,}', "`r`n")
    $result = $result.Trim()

    if ([string]::IsNullOrWhiteSpace($result)) {
        return $MissingValue
    }

    return $result
}
