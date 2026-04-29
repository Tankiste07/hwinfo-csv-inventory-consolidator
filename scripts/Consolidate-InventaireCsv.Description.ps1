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

    $haystack = Normalize-Key "${TypeBoitier} ${Model} ${Brand} ${Motherboard}"

    if ($haystack -match '\b(laptop|notebook|portable|latitude|thinkpad|probook|elitebook)\b') {
        return 'Portable'
    }

    if ($haystack -match '\b(dell\s*3040|optiplex\s*3050|optiplex\s*3046|thinkcentre\s*m900|micro|tiny|mini|usff|ultra small form factor)\b') {
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

function Get-DescriptionLookupFromComputerData {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$ComputerData,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    $lookup = @{}
    foreach ($property in $ComputerData.PSObject.Properties) {
        $key = Convert-ToLookupKey $property.Name
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $lookup[$key] = [string]$property.Value
        }
    }

    $lookup.Remove((Convert-ToLookupKey 'TauxUsure')) | Out-Null

    $aliasToProperty = @{
        'marque'                     = 'NomMarqueOrdinateur'
        'nommarque'                  = 'NomMarqueOrdinateur'
        'modele'                     = 'ModeleOrdinateur'
        'processeur'                 = 'Processeur'
        'memoiretotale'              = 'MemoireTotale'
        'typeddrsupporte'            = 'TypeDDRSupporte'
        'typedisque'                 = 'TypeDisque'
        'tailledisque'               = 'CapaciteSSD'
        'modelessd'                  = 'ModeleSSD'
        'cartegraphique'             = 'CarteGraphique'
        'systeme'                    = 'SystemeOperateur'
        'typeboit'                   = 'TypeBoitier'
        'typeboitier'                = 'TypeBoitier'
        'tailleecran'                = 'TailleEcran'
        'nombrecoeurs'               = 'NombreCoeurs'
        'nombreprocesseurslogiques'  = 'NombreProcesseursLogiques'
        'modelecartemere'            = 'ModeleCarteMere'
        'modele carte mere'          = 'ModeleCarteMere'
        'modele de carte mere'       = 'ModeleCarteMere'
        'carte mere modele'          = 'ModeleCarteMere'
    }

    foreach ($alias in $aliasToProperty.Keys) {
        $propertyKey = (Convert-ToLookupKey $aliasToProperty[$alias])
        if ($lookup.ContainsKey($propertyKey)) {
            $value = [string]$lookup[$propertyKey]
            if (-not [string]::IsNullOrWhiteSpace($value) -and $value -ne $MissingValue) {
                $lookup[(Convert-ToLookupKey $alias)] = $value
            }
        }
    }

    return $lookup
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

    $lookup = Get-DescriptionLookupFromComputerData -ComputerData $ComputerData -MissingValue $MissingValue
    $missingSentinel = '__MISSING__'

    $result = [regex]::Replace($TemplateText, '\[(?<token>[^\[\]\r\n]+)\]', {
        param($match)

        $key = Convert-ToLookupKey $match.Groups['token'].Value
        if ($lookup.ContainsKey($key)) {
            $value = [string]$lookup[$key]
            if (-not [string]::IsNullOrWhiteSpace($value) -and $value -ne $MissingValue) {
                return $value
            }
        }

        return $missingSentinel
    })

    $result = [regex]::Replace($result, '^[^\r\n]*' + [regex]::Escape($missingSentinel) + '[^\r\n]*\r?\n?', '', [Text.RegularExpressions.RegexOptions]::Multiline)
    $result = [regex]::Replace($result, '^[^\r\n]*\[[^\]\r\n]+\][^\r\n]*\r?\n?', '', [Text.RegularExpressions.RegexOptions]::Multiline)
    $result = [regex]::Replace($result, '(\r\n|\r|\n){2,}', "`r`n")
    $result = $result.Trim()

    if ([string]::IsNullOrWhiteSpace($result)) {
        return $MissingValue
    }

    return $result
}
