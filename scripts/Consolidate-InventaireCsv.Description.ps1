<#
.SYNOPSIS
    Fonctions de selection et de generation de descriptions a partir de templates.
#>

function Select-DescriptionTemplatePath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateMap,

        [AllowNull()]
        [string]$Brand,

        [AllowNull()]
        [string]$TypeBoitier
    )

    $brandKey = Convert-ToLookupKey $Brand
    if ([string]::IsNullOrWhiteSpace($brandKey) -or $brandKey -notlike '*dell*') {
        return $null
    }

    $brandLabel = Normalize-Key $Brand
    $boitierKey = Normalize-Key $TypeBoitier

    $candidate = 'dell'
    if ($brandLabel -match '\blatitude\b') {
        $candidate = 'delllatitude'
    }
    elseif ($boitierKey -match 'space[\s\-]*saving') {
        $candidate = 'dellspacesaving'
    }

    if ($TemplateMap.ContainsKey($candidate)) {
        return $TemplateMap[$candidate]
    }

    $matchKey = $null
    if ($candidate -eq 'dell') {
        $matchKey = $TemplateMap.Keys |
            Where-Object { $_ -like '*dell*' -and $_ -notlike '*latitude*' -and $_ -notlike '*spacesaving*' } |
            Sort-Object Length |
            Select-Object -First 1
    }
    else {
        $matchKey = $TemplateMap.Keys |
            Where-Object { $_ -like "*$candidate*" } |
            Sort-Object Length -Descending |
            Select-Object -First 1
    }

    if (-not [string]::IsNullOrWhiteSpace($matchKey)) {
        return $TemplateMap[$matchKey]
    }

    if ($TemplateMap.ContainsKey('dell')) {
        return $TemplateMap['dell']
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

    $memoireTotaleKey = Convert-ToLookupKey 'MemoireTotale'
    if ($valueLookup.ContainsKey($memoireTotaleKey)) {
        $mem = [string]$valueLookup[$memoireTotaleKey]
        $memMatch = [regex]::Match($mem, '([0-9]+(?:[\.,][0-9]+)?)')
        if ($memMatch.Success) {
            $memNum = $memMatch.Groups[1].Value
            $valueLookup['ramgo'] = $memNum
            $valueLookup['memoirego'] = $memNum
        }
    }

    $capaciteSsdKey = Convert-ToLookupKey 'CapaciteSSD'
    if ($valueLookup.ContainsKey($capaciteSsdKey)) {
        $ssd = [string]$valueLookup[$capaciteSsdKey]
        $ssdMatch = [regex]::Match($ssd, '([0-9]+(?:[\.,][0-9]+)?)')
        if ($ssdMatch.Success) {
            $ssdNum = $ssdMatch.Groups[1].Value
            $valueLookup['ssdgo'] = $ssdNum
            $valueLookup['capacitessdgo'] = $ssdNum
        }
    }

    $aliasToProperty = @{
        'marque'      = 'NomMarqueOrdinateur'
        'nommarque'   = 'NomMarqueOrdinateur'
        'serial'      = 'NumeroSerie'
        'numeroserie' = 'NumeroSerie'
        'processeur'  = 'NomProcesseur'
        'cpu'         = 'NomProcesseur'
        'coeur'       = 'NombreCoeurs'
        'coeurs'      = 'NombreCoeurs'
        'cœur'        = 'NombreCoeurs'
        'cœurs'       = 'NombreCoeurs'
        'nombredecoeurs' = 'NombreCoeurs'
        'threads'     = 'NombreProcesseursLogiques'
        'thread'      = 'NombreProcesseursLogiques'
        'nombrethreads' = 'NombreProcesseursLogiques'
        'os'          = 'SystemeOperateur'
        'systeme'     = 'SystemeOperateur'
        'ram'         = 'MemoireTotale'
        'memoire'     = 'MemoireTotale'
        'memoireddr3l' = 'MemoireTotale'
        'memoireddr4' = 'MemoireTotale'
        'ssd'         = 'CapaciteSSD'
        'cartegraphique' = 'Carte graphique'
        'carte graphique' = 'Carte graphique'
    }

    foreach ($alias in $aliasToProperty.Keys) {
        $propertyName = $aliasToProperty[$alias]
        $propertyKey = Convert-ToLookupKey $propertyName
        if ($valueLookup.ContainsKey($propertyKey) -and -not [string]::IsNullOrWhiteSpace($valueLookup[$propertyKey])) {
            $valueLookup[(Convert-ToLookupKey $alias)] = $valueLookup[$propertyKey]
        }
    }

    $latitudeScreenSize = $null
    $brandValue = [string]$ComputerData.NomMarqueOrdinateur
    $latitudeScreenSize = Get-DellLatitudeScreenSizeInches -Brand $brandValue
    if ($null -ne $latitudeScreenSize) {
        $screenValue = ('{0} pouces' -f $latitudeScreenSize)
        $valueLookup['tailleecran'] = $screenValue
        $valueLookup['tailleecranpouces'] = $screenValue
        $valueLookup['ecran'] = $screenValue
        $valueLookup['ecranpouces'] = $screenValue
        $valueLookup['pouces'] = [string]$latitudeScreenSize
    }

    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($match)

        $token = $match.Groups['token'].Value
        $lookupKey = Convert-ToLookupKey $token
        if ($valueLookup.ContainsKey($lookupKey) -and -not [string]::IsNullOrWhiteSpace($valueLookup[$lookupKey])) {
            return $valueLookup[$lookupKey]
        }

        $tokenLoose = $lookupKey -replace '^(type|nombre|nom|de|du|des|la|le)+', ''
        if (-not [string]::IsNullOrWhiteSpace($tokenLoose)) {
            $bestKey = $valueLookup.Keys |
                Where-Object {
                    $k = [string]$_
                    -not [string]::IsNullOrWhiteSpace($k) -and
                    ($k -like "*$lookupKey*" -or $lookupKey -like "*$k*" -or $k -like "*$tokenLoose*" -or $tokenLoose -like "*$k*")
                } |
                Sort-Object Length -Descending |
                Select-Object -First 1

            if (-not [string]::IsNullOrWhiteSpace($bestKey) -and -not [string]::IsNullOrWhiteSpace($valueLookup[$bestKey])) {
                return $valueLookup[$bestKey]
            }
        }

        return $MissingValue
    }

    $result = [regex]::Replace($TemplateText, '\[(?<token>[^\[\]\r\n]+)\]', $evaluator)
    $result = [regex]::Replace($result, '(?i)\b(go|to)\s+\1\b', '$1')
    if ([string]::IsNullOrWhiteSpace($result)) {
        return $MissingValue
    }

    if ($null -ne $latitudeScreenSize -and
        $result -ne $MissingValue -and
        $result -notmatch '(?i)\b[0-9]{2}\s*pouces?\b') {
        $result = ('{0} - Ecran {1} pouces' -f $result.Trim().TrimEnd('.'), $latitudeScreenSize)
    }

    return $result.Trim()
}
