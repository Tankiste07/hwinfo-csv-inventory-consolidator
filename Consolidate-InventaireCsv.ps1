#requires -Version 5.1
<#!
.SYNOPSIS
    Consolide des fichiers CSV d'inventaire (format cle,valeur) vers un fichier Excel.

.DESCRIPTION
    - Parcourt tous les CSV d'un dossier.
    - Extrait les informations importantes par machine.
    - Exporte un fichier Excel final avec une ligne par ordinateur.

.PARAMETER Folder
    Dossier contenant les CSV source.

.PARAMETER Output
    Fichier Excel de sortie (.xlsx).

.EXAMPLE
    .\Consolidate-InventaireCsv.ps1 -Folder "C:\Inventaire\CSV" -Output "C:\Inventaire\InventaireFinal.xlsx"

.EXAMPLE
    .\Consolidate-InventaireCsv.ps1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Folder = 'E:\Programs\HWInfo',

    [Parameter(Mandatory = $false)]
    [string]$Output,

    [Parameter(Mandatory = $false)]
    [string]$DescriptionTemplateFolder = 'E:\Programs\HWInfo\SquelletteDescription'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MissingValue = 'N/A'

$script:LogDirectory = Join-Path -Path $PSScriptRoot -ChildPath 'logs'
$script:LogFile = Join-Path -Path $script:LogDirectory -ChildPath ("inventaire-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

if (-not (Test-Path -Path $script:LogDirectory -PathType Container)) {
    New-Item -Path $script:LogDirectory -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    try {
        Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    }
    catch {
        # Ne bloque jamais le traitement principal si le log echoue.
    }
}

function Write-AppWarning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[ATTENTION] $Message"
    Write-Log -Level 'WARN' -Message $Message
}

function Write-AppInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host $Message
    Write-Log -Level 'INFO' -Message $Message
}

Write-Log -Level 'INFO' -Message "Demarrage du script. Dossier source='$Folder', sortie='$Output'"

trap {
    $trapMessage = $_.Exception.Message
    Write-Host "[ERREUR] $trapMessage"
    Write-Log -Level 'ERROR' -Message $trapMessage
    exit 99
}

function Exit-WithMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [int]$Code = 1
    )

    Write-Host "[ERREUR] $Message"
    Write-Log -Level 'ERROR' -Message $Message
    exit $Code
}

# Nom de fichier Excel fixe (toujours le meme, ecrase a chaque execution).
$OutputFileName = 'Inventaire_HWInfo.xlsx'

# Si non fourni, genere le chemin fixe dans le dossier source.
if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Join-Path -Path $Folder -ChildPath $OutputFileName
}

# Supprime l'ancien fichier Excel s'il existe deja.
if (Test-Path -Path $Output) {
    try {
        Remove-Item -Path $Output -Force -ErrorAction Stop
    }
    catch {
        Exit-WithMessage -Message "Impossible de supprimer l'ancien fichier Excel '$Output' (fichier ouvert ?). Fermez-le et relancez." -Code 20
    }
}

# Verifie que le dossier source existe
if (-not (Test-Path -Path $Folder -PathType Container)) {
    Exit-WithMessage -Message "Le dossier source n'existe pas : $Folder" -Code 11
}

# Verifie la presence du module ImportExcel
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Exit-WithMessage -Message "Le module 'ImportExcel' est introuvable. Lance d'abord .\\Install-Dependencies.ps1" -Code 12
}
Import-Module ImportExcel -ErrorAction Stop

# Variables de contexte utilisees par GetValue
$script:CurrentIndex = @{}
$script:CurrentFile = ''

function Normalize-Key {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $t = $Text.Trim().Trim('"')
    $t = $t -replace ':\s*$', ''
    $t = $t.ToLowerInvariant()

    # Supprime les accents pour faire correspondre systeme/systeme, securise/secure, etc.
    $formD = $t.Normalize([Text.NormalizationForm]::FormD)
    $chars = New-Object System.Collections.Generic.List[char]
    foreach ($c in $formD.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$chars.Add($c)
        }
    }

    $t = -join $chars
    $t = $t.Normalize([Text.NormalizationForm]::FormC)
    $t = $t -replace '\s+', ' '
    return $t.Trim()
}

function GetValue($key) {
    <#
    .SYNOPSIS
        Retourne la valeur associee a une cle dans le CSV courant  .

    .DESCRIPTION
        - Recherche exacte de la cle dans l'index en memoire.
        - Retourne la premiere valeur non vide trouvee.
        - Si la cle est absente, retourne $null et ecrit un warning.
    #>
    if ([string]::IsNullOrWhiteSpace($key)) {
        return $null
    }

    $normalizedKey = Normalize-Key $key
    if ([string]::IsNullOrWhiteSpace($normalizedKey)) {
        return $null
    }

    if ($script:CurrentIndex.ContainsKey($normalizedKey)) {
        $value = $script:CurrentIndex[$normalizedKey] |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 1

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return $null
}

function GetFirstValue {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Keys,

        [Parameter(Mandatory = $true)]
        [string]$FieldName,

        [Parameter(Mandatory = $false)]
        [string[]]$FallbackContains,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$DefaultValue = $null
    )

    foreach ($k in $Keys) {
        $v = GetValue $k
        if (-not [string]::IsNullOrWhiteSpace($v)) {
            return $v
        }
    }

    if ($FallbackContains -and $FallbackContains.Count -gt 0) {
        foreach ($indexKey in $script:CurrentIndex.Keys) {
            $allMatch = $true
            foreach ($token in $FallbackContains) {
                $nt = Normalize-Key $token
                if ([string]::IsNullOrWhiteSpace($nt)) {
                    continue
                }

                if ($indexKey -notlike "*$nt*") {
                    $allMatch = $false
                    break
                }
            }

            if ($allMatch) {
                $fallbackValue = $script:CurrentIndex[$indexKey] |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -First 1

                if (-not [string]::IsNullOrWhiteSpace($fallbackValue)) {
                    return $fallbackValue
                }
            }
        }
    }

    # Gestion propre des cles absentes : un warning par champ
    Write-AppWarning "[$script:CurrentFile] Champ introuvable : '$FieldName'"
    if ($PSBoundParameters.ContainsKey('DefaultValue')) {
        return $DefaultValue
    }

    return $null
}

function GetFirstValueNoWarning {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Keys,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$DefaultValue = $null
    )

    foreach ($k in $Keys) {
        $v = GetValue $k
        if (-not [string]::IsNullOrWhiteSpace($v)) {
            return $v
        }
    }

    if ($PSBoundParameters.ContainsKey('DefaultValue')) {
        return $DefaultValue
    }

    return $null
}

function Test-IsSupportedValue {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $v = Normalize-Key $Value

    if ($v -match '^(non|not supported|unsupported|false|no|0)') {
        return $false
    }

    if ($v -match 'support|present|presen|active|activee|enabled|enable|oui|yes|true|1') {
        return $true
    }

    return $false
}

function Normalize-OperatingSystem {
    param(
        [AllowNull()]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq $MissingValue) {
        return $MissingValue
    }

    $v = Normalize-Key $Value
    if ($v -match 'windows\s*11') {
        return 'Windows 11'
    }

    if ($v -match 'windows\s*10') {
        return 'Windows 10'
    }

    return $Value.Trim()
}

function Normalize-StorageCapacity {
    param(
        [AllowNull()]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq $MissingValue) {
        return $MissingValue
    }

    # Priorite a la valeur entre parentheses, ex: "244,198 Megaoctets (256 Go)" -> "256 Go".
    $match = [regex]::Match($Value, '\(([0-9]+(?:[\.,][0-9]+)?)\s*(Go|GB|To|TB)\)', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        $num = $match.Groups[1].Value -replace '\.0+$', ''
        $unit = $match.Groups[2].Value.ToUpperInvariant()
        if ($unit -eq 'GO') { $unit = 'Go' }
        if ($unit -eq 'GB') { $unit = 'Go' }
        if ($unit -eq 'TO') { $unit = 'To' }
        if ($unit -eq 'TB') { $unit = 'To' }
        return ('{0} {1}' -f $num, $unit)
    }

    # Sinon, extrait une capacite deja exprimee en Go/To.
    $match = [regex]::Match($Value, '([0-9]+(?:[\.,][0-9]+)?)\s*(Go|GB|To|TB)', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        $num = $match.Groups[1].Value -replace '\.0+$', ''
        $unit = $match.Groups[2].Value.ToUpperInvariant()
        if ($unit -eq 'GO') { $unit = 'Go' }
        if ($unit -eq 'GB') { $unit = 'Go' }
        if ($unit -eq 'TO') { $unit = 'To' }
        if ($unit -eq 'TB') { $unit = 'To' }
        return ('{0} {1}' -f $num, $unit)
    }

    return $Value.Trim()
}

function Convert-ToLookupKey {
    param(
        [AllowNull()]
        [string]$Text
    )

    $normalized = Normalize-Key $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    return ($normalized -replace '[^a-z0-9]', '')
}

# Regle metier specifique pour les DELL Latitude : la taille d'ecran en pouces est deduite 
#du chiffre des centaines dans le numero de modele a 4 chiffres (ex: Latitude 5490 => 14 pouces, 
#Latitude 5590 => 15 pouces, etc.).

function Get-DellLatitudeScreenSizeInches {
    param(
        [AllowNull()]
        [string]$Brand
    )

    if ([string]::IsNullOrWhiteSpace($Brand)) {
        return $null
    }

    $brandNormalized = Normalize-Key $Brand
    if ($brandNormalized -notlike '*dell*' -or $brandNormalized -notlike '*latitude*') {
        return $null
    }

    # Regle metier demandee : sur un modele a 4 chiffres, le chiffre des centaines
    # indique la taille d'ecran (4 => 14 pouces, 5 => 15 pouces, etc.).
    $modelMatch = [regex]::Match($Brand, '(?i)latitude[^0-9]*(?<model>\d{4})')
    if (-not $modelMatch.Success) {
        $modelMatch = [regex]::Match($Brand, '(?<model>\d{4})')
    }

    if (-not $modelMatch.Success) {
        return $null
    }

    $modelRaw = $modelMatch.Groups['model'].Value
    $modelNumber = 0
    if (-not [int]::TryParse($modelRaw, [ref]$modelNumber)) {
        return $null
    }

    $hundredsDigit = [math]::Floor(($modelNumber % 1000) / 100)
    if ($hundredsDigit -lt 1 -or $hundredsDigit -gt 9) {
        return $null
    }

    return [int](10 + $hundredsDigit)
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

    # Valeurs derivees utiles pour les templates marketing (ex: [ramgo], [ssdgo]).
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

        # Fallback souple : autorise des balises "naturelles" proches des noms de colonnes.
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

    # Fallback final pour les DELL si aucun template specifique n'existe.
    if ($TemplateMap.ContainsKey('dell')) {
        return $TemplateMap['dell']
    }

    return $null
}

# Recupere tous les CSV du dossier
$csvFiles = Get-ChildItem -Path $Folder -Filter '*.csv' -File
if (-not $csvFiles) {
    Exit-WithMessage -Message "Pas de CSV dans : $Folder" -Code 10
}

# Liste finale d'objets PowerShell (1 objet = 1 PC)
$results = New-Object System.Collections.Generic.List[object]

$descriptionTemplates = @{}
if (Test-Path -Path $DescriptionTemplateFolder -PathType Container) {
    $templateFiles = Get-ChildItem -Path $DescriptionTemplateFolder -Filter '*.txt' -File -ErrorAction SilentlyContinue
    foreach ($templateFile in $templateFiles) {
        $templateKey = Convert-ToLookupKey $templateFile.BaseName
        if (-not [string]::IsNullOrWhiteSpace($templateKey) -and -not $descriptionTemplates.ContainsKey($templateKey)) {
            $descriptionTemplates[$templateKey] = $templateFile.FullName
        }
    }
}
else {
    Write-AppWarning "Dossier des descriptions introuvable : '$DescriptionTemplateFolder'. La colonne DescriptionOrdi sera a '$MissingValue'."
}

foreach ($file in $csvFiles) {
    try {
        $script:CurrentFile = $file.Name

        # Parse ligne par ligne pour supporter les exports HWInfo (sections + format "cle","valeur")
        $lines = Get-Content -Path $file.FullName

        # Construit un index cle -> liste de valeurs
        $script:CurrentIndex = @{}
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $k = $null
            $v = $null

            if ($line -match '^"(?<k>(?:[^"]|"")*)","(?<v>(?:[^"]|"")*)"$') {
                $k = ($Matches.k -replace '""', '"').Trim()
                $v = ($Matches.v -replace '""', '"').Trim()
            }
            elseif ($line -match '^(?<k>[^,]+),(?<v>.*)$') {
                $k = ($Matches.k -as [string]).Trim().Trim('"')
                $v = ($Matches.v -as [string]).Trim().Trim('"')
            }
            else {
                continue
            }

            if ([string]::IsNullOrWhiteSpace($k)) {
                continue
            }

            # Ignore une potentielle ligne d'entete
            $kn = Normalize-Key $k
            if ($kn -in @('key', 'cle') -and (Normalize-Key $v) -in @('value', 'valeur')) {
                continue
            }

            if ([string]::IsNullOrWhiteSpace($kn)) {
                continue
            }

            if (-not $script:CurrentIndex.ContainsKey($kn)) {
                $script:CurrentIndex[$kn] = New-Object System.Collections.Generic.List[string]
            }

            [void]$script:CurrentIndex[$kn].Add($v)
        }

        # Recuperation specifique du support DDR (sans warning si absent)
        $ddr4Value = GetFirstValueNoWarning -Keys @('DDR4', 'Support DDR4') -DefaultValue $MissingValue
        $ddr3lValue = GetFirstValueNoWarning -Keys @('DDR3L', 'Support DDR3L') -DefaultValue $MissingValue
        $ddr3Value = GetFirstValueNoWarning -Keys @('DDR3', 'Support DDR3') -DefaultValue $MissingValue

        $supportedDdrTypes = New-Object System.Collections.Generic.List[string]
        if (Test-IsSupportedValue $ddr4Value) { [void]$supportedDdrTypes.Add('DDR4') }
        if (Test-IsSupportedValue $ddr3lValue) { [void]$supportedDdrTypes.Add('DDR3L') }
        if (Test-IsSupportedValue $ddr3Value) { [void]$supportedDdrTypes.Add('DDR3') }

        if ($supportedDdrTypes.Count -gt 0) {
            $typeDdrSupporte = ($supportedDdrTypes -join ', ')
        }
        elseif ($ddr4Value -eq $MissingValue -and $ddr3lValue -eq $MissingValue -and $ddr3Value -eq $MissingValue) {
            $typeDdrSupporte = $MissingValue
        }
        else {
            $typeDdrSupporte = 'Aucun'
        }

        $nomMarqueOrdinateur = GetFirstValue -FieldName 'NomMarqueOrdinateur' -Keys @("Nom de marque de l'ordinateur", 'Computer Brand Name', 'Brand Name', 'System Brand') -FallbackContains @('marque', 'ordinateur') -DefaultValue $MissingValue

        # Cree l'objet final pour ce PC
        $pcObject = [PSCustomObject]@{
            SystemeOperateur          = Normalize-OperatingSystem -Value (GetFirstValue -FieldName 'SystemeOperateur' -Keys @('Systeme operateur', 'OperatingSystem', 'OS', 'Caption') -DefaultValue $MissingValue) -MissingValue $MissingValue
            NomMarqueOrdinateur       = $nomMarqueOrdinateur
            TypeBoitier               = GetFirstValue -FieldName 'TypeBoitier' -Keys @('Type de boitier', 'Type de boîtier', 'Chassis Type', 'Type de chassis') -DefaultValue $MissingValue
            NumeroSerie               = GetFirstValue -FieldName 'NumeroSerie' -Keys @('Numero de serie', 'Numéro de série', 'Serial Number', 'SerialNumber', 'System Serial Number', 'BIOS serial number') -FallbackContains @('serial') -DefaultValue $MissingValue
            NomProcesseur             = GetFirstValue -FieldName 'NomProcesseur' -Keys @('Nom du processeur', 'ProcessorName', 'CPUName', 'CPU') -DefaultValue $MissingValue
            NombreCoeurs              = GetFirstValue -FieldName 'NombreCoeurs' -Keys @('Nombre de coeurs de processeur', 'NumberOfCores', 'CPUCores', 'Cores') -DefaultValue $MissingValue
            NombreProcesseursLogiques = GetFirstValue -FieldName 'NombreProcesseursLogiques' -Keys @('Nombre de processeurs logiques', 'NumberOfLogicalProcessors', 'LogicalProcessors') -DefaultValue $MissingValue
            MemoireTotale             = GetFirstValue -FieldName 'MemoireTotale' -Keys @('Taille totale de la memoire', 'TotalMemory', 'MemoryTotal', 'RAM', 'Memoire physique totale') -DefaultValue $MissingValue
            TypeDDRSupporte           = $typeDdrSupporte
            TauxUsure                 = GetFirstValue -FieldName 'TauxUsure' -Keys @('Taux d usure', 'Taux d usure de la batterie', "Taux d'usure", 'BatteryWearLevel', 'BatteryWear', 'Usure batterie') -FallbackContains @('taux', 'usure') -DefaultValue $MissingValue
            JeuDePucesGraphiques      = GetFirstValue -FieldName 'JeuDePucesGraphiques' -Keys @('Jeu de puces graphiques', 'Graphics Chipset', 'Graphic Chipset', 'GPU Chipset', 'Nom de la puce graphique') -FallbackContains @('puces', 'graphi') -DefaultValue $MissingValue
            ModeleSSD                 = GetFirstValue -FieldName 'ModeleSSD' -Keys @('Modele du SSD', 'SSDModel', 'DiskModel', 'StorageModel', 'Modele de lecteur') -DefaultValue $MissingValue
            CapaciteSSD               = Normalize-StorageCapacity -Value (GetFirstValue -FieldName 'CapaciteSSD' -Keys @('Capacite du SSD', 'SSDCapacity', 'DiskSize', 'StorageCapacity', 'Capacite du lecteur') -DefaultValue $MissingValue) -MissingValue $MissingValue
            ModeleCarteMere           = GetFirstValue -FieldName 'ModeleCarteMere' -Keys @('Modele de carte mere', 'BaseBoardModel', 'MotherboardModel', 'Carte mere') -DefaultValue $MissingValue
            }

        $descriptionOrdi = $MissingValue
        $matchedTemplatePath = Select-DescriptionTemplatePath -TemplateMap $descriptionTemplates -Brand $nomMarqueOrdinateur -TypeBoitier ([string]$pcObject.TypeBoitier)
        if ($null -ne $matchedTemplatePath) {
            $templatePath = $matchedTemplatePath
            try {
                $templateContent = Get-Content -Path $templatePath -Raw -Encoding UTF8
                $descriptionOrdi = Get-DescriptionFromTemplate -TemplateText $templateContent -ComputerData $pcObject -MissingValue $MissingValue
            }
            catch {
                Write-AppWarning "[$script:CurrentFile] Impossible de lire le fichier de description '$templatePath' : $($_.Exception.Message)"
            }
        }

        $pcObject | Add-Member -NotePropertyName 'DescriptionOrdi' -NotePropertyValue $descriptionOrdi

        [void]$results.Add($pcObject)
    }
    catch {
        Write-AppWarning "Fichier ignore '$($file.Name)' : $($_.Exception.Message)"
        continue
    }
}

if ($results.Count -eq 0) {
    Exit-WithMessage -Message "Aucune donnee exploitable n'a ete extraite des CSV." -Code 13
}

# Exporte toutes les lignes vers Excel (AutoSize active)
$results |
    Export-Excel -Path $Output `
                 -WorksheetName 'Inventaire' `
                 -TableName 'InventairePC' `
                 -AutoSize `
                 -BoldTopRow `
                 -FreezeTopRow `
                 -ClearSheet

# Met en rouge toute ligne dont le taux d'usure est > 38%
$package = Open-ExcelPackage -Path $Output
try {
    $ws = $package.Workbook.Worksheets['Inventaire']
    if ($null -ne $ws -and $null -ne $ws.Dimension) {
        $maxRow = $ws.Dimension.End.Row
        $maxCol = $ws.Dimension.End.Column

        $tauxCol = $null
        for ($c = 1; $c -le $maxCol; $c++) {
            $header = $ws.Cells[1, $c].Text
            if ((Normalize-Key $header) -eq (Normalize-Key 'TauxUsure')) {
                $tauxCol = $c
                break
            }
        }

        if ($null -ne $tauxCol) {
            for ($r = 2; $r -le $maxRow; $r++) {
                $raw = $ws.Cells[$r, $tauxCol].Text
                if ([string]::IsNullOrWhiteSpace($raw) -or $raw -eq $MissingValue) {
                    continue
                }

                $numText = ($raw -replace '%', '').Trim()
                $numText = $numText -replace ',', '.'
                $wear = 0.0

                if ([double]::TryParse($numText, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$wear)) {
                    if ($wear -gt 38.0) {
                        $rowRange = $ws.Cells[$r, 1, $r, $maxCol]
                        $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightCoral)
                        $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
                    }
                }
            }
        }
    }
    $package.Save()
}
finally {
    if ($null -ne $package) {
        $package.Dispose()
    }
}

# Supprime les CSV source apres creation reussie du fichier Excel.
foreach ($csv in $csvFiles) {
    try {
        Remove-Item -Path $csv.FullName -Force -ErrorAction Stop
    }
    catch {
        Write-AppWarning "Impossible de supprimer '$($csv.FullName)' : $($_.Exception.Message)"
    }
}

Write-AppInfo "Export termine : $Output"
Write-AppInfo "Nombre de postes exportes : $($results.Count)"
Write-AppInfo "Log : $script:LogFile"
exit 0
