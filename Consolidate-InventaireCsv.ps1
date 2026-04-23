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
    [string]$Output
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
        Retourne la valeur associee a une cle dans le CSV courant.

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

# Recupere tous les CSV du dossier
$csvFiles = Get-ChildItem -Path $Folder -Filter '*.csv' -File
if (-not $csvFiles) {
    Exit-WithMessage -Message "Pas de CSV dans : $Folder" -Code 10
}

# Liste finale d'objets PowerShell (1 objet = 1 PC)
$results = New-Object System.Collections.Generic.List[object]

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

        # Cree l'objet final pour ce PC
        $pcObject = [PSCustomObject]@{
            SystemeOperateur          = GetFirstValue -FieldName 'SystemeOperateur' -Keys @('Systeme operateur', 'OperatingSystem', 'OS', 'Caption') -DefaultValue $MissingValue
            NomMarqueOrdinateur       = GetFirstValue -FieldName 'NomMarqueOrdinateur' -Keys @("Nom de marque de l'ordinateur", 'Computer Brand Name', 'Brand Name', 'System Brand') -FallbackContains @('marque', 'ordinateur') -DefaultValue $MissingValue
            NumeroSerie               = GetFirstValue -FieldName 'NumeroSerie' -Keys @('Numero de serie', 'Numéro de série', 'Serial Number', 'SerialNumber', 'System Serial Number', 'BIOS serial number') -FallbackContains @('serial') -DefaultValue $MissingValue
            NomProcesseur             = GetFirstValue -FieldName 'NomProcesseur' -Keys @('Nom du processeur', 'ProcessorName', 'CPUName', 'CPU') -DefaultValue $MissingValue
            NombreCoeurs              = GetFirstValue -FieldName 'NombreCoeurs' -Keys @('Nombre de coeurs de processeur', 'NumberOfCores', 'CPUCores', 'Cores') -DefaultValue $MissingValue
            NombreProcesseursLogiques = GetFirstValue -FieldName 'NombreProcesseursLogiques' -Keys @('Nombre de processeurs logiques', 'NumberOfLogicalProcessors', 'LogicalProcessors') -DefaultValue $MissingValue
            MemoireTotale             = GetFirstValue -FieldName 'MemoireTotale' -Keys @('Taille totale de la memoire', 'TotalMemory', 'MemoryTotal', 'RAM', 'Memoire physique totale') -DefaultValue $MissingValue
            TypeDDRSupporte           = $typeDdrSupporte
            TauxUsure                 = GetFirstValue -FieldName 'TauxUsure' -Keys @('Taux d usure', 'Taux d usure de la batterie', "Taux d'usure", 'BatteryWearLevel', 'BatteryWear', 'Usure batterie') -FallbackContains @('taux', 'usure') -DefaultValue $MissingValue
            JeuDePucesGraphiques      = GetFirstValue -FieldName 'JeuDePucesGraphiques' -Keys @('Jeu de puces graphiques', 'Graphics Chipset', 'Graphic Chipset', 'GPU Chipset', 'Nom de la puce graphique') -FallbackContains @('puces', 'graphi') -DefaultValue $MissingValue
            ModeleSSD                 = GetFirstValue -FieldName 'ModeleSSD' -Keys @('Modele du SSD', 'SSDModel', 'DiskModel', 'StorageModel', 'Modele de lecteur') -DefaultValue $MissingValue
            CapaciteSSD               = GetFirstValue -FieldName 'CapaciteSSD' -Keys @('Capacite du SSD', 'SSDCapacity', 'DiskSize', 'StorageCapacity', 'Capacite du lecteur') -DefaultValue $MissingValue
            ModeleCarteMere           = GetFirstValue -FieldName 'ModeleCarteMere' -Keys @('Modele de carte mere', 'BaseBoardModel', 'MotherboardModel', 'Carte mere') -DefaultValue $MissingValue
            VersionBIOS               = GetFirstValue -FieldName 'VersionBIOS' -Keys @('Version du BIOS', 'BIOSVersion', 'SMBIOSBIOSVersion', 'Version du BIOS du systeme') -DefaultValue $MissingValue
            SecureBoot                = GetFirstValue -FieldName 'SecureBoot' -Keys @('Etat du Secure Boot', 'SecureBootState', 'SecureBoot', 'Demarrage securise') -DefaultValue $MissingValue
        }

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
