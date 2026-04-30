<#
.SYNOPSIS
    Fonctions communes de journalisation et de normalisation.
#>

function Write-Log {
    param(
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level,

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

function GetValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return $null
    }

    $normalizedKey = Normalize-Key $Key
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

function Find-FirstIndexedValue {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FallbackContains
    )

    if (-not $FallbackContains -or $FallbackContains.Count -eq 0) {
        return $null
    }

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

        if (-not $allMatch) {
            continue
        }

        $fallbackValue = $script:CurrentIndex[$indexKey] |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 1

        if (-not [string]::IsNullOrWhiteSpace($fallbackValue)) {
            return $fallbackValue
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
        $fallbackValue = Find-FirstIndexedValue -FallbackContains $FallbackContains
        if (-not [string]::IsNullOrWhiteSpace($fallbackValue)) {
            return $fallbackValue
        }
    }

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
        return Find-FirstIndexedValue -FallbackContains $FallbackContains
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

function Get-StorageCapacityFromModel {
    param(
        [AllowNull()]
        [string]$Model
    )

    if ([string]::IsNullOrWhiteSpace($Model)) {
        return $null
    }

    $normalized = Normalize-Key $Model
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $match = [regex]::Match($normalized, '([0-9]+(?:[\.,][0-9]+)?)\s*(gb|go|tb|to)\b', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        $num = $match.Groups[1].Value -replace '\.0+$', ''
        $unit = $match.Groups[2].Value.ToLowerInvariant()
        switch ($unit) {
            'gb' { $unit = 'Go' }
            'go' { $unit = 'Go' }
            'tb' { $unit = 'To' }
            'to' { $unit = 'To' }
        }
        return ('{0} {1}' -f $num, $unit)
    }

    return $null
}

function Get-StorageTypeFromModel {
    param(
        [AllowNull()]
        [string]$Model
    )

    if ([string]::IsNullOrWhiteSpace($Model)) {
        return $null
    }

    $normalized = Normalize-Key $Model
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $lower = $normalized.ToLowerInvariant()

    $ssdPatterns = @(
        'ssd',
        'solid state',
        'nvme',
        'm\.2',
        '\bm2\b',
        'pci[e]?e',
        '\bmz7',
        '\bmzv',
        '\bmzn',
        '\bkxg',
        '\bkbg',
        '\bhfs',
        'liteon cv',
        'liteoncv',
        '\bpm871a?',
        '\bpm991',
        '\bpm9a1',
        'samsung mz7',
        'samsung mzv',
        'samsung mzn',
        'samsung mz',
        'samsung pm9',
        'hynix',
        'intel',
        'crucial',
        'micron',
        'kingston',
        'sandisk',
        'lexar'
    )

    $hddPatterns = @(
        'hdd',
        'hard drive',
        'hard disk',
        '5400\s*rpm',
        '7200\s*rpm',
        '5400rpm',
        '7200rpm',
        'rotational',
        'rotationnel',
        '\bst',
        '\bwd',
        '\bhts',
        '\bmk',
        '\bhgst',
        'seagate',
        'western digital',
        'hitachi',
        'toshiba'
    )

    foreach ($pattern in $ssdPatterns) {
        if ($lower -match $pattern) {
            return 'SSD'
        }
    }

    foreach ($pattern in $hddPatterns) {
        if ($lower -match $pattern) {
            return 'HDD'
        }
    }

    return $null
}

function Normalize-DiskType {
    param(
        [AllowNull()]
        [string]$RawType,

        [AllowNull()]
        [string]$Capacity,

        [AllowNull()]
        [string]$ModelSSD,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    $storageType = Get-StorageTypeFromModel -Model $ModelSSD
    if ($storageType) {
        return $storageType
    }

    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($RawType) -and $RawType -ne $MissingValue) { $parts += $RawType }
    if (-not [string]::IsNullOrWhiteSpace($Capacity) -and $Capacity -ne $MissingValue) { $parts += $Capacity }
    if (-not [string]::IsNullOrWhiteSpace($ModelSSD) -and $ModelSSD -ne $MissingValue) { $parts += $ModelSSD }

    $source = ($parts -join ' ').ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($source)) {
            if ($source -match 'ssd|nvme|pci[e]?e|pcie|solid state|flash|m2\b|sata.*ssd') {
        }

        if ($source -match 'hdd|hard disk|hard drive|disk drive|disque dur|seagate|western digital|wd|toshiba|hitachi|mq') {
            return 'HDD'
        }

        if ($source -match 'sata') {
            if ($source -match 'ssd') {
                return 'SSD'
            }
            return 'HDD'
        }
    }

    return 'Inconnu'
}

function Split-BrandAndModelFromName {
    param(
        [AllowNull()]
        [string]$RawName
    )

    if ([string]::IsNullOrWhiteSpace($RawName)) {
        return @{ Brand = $null; Model = $null }
    }

    $clean = $RawName.Trim()
    $knownBrands = @(
        'dell', 'hp', 'hewlett packard', 'lenovo', 'asus', 'acer', 'toshiba', 'msi',
        'apple', 'samsung', 'lg', 'huawei', 'sony', 'fujitsu', 'panasonic', 'zte', 'gigabyte'
    )

    foreach ($brand in $knownBrands) {
        if ((Normalize-Key $clean) -like "$brand *" -or (Normalize-Key $clean) -eq $brand) {
            $pattern = "(?i)^\s*($brand)\s+(.+)$"
            $match = [regex]::Match($clean, $pattern)
            if ($match.Success) {
                return @{ Brand = $match.Groups[1].Value.Trim(); Model = $match.Groups[2].Value.Trim() }
            }

            return @{ Brand = $brand; Model = $null }
        }
    }

    $parts = $clean -split '\s+'
    if ($parts.Count -gt 1) {
        return @{ Brand = $parts[0]; Model = ($parts[1..($parts.Count - 1)] -join ' ') }
    }

    return @{ Brand = $clean; Model = $null }
}

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
