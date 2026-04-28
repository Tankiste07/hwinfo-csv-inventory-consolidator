<#
.SYNOPSIS
    Fonctions de lecture des CSV et de construction des objets PC.
#>

function Load-DescriptionTemplates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DescriptionTemplateFolder,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

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

    return $descriptionTemplates
}

function Read-HwInfoCsvIndex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $lines = Get-Content -Path $FilePath
    $index = @{}

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

        $kn = Normalize-Key $k
        if ($kn -in @('key', 'cle') -and (Normalize-Key $v) -in @('value', 'valeur')) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($kn)) {
            continue
        }

        if (-not $index.ContainsKey($kn)) {
            $index[$kn] = New-Object System.Collections.Generic.List[string]
        }

        [void]$index[$kn].Add($v)
    }

    return $index
}

function Get-SupportedDdrType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    $ddr4Value = GetFirstValueNoWarning -Keys @('DDR4', 'Support DDR4') -DefaultValue $MissingValue
    $ddr3lValue = GetFirstValueNoWarning -Keys @('DDR3L', 'Support DDR3L') -DefaultValue $MissingValue
    $ddr3Value = GetFirstValueNoWarning -Keys @('DDR3', 'Support DDR3') -DefaultValue $MissingValue

    $supportedDdrTypes = New-Object System.Collections.Generic.List[string]
    if (Test-IsSupportedValue $ddr4Value) { [void]$supportedDdrTypes.Add('DDR4') }
    if (Test-IsSupportedValue $ddr3lValue) { [void]$supportedDdrTypes.Add('DDR3L') }
    if (Test-IsSupportedValue $ddr3Value) { [void]$supportedDdrTypes.Add('DDR3') }

    if ($supportedDdrTypes.Count -gt 0) {
        return ($supportedDdrTypes -join ', ')
    }

    if ($ddr4Value -eq $MissingValue -and $ddr3lValue -eq $MissingValue -and $ddr3Value -eq $MissingValue) {
        return $MissingValue
    }

    return 'Aucun'
}

function Build-ComputerObject {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$DescriptionTemplates,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue
    )

    $nomMarqueOrdinateurRaw = GetFirstValue -FieldName 'NomMarqueOrdinateur' -Keys @("Nom de marque de l'ordinateur", 'Computer Brand Name', 'Brand Name', 'System Brand') -FallbackContains @('marque', 'ordinateur') -DefaultValue $MissingValue
    $modeleOrdinateurRaw = GetFirstValueNoWarning -Keys @('Modele', 'Model', 'Product Name', 'Model Name', 'System Model', 'Computer Model', 'Nom du modele', 'Nom du modèle') -DefaultValue $null
    $typeDdrSupporte = Get-SupportedDdrType -MissingValue $MissingValue
    $diskTypeRaw = GetFirstValue -FieldName 'TypeDisque' -Keys @('Type de disque', 'Drive Type', 'Media Type', 'Disk Type', 'Storage Type', 'Type du disque') -DefaultValue $MissingValue
    $rawCapacity = GetFirstValue -FieldName 'CapaciteSSD' -Keys @('Capacite du SSD', 'SSDCapacity', 'DiskSize', 'StorageCapacity', 'Capacite du lecteur') -DefaultValue $MissingValue
    $rawCpu = GetFirstValue -FieldName 'NomProcesseur' -Keys @('Nom du processeur', 'ProcessorName', 'CPUName', 'CPU') -DefaultValue $MissingValue
    $rawModeleSSD = GetFirstValue -FieldName 'ModeleSSD' -Keys @('Modele du SSD', 'SSDModel', 'DiskModel', 'StorageModel', 'Modele de lecteur') -DefaultValue $MissingValue

    $brandModel = Split-BrandAndModelFromName -RawName $nomMarqueOrdinateurRaw
    $nomMarqueOrdinateur = if (-not [string]::IsNullOrWhiteSpace($brandModel.Brand)) { $brandModel.Brand } else { $nomMarqueOrdinateurRaw }
    $modeleOrdinateur = if ([string]::IsNullOrWhiteSpace($modeleOrdinateurRaw) -or $modeleOrdinateurRaw -eq $MissingValue) { $brandModel.Model } else { $modeleOrdinateurRaw }

    $pcObject = [PSCustomObject]@{
        SystemeOperateur          = Normalize-OperatingSystem -Value (GetFirstValue -FieldName 'SystemeOperateur' -Keys @('Systeme operateur', 'OperatingSystem', 'OS', 'Caption') -DefaultValue $MissingValue) -MissingValue $MissingValue
        NomMarqueOrdinateur       = $nomMarqueOrdinateur
        ModeleOrdinateur          = if ([string]::IsNullOrWhiteSpace($modeleOrdinateur)) { $MissingValue } else { $modeleOrdinateur }
        Processeur                = $rawCpu
        TypeBoitier               = GetFirstValue -FieldName 'TypeBoitier' -Keys @('Type de boitier', 'Type de boîtier', 'Chassis Type', 'Type de chassis') -DefaultValue $MissingValue
        NumeroSerie               = GetFirstValue -FieldName 'NumeroSerie' -Keys @('Numero de serie', 'Numéro de série', 'Serial Number', 'SerialNumber', 'System Serial Number', 'BIOS serial number') -FallbackContains @('serial') -DefaultValue $MissingValue
        MemoireTotale             = GetFirstValue -FieldName 'MemoireTotale' -Keys @('Taille totale de la memoire', 'TotalMemory', 'MemoryTotal', 'RAM', 'Memoire physique totale') -DefaultValue $MissingValue
        TypeDDRSupporte           = $typeDdrSupporte
        TauxUsure                 = GetFirstValue -FieldName 'TauxUsure' -Keys @('Taux d usure', 'Taux d usure de la batterie', "Taux d'usure", 'BatteryWearLevel', 'BatteryWear', 'Usure batterie') -FallbackContains @('taux', 'usure') -DefaultValue $MissingValue
        CarteGraphique            = GetFirstValue -FieldName 'CarteGraphique' -Keys @('Jeu de puces graphiques', 'Graphics Chipset', 'Graphic Chipset', 'GPU Chipset', 'Nom de la puce graphique') -FallbackContains @('puces', 'graphi') -DefaultValue $MissingValue
        ModeleSSD                 = $rawModeleSSD
        CapaciteSSD               = Normalize-StorageCapacity -Value $rawCapacity -MissingValue $MissingValue
        TypeDisque                = Normalize-DiskType -RawType $diskTypeRaw -Capacity $rawCapacity -ModelSSD $rawModeleSSD -MissingValue $MissingValue
    }

    $descriptionOrdi = $MissingValue
    $matchedTemplatePath = Select-DescriptionTemplatePath -TemplateMap $DescriptionTemplates -Brand $nomMarqueOrdinateur -TypeBoitier ([string]$pcObject.TypeBoitier)
    if ($null -ne $matchedTemplatePath) {
        try {
            $templateContent = Get-Content -Path $matchedTemplatePath -Raw -Encoding UTF8
            $descriptionOrdi = Get-DescriptionFromTemplate -TemplateText $templateContent -ComputerData $pcObject -MissingValue $MissingValue
        }
        catch {
            Write-AppWarning "[$script:CurrentFile] Impossible de lire le fichier de description '$matchedTemplatePath' : $($_.Exception.Message)"
        }
    }

    $pcObject | Add-Member -NotePropertyName 'DescriptionOrdi' -NotePropertyValue $descriptionOrdi
    return $pcObject
}
