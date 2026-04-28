<#
.SYNOPSIS
    Fonctions d'export Excel et de nettoyage des fichiers sources.
#>

function Export-InventoryToExcel {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Items,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $Items |
        Select-Object `
            @{Name='Marque'; Expression={ $_.NomMarqueOrdinateur }}, `
            @{Name='Modele'; Expression={ $_.ModeleOrdinateur }}, `
            @{Name='N de serie'; Expression={ $_.NumeroSerie }}, `
            @{Name='Processeur'; Expression={ $_.Processeur }}, `
            @{Name='Type de boitier'; Expression={ $_.TypeBoitier }}, `
            @{Name='Taille disque'; Expression={ $_.CapaciteSSD }}, `
            @{Name='Type de disque'; Expression={ $_.TypeDisque }}, `
            @{Name='Taille memoire'; Expression={ $_.MemoireTotale }}, `
            @{Name='Type DDR supporte'; Expression={ $_.TypeDDRSupporte }}, `
            @{Name='Taux d usure'; Expression={ $_.TauxUsure }}, `
            @{Name='Carte Graphique'; Expression={ if (-not [string]::IsNullOrWhiteSpace($_.CarteGraphique) -and $_.CarteGraphique -ne 'N/A') { $_.CarteGraphique } else { 'Aucune' } }}, `
            @{Name='Modele SSD'; Expression={ $_.ModeleSSD }}, `
            @{Name='Commentaire'; Expression={ '' }}, `
            @{Name='Grade'; Expression={ if ((Convert-WearToPercentValue $_.TauxUsure) -gt 38.0) { 'HS' } else { '' } } }, `
            @{Name='Valorisation'; Expression={ '' }} |
        Export-Excel -Path $OutputPath `
                     -WorksheetName 'Inventaire' `
                     -TableName 'InventairePC' `
                     -AutoSize `
                     -BoldTopRow `
                     -FreezeTopRow `
                     -ClearSheet

    $package = Open-ExcelPackage -Path $OutputPath
    try {
        $ws = $package.Workbook.Worksheets['Inventaire']
        if ($null -ne $ws -and $null -ne $ws.Dimension) {
            $gradeCol = $null
            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                if ($ws.Cells[1, $c].Text -eq 'Grade') {
                    $gradeCol = $c
                    break
                }
            }

            if ($gradeCol -ne $null -and $ws.Dimension.End.Row -ge 2) {
                $gradeRange = $ws.Cells[2, $gradeCol, $ws.Dimension.End.Row, $gradeCol].Address
                $validation = $ws.DataValidations.AddListValidation($gradeRange)
                $validation.ShowErrorMessage = $true
                $validation.ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::Stop
                $validation.ErrorTitle = 'Valeur invalide'
                $validation.Error = 'Choisissez A, B, C ou HS dans la liste.'
                $validation.ShowInputMessage = $true
                $validation.PromptTitle = 'Grade'
                $validation.Prompt = 'Sélectionnez A, B, C ou HS'
                $validation.Formula.Values.Add('A')
                $validation.Formula.Values.Add('B')
                $validation.Formula.Values.Add('C')
                $validation.Formula.Values.Add('HS')
            }
        }
        $package.Save()
    }
    finally {
        if ($null -ne $package) {
            $package.Dispose()
        }
    }
}

function Convert-WearToPercentValue {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 0.0
    }

    $text = $Value.Trim() -replace '%', ''
    $text = $text -replace ',', '.'
    $parsed = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }

    return 0.0
}

function Export-AnnouncementInventoryToExcel {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Items,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $Items |
        Select-Object `
            @{Name='Marque'; Expression={ $_.NomMarqueOrdinateur }}, `
            @{Name='Modele'; Expression={ $_.ModeleOrdinateur }}, `
            @{Name='N de serie'; Expression={ $_.NumeroSerie }}, `
            @{Name='Systeme'; Expression={ $_.SystemeOperateur }}, `
            @{Name='Processeur'; Expression={ $_.Processeur }}, `
            @{Name='Type de boitier'; Expression={ $_.TypeBoitier }}, `
            @{Name='Memoire totale'; Expression={ $_.MemoireTotale }}, `
            @{Name='Type DDR supporte'; Expression={ $_.TypeDDRSupporte }}, `
            @{Name='Disque'; Expression={
                if (-not [string]::IsNullOrWhiteSpace($_.CapaciteSSD) -and $_.CapaciteSSD -ne $script:MissingValue -and -not [string]::IsNullOrWhiteSpace($_.TypeDisque) -and $_.TypeDisque -ne $script:MissingValue) {
                    "$($_.CapaciteSSD) $($_.TypeDisque)"
                }
                elseif (-not [string]::IsNullOrWhiteSpace($_.CapaciteSSD) -and $_.CapaciteSSD -ne $script:MissingValue) {
                    $_.CapaciteSSD
                }
                elseif (-not [string]::IsNullOrWhiteSpace($_.TypeDisque) -and $_.TypeDisque -ne $script:MissingValue) {
                    $_.TypeDisque
                }
                else {
                    $script:MissingValue
                }
            }}, `
            @{Name='Modele SSD'; Expression={ $_.ModeleSSD }}, `
            @{Name='Taux d usure'; Expression={ $_.TauxUsure }}, `
            @{Name='Carte Graphique'; Expression={ if (-not [string]::IsNullOrWhiteSpace($_.CarteGraphique) -and $_.CarteGraphique -ne 'N/A') { $_.CarteGraphique } else { 'Aucune' } }}, `
            @{Name='Description'; Expression={ $_.DescriptionOrdi }} |
        Export-Excel -Path $OutputPath `
                     -WorksheetName 'Annonce' `
                     -TableName 'InventaireAnnonce' `
                     -TableStyle Light16 `
                     -AutoSize `
                     -BoldTopRow `
                     -FreezeTopRow `
                     -AutoFilter `
                     -ClearSheet

    $package = Open-ExcelPackage -Path $OutputPath
    try {
        $ws = $package.Workbook.Worksheets['Annonce']
        if ($null -ne $ws -and $null -ne $ws.Dimension) {
            $headerRange = $ws.Cells[1, 1, 1, $ws.Dimension.End.Column]
            $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(221,234,242))
            $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(31,73,125))
            $headerRange.Style.Font.Bold = $true

            $fullRange = $ws.Cells[1, 1, $ws.Dimension.End.Row, $ws.Dimension.End.Column]
            $fullRange.Style.WrapText = $false

            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                $ws.Column($c).AutoFit()
            }

            $descriptionCol = $null
            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                if ($ws.Cells[1, $c].Text -eq 'Description') {
                    $descriptionCol = $c
                    break
                }
            }

            if ($descriptionCol -ne $null) {
                $ws.Column($descriptionCol).Width = 50
                $ws.Column($descriptionCol).Style.WrapText = $true
            }
        }
        $package.Save()
    }
    finally {
        if ($null -ne $package) {
            $package.Dispose()
        }
    }
}

function Highlight-HighWearRows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$MissingValue,

        [Parameter(Mandatory = $false)]
        [double]$Threshold = 38.0
    )

    $package = Open-ExcelPackage -Path $OutputPath
    try {
        $ws = $package.Workbook.Worksheets['Inventaire']
        if ($null -ne $ws -and $null -ne $ws.Dimension) {
            $maxRow = $ws.Dimension.End.Row
            $maxCol = $ws.Dimension.End.Column

            $tauxCol = $null
            for ($c = 1; $c -le $maxCol; $c++) {
                $header = $ws.Cells[1, $c].Text
                $headerKey = Convert-ToLookupKey $header
                if ($headerKey -match 'taux.*usure|usure.*taux') {
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
                        if ($wear -gt $Threshold) {
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
}

function Remove-SourceCsvFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$CsvFiles
    )

    foreach ($csv in $CsvFiles) {
        try {
            Remove-Item -Path $csv.FullName -Force -ErrorAction Stop
        }
        catch {
            Write-AppWarning "Impossible de supprimer '$($csv.FullName)' : $($_.Exception.Message)"
        }
    }
}
