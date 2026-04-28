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
            @{Name='Carte Graphique'; Expression={ $_.CarteGraphique }}, `
            @{Name='Modele SSD'; Expression={ $_.ModeleSSD }}, `
            @{Name='Commentaire'; Expression={ '' }}, `
            @{Name='Grade'; Expression={ '' }}, `
            @{Name='Valorisation'; Expression={ '' }} |
        Export-Excel -Path $OutputPath `
                     -WorksheetName 'Inventaire' `
                     -TableName 'InventairePC' `
                     -AutoSize `
                     -BoldTopRow `
                     -FreezeTopRow `
                     -ClearSheet
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
