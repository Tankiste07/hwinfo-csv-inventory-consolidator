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
    [ValidateNotNullOrEmpty()]
    [string]$Folder = 'E:\Programs\HWInfo',
    [string]$Output,
    [string]$DescriptionTemplateFolder = 'E:\Programs\HWInfo\SquelletteDescription'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MissingValue = 'N/A'

$script:LogDirectory = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'logs'
$script:LogFile = Join-Path -Path $script:LogDirectory -ChildPath ("inventaire-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
if (-not (Test-Path -Path $script:LogDirectory -PathType Container)) {
    New-Item -Path $script:LogDirectory -ItemType Directory -Force | Out-Null
}

. "$PSScriptRoot\Consolidate-InventaireCsv.Common.ps1"
. "$PSScriptRoot\Consolidate-InventaireCsv.Data.ps1"
. "$PSScriptRoot\Consolidate-InventaireCsv.Description.ps1"
. "$PSScriptRoot\Consolidate-InventaireCsv.Export.ps1"

Write-Log -Level 'INFO' -Message "Demarrage du script. Dossier source='$Folder', sortie='$Output'"

trap {
    $trapMessage = $_.Exception.Message
    Write-Host "[ERREUR] $trapMessage"
    Write-Log -Level 'ERROR' -Message $trapMessage
    exit 99
}

if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Join-Path -Path $Folder -ChildPath 'Inventaire_HWInfo.xlsx'
}

if (Test-Path -Path $Output) {
    try {
        Remove-Item -Path $Output -Force -ErrorAction Stop
    }
    catch {
        Exit-WithMessage -Message "Impossible de supprimer l'ancien fichier Excel '$Output' (fichier ouvert ?). Fermez-le et relancez." -Code 20
    }
}

$AnnouncementOutput = Join-Path -Path (Split-Path -Path $Output -Parent) -ChildPath 'Inventaire_Annonce.xlsx'
if (Test-Path -Path $AnnouncementOutput) {
    try {
        Remove-Item -Path $AnnouncementOutput -Force -ErrorAction Stop
    }
    catch {
        Exit-WithMessage -Message "Impossible de supprimer l'ancien fichier Excel '$AnnouncementOutput' (fichier ouvert ?). Fermez-le et relancez." -Code 21
    }
}

if (-not (Test-Path -Path $Folder -PathType Container)) {
    Exit-WithMessage -Message "Le dossier source n'existe pas : $Folder" -Code 11
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Exit-WithMessage -Message "Le module 'ImportExcel' est introuvable. Lance d'abord .\Install-Dependencies.ps1" -Code 12
}
Import-Module ImportExcel -ErrorAction Stop

function Render-ProgressBar {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Current,

        [Parameter(Mandatory = $true)]
        [int]$Total,

        [string]$Label = 'Traitement',
        [int]$Width = 20
    )

    if ($Total -le 0) {
        return
    }

    $percent = [math]::Round(($Current / $Total) * 100)
    if ($percent -gt 100) { $percent = 100 }

    $filled = [int]([math]::Floor($Width * $Current / $Total))
    if ($filled -lt 0) { $filled = 0 }
    if ($filled -gt $Width) { $filled = $Width }

    $empty = $Width - $filled
    $fill = if ($filled -gt 0) { -join (1..$filled | ForEach-Object { 'o' }) } else { '' }
    $dash = if ($empty -gt 0) { -join (1..$empty | ForEach-Object { '-' }) } else { '' }

    $bar = "|$fill$dash|"
    $display = "{0} {1} {2}`%" -f $Label, $bar, $percent
    $padding = ' ' * [math]::Max(0, 80 - $display.Length)

    Write-Host -NoNewline "`r$display$padding"
    if ($Current -eq $Total) {
        Write-Host
    }
}

$csvFiles = Get-ChildItem -Path $Folder -Filter '*.csv' -File
if (-not $csvFiles) {
    Exit-WithMessage -Message "Pas de CSV dans : $Folder" -Code 10
}

$totalSteps = $csvFiles.Count + 4
$currentStep = 0
$results = New-Object System.Collections.Generic.List[object]
$descriptionTemplates = Load-DescriptionTemplates -DescriptionTemplateFolder $DescriptionTemplateFolder -MissingValue $MissingValue

foreach ($file in $csvFiles) {
    $currentStep++
    try {
        $script:CurrentFile = $file.Name
        $script:CurrentIndex = Read-HwInfoCsvIndex -FilePath $file.FullName

        $pcObject = Build-ComputerObject -DescriptionTemplates $descriptionTemplates -MissingValue $MissingValue
        [void]$results.Add($pcObject)
    }
    catch {
        Write-AppWarning "Fichier ignore '$($file.Name)' : $($_.Exception.Message)"
        continue
    }
    Render-ProgressBar -Current $currentStep -Total $totalSteps -Label "Traitement"
}

if ($results.Count -eq 0) {
    Exit-WithMessage -Message "Aucune donnee exploitable n'a ete extraite des CSV." -Code 13
}

Export-InventoryToExcel -Items $results -OutputPath $Output
$currentStep++
Render-ProgressBar -Current $currentStep -Total $totalSteps -Label "Export inventaire"

Export-AnnouncementInventoryToExcel -Items $results -OutputPath $AnnouncementOutput
$currentStep++
Render-ProgressBar -Current $currentStep -Total $totalSteps -Label "Export annonce"

Highlight-HighWearRows -OutputPath $Output -MissingValue $MissingValue -Threshold 38.0
$currentStep++
Render-ProgressBar -Current $currentStep -Total $totalSteps -Label "Mise en surbrillance"

Remove-SourceCsvFiles -CsvFiles $csvFiles
$currentStep++
Render-ProgressBar -Current $currentStep -Total $totalSteps -Label "Nettoyage"

Write-AppInfo "Export termine : $Output"
Write-AppInfo "Export annonce termine : $AnnouncementOutput"
Write-AppInfo "Nombre de postes exportes : $($results.Count)"
Write-AppInfo "Log : $script:LogFile"
exit 0
