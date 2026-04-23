#requires -Version 5.1
<#!
.SYNOPSIS
    Installe la dependance necessaire au script de consolidation CSV -> Excel.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Prepare package source to avoid interactive prompts on first install.
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installation du module ImportExcel..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
}
else {
    Write-Host "Le module ImportExcel est deja installe."
}

Write-Host "Dependances OK."
