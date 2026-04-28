@echo off
setlocal EnableDelayedExpansion

REM Se place dans le dossier du .bat
cd /d "%~dp0"

set "OUTPUT_DIR=E:\Programs\HWInfo"
REM set "OUTPUT_DIR =C:\Users\33695\Documents\ScriptBrokerInfo\CSV_Test"
set "OUTPUT_FILE=Inventaire_HWInfo.xlsx"
set "LAST_XLSX="

echo Verification des dependances...
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Get-Module -ListAvailable -Name ImportExcel) { exit 0 } else { exit 1 }"
if errorlevel 1 (
    echo ImportExcel non detecte. Installation en cours...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install-Dependencies.ps1"
    if not "%ERRORLEVEL%"=="0" (
        echo.
        echo Erreur: impossible d'installer les dependances.
        echo Appuyez sur une touche pour fermer...
        pause >nul
        exit /b 1
    )
)

REM Lance le script PowerShell avec une policy temporaire pour cette execution
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Consolidate-InventaireCsv.ps1"
set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
    echo.
    if "%EXITCODE%"=="10" (
        echo Pas de CSV trouve dans %OUTPUT_DIR%.
    ) else if "%EXITCODE%"=="11" (
        echo Le dossier source est introuvable : %OUTPUT_DIR%.
    ) else if "%EXITCODE%"=="12" (
        echo Dependance ImportExcel absente ou inutilisable.
    ) else (
        echo Une erreur est survenue pendant la consolidation ^(code %EXITCODE%^).
    )
) else (
    if exist "%OUTPUT_DIR%\%OUTPUT_FILE%" (
        set "LAST_XLSX=%OUTPUT_DIR%\%OUTPUT_FILE%"
    )

    echo.
    echo Execution terminee avec succes.

    if defined LAST_XLSX (
        echo Ouverture du fichier : !LAST_XLSX!
        start "" "!LAST_XLSX!"
    ) else (
        echo Fichier Excel non trouve dans %OUTPUT_DIR%.
    )
)

echo.
echo Appuyez sur une touche pour fermer...
pause >nul
exit /b %EXITCODE%
