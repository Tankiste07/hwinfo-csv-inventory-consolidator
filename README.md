# Consolidation Inventaire CSV vers Excel

## Fichiers
- `scripts/Consolidate-InventaireCsv.ps1` : script principal
- `scripts/Consolidate-InventaireCsv.Common.ps1` : fonctions communes de journalisation et normalisation
- `scripts/Consolidate-InventaireCsv.Data.ps1` : lecture des CSV et construction des objets PC
- `scripts/Consolidate-InventaireCsv.Description.ps1` : generation de descriptions a partir de templates
- `scripts/Consolidate-InventaireCsv.Export.ps1` : export Excel et nettoyage des sources
- `scripts/Install-Dependencies.ps1` : installe le module `ImportExcel`
- `scripts/Lancer-Inventaire.bat` : lance la consolidation avec verification des dependances

## Prerequis
- Windows PowerShell 5.1+
- Droit d'installer un module utilisateur (`CurrentUser`)

## Utilisation rapide
1. Ouvrir PowerShell dans ce dossier.
2. Installer la dependance :

```powershell
.\scripts\Install-Dependencies.ps1
```

3. Lancer la consolidation :

```powershell
.\scripts\Consolidate-InventaireCsv.ps1 -Folder "C:\Chemin\Vers\DossierCSV" -Output "C:\Chemin\Vers\InventaireFinal.xlsx"
```

4. Ou lancer avec les valeurs par defaut :

```powershell
.\scripts\Consolidate-InventaireCsv.ps1
```

5. Ou lancer en double-cliquant sur :

```text
Lancer-Inventaire.bat
```

Valeurs par defaut :
- `Folder` : `E:\Programs\HWInfo`
- `Output` : `E:\Programs\HWInfo\Inventaire_HWInfo.xlsx`

Champs manquants :
- Si une information n'existe pas dans le CSV source, la valeur `N/A` est ecrite dans Excel.

## Gestion des erreurs
- Les erreurs sont capturees avec des messages lisibles en console.
- Exemple si aucun CSV : `Pas de CSV dans : ...` (code de sortie `10`).
- Codes principaux :
- `10` : aucun CSV trouve
- `11` : dossier source introuvable
- `12` : module `ImportExcel` absent
- `13` : aucune donnee exploitable extraite

## Journaux (logs)
- Un fichier log est cree automatiquement a chaque execution.
- Dossier : `logs`
- Nom : `inventaire-YYYYMMDD-HHMMSS.log`
- Contenu : informations, avertissements et erreurs (`INFO`, `WARN`, `ERROR`).

## Resultat
- Un fichier Excel est genere avec :
- 1 ligne par ordinateur
- 1 colonne par information extraite
