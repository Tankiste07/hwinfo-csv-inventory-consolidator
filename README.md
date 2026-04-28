# Consolidation Inventaire CSV vers Excel

## Fichiers
- `scripts/Consolidate-InventaireCsv.ps1` : script principal
- `scripts/Consolidate-InventaireCsv.Common.ps1` : fonctions communes de journalisation et normalisation
- `scripts/Consolidate-InventaireCsv.Data.ps1` : lecture des CSV et construction des objets PC
- `scripts/Consolidate-InventaireCsv.Description.ps1` : génération de descriptions à partir de templates
- `scripts/Consolidate-InventaireCsv.Export.ps1` : export Excel et nettoyage des sources
- `scripts/Install-Dependencies.ps1` : installe le module `ImportExcel`
- `Lancer-Inventaire.bat` : lance la consolidation avec vérification des dépendances

## Prérequis
- Windows PowerShell 5.1+
- Droits pour installer un module utilisateur (`CurrentUser`)

## Utilisation rapide
1. Ouvrir PowerShell dans le dossier racine du projet.
2. Installer la dépendance :

```powershell
.\scripts\Install-Dependencies.ps1
```

3. Lancer la consolidation :

```powershell
.\scripts\Consolidate-InventaireCsv.ps1 -Folder "C:\Chemin\Vers\DossierCSV" -Output "C:\Chemin\Vers\InventaireFinal.xlsx"
```

4. Lancer avec les valeurs par défaut :

```powershell
.\scripts\Consolidate-InventaireCsv.ps1
```

5. Lancer avec le batch :

```text
Lancer-Inventaire.bat
```

Valeurs par défaut :
- `Folder` : `E:\Programs\HWInfo`
- `Output` : `E:\Programs\HWInfo\Inventaire_HWInfo.xlsx`

## Comportement du traitement
- Un second fichier `Inventaire_Annonce.xlsx` est généré automatiquement dans le même dossier que la sortie principale.
- Le champ `Carte Graphique` est présent dans les deux exports et vaut `Aucune` s'il n'existe pas.
- Dans le premier fichier, la colonne `Grade` contient un menu déroulant (`A`, `B`, `C`, `HS`).
- Si le `Taux d usure` dépasse `38 %`, le `Grade` est rempli automatiquement avec `HS` dans le premier fichier.
- Les champs manquants sont remplis avec `N/A`.

## Gestion des erreurs
- Les erreurs sont capturées et affichées en console.
- Si aucun CSV n'est trouvé : `Pas de CSV dans : ...` (code de sortie `10`).
- Codes principaux :
  - `10` : aucun CSV trouvé
  - `11` : dossier source introuvable
  - `12` : module `ImportExcel` absent
  - `13` : aucune donnée exploitable extraite

## Journaux (logs)
- Un fichier log est créé automatiquement à chaque exécution.
- Dossier : `logs`
- Nom : `inventaire-YYYYMMDD-HHMMSS.log`
- Contenu : informations, avertissements et erreurs (`INFO`, `WARN`, `ERROR`).

## Résultat
- Deux fichiers Excel sont générés :
  - `Inventaire_HWInfo.xlsx`
  - `Inventaire_Annonce.xlsx`
- 1 ligne par ordinateur
- 1 colonne par information extraite
