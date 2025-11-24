# Google Drive Duplicator

Ce projet permet de dupliquer une arborescence complète d'un dossier Google Drive (partagé ou non) vers votre propre Drive.

## Fonctionnalités

*   **Copie Récursive** : Copie tous les fichiers et sous-dossiers.
*   **Gestion des Limites** : Gère la limite de temps de 6 minutes de Google Apps Script. Si la copie n'est pas finie, vous pouvez la relancer et elle reprendra exactement où elle s'est arrêtée.
*   **Vérification** : Compare le nombre de fichiers entre la source et la destination pour s'assurer de l'intégrité.
*   **Interface Google Sheets** : Pilotage facile via un tableau.

## Installation

### Option A : Copier-Coller (Simple)

1.  Créez un nouveau **Google Sheet**.
2.  Allez dans **Extensions** > **Apps Script**.
3.  Supprimez le code existant dans `Code.gs`.
4.  Copiez le contenu du fichier `src/Code.js` de ce dépôt et collez-le dans l'éditeur.
5.  Sauvegardez.
6.  Rechargez votre Google Sheet. Un menu "Drive Duplicator" apparaîtra.

### Option B : Utilisation de CLASP (Avancé)

Si vous avez Node.js installé :

1.  Installez clasp : `npm install -g @google/clasp`
2.  Connectez-vous : `clasp login`
3.  Créez un sheet : `clasp create --type sheets --title "Drive Duplicator"` (ou clonez un projet existant).
4.  Poussez le code : `clasp push`

## Configuration du Google Sheet

Le script s'attend à trouver les colonnes suivantes (l'ordre n'est pas strict, mais c'est mieux de suivre cet en-tête) :

| Ligne 1 | A | B | C | D |
| :--- | :--- | :--- | :--- | :--- |
| **En-têtes** | **Source Folder ID** | **Status** | **Destination URL** | **Verification** |

*   **Source Folder ID** : L'ID du dossier que vous voulez copier (la partie à la fin de l'URL du dossier).
*   **Status** : Laissez vide au début. Le script mettra "Pending", "Processing", "Done", ou "Error".

## Utilisation

1.  Remplissez l'ID du dossier source dans la colonne A.
2.  Allez dans le menu **Drive Duplicator** > **Start Copy**.
3.  Si le script s'arrête (limite de temps), le statut restera "Processing" (ou indiquera "Time Limit"). Relancez simplement **Start Copy** pour continuer.
4.  Une fois terminé, le statut sera "Done".
5.  Pour vérifier, cliquez sur **Drive Duplicator** > **Verify Folder**.
