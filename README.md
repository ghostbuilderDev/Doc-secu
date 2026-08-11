# Docu Chantier

Application web mobile de génération de documents Word à partir de trames `.docx`.

## Ce qui est inclus

- Les trames **ISF**, **CSF 2026** et **PPSPS** fournies, sans modification.
- Une fiche dossier réutilisable (opération, période, entreprise, horaires, contacts…).
- Un éditeur de toutes les zones surlignées de chaque trame, y compris les cases `X` et les pieds de page.
- Une génération `.docx` locale qui conserve la mise en page, les tableaux, les images et les annexes Word.
- L’ajout de nouvelles trames Word depuis le téléphone ou l’ordinateur.
- Un mode PWA : une fois ouvert une première fois, l’application reste utilisable hors connexion.

## Mise en ligne avec GitHub Pages

1. Créer un dépôt GitHub, par exemple `docu-chantier`.
2. Déposer **le contenu** du dossier `app-documents` à la racine du dépôt.
3. Dans GitHub : `Settings` → `Pages` → `Deploy from a branch` → sélectionner `main` et le dossier `/ (root)`.
4. Ouvrir l’adresse GitHub Pages obtenue depuis Chrome ou Edge sur le téléphone, puis l’ajouter à l’écran d’accueil si besoin.

Chaque mise à jour envoyée sur GitHub conserve le même lien. Pour récupérer une nouvelle version, fermer puis rouvrir l’application lorsqu’une connexion est disponible.

## Utilisation

1. Ouvrir une trame dans **Documents**.
2. Renseigner la **Fiche dossier**.
3. Dans **Champs à remplir**, préremplir les zones reconnues puis vérifier les autres informations.
4. Dans **Générer**, sélectionner la ou les trames et télécharger le fichier Word final.

Lorsque plusieurs trames sont générées en même temps, elles sont regroupées dans une archive `.zip` pour éviter les téléchargements bloqués par le téléphone.

Les données restent dans le navigateur de l’appareil : aucune donnée de chantier n’est transmise par l’application.

## Licence de la bibliothèque intégrée

Le moteur ZIP `JSZip 3.10.1` est fourni localement dans `assets/jszip.min.js`. Sa licence est disponible dans `assets/JSZIP-LICENSE.md`.
