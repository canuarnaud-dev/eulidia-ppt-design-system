---
description: Mettre à jour le design system et le skill eulidia-powerpoint à partir d'une nouvelle galerie des composants (.pptx)
argument-hint: <chemin de la nouvelle galerie .pptx>
---

Nouvelle galerie : $ARGUMENTS

Tu mets à jour le kit à partir de cette nouvelle version de `galerie-composants-eulidia.pptx`, en suivant `MAINTENANCE.md` (procédure, décisions de la section 3, pièges de la section 4). Arnaud veut le moins d'étapes possible : tu fais tout le mécanique, tu ne lui demandes que les validations.

Python : toujours `.venv\Scripts\python.exe` (Python 3.12, avec pyoxipng). Si `.venv` n'existe pas, le créer : `uv venv --python 3.12 .venv` puis `uv pip install --python .venv\Scripts\python.exe -r requirements.txt`.

## 1. Préparer
- Si aucun chemin n'est donné ci-dessus, demander à Arnaud de glisser le fichier .pptx dans le chat.
- Vérifier que l'arbre git est propre, puis créer une branche `maj-galerie-AAAA-MM-JJ` depuis `main` à jour (`git pull`).
- Vérifier que `design-system/` correspond à l'artefact publié (étape 1 de MAINTENANCE.md) : lire l'artefact avec l'outil Artifact (`action: "read"` / `list` avec `scope: "files"`) et comparer. S'il diffère, reprendre la version publiée avant toute modification et le signaler.

## 2. Analyser et faire valider
- `.venv\Scripts\python.exe scripts\mettre_a_jour.py "<galerie>" --analyse`, puis lire `build\maj-galerie\rapport.json`.
- Présenter à Arnaud un résumé clair et court : composants ajoutés / retirés / modifiés (visuel ou notes), déplacements, masque, plaquette, socle à reconstruire, fiches à revoir. Pour un visuel modifié, rendre l'avant/après : la vignette actuelle `skill\assets\exemples\slide-NN.jpg` et la nouvelle (exportée à l'étape 3) ou la slide elle-même.
- Signaler tout ce qui contredit une décision de la section 3 (plaquette ≠ 7 slides, dispositions du masque ≠ les 5 attendues, etc.) : ne pas trancher seul.
- Demander la confirmation d'Arnaud. Sans « oui », s'arrêter.

## 3. Appliquer la partie mécanique
- `.venv\Scripts\python.exe scripts\mettre_a_jour.py "<galerie>" --oui` : copie la galerie dans `reference\`, reconstruit le socle si besoin, exporte les vignettes dans `build\maj-galerie\vignettes\` via PowerPoint, puis génère et teste le skill s'il n'y a rien à rédiger.

## 4. Rédiger ce qui doit l'être (liste « a_revoir » du rapport)
- **Renumérotation** : renommer les vignettes de `skill\assets\exemples\` (en deux temps pour éviter les collisions) et mettre à jour chaque mention `slide-NN.jpg` (fiches, `design-system\assets\Exemples\README.md`, `skill\SKILL.md`, table de `scripts\generer_skill.py`).
- **Vignettes** : copier les nouvelles depuis `build\maj-galerie\vignettes\` vers `skill\assets\exemples\` ; supprimer celles des composants retirés.
- **Composant nouveau ou modifié** : lire ses notes dans le rapport (`texte_notes`) et regarder la vignette ; écrire ou ajuster `design-system\components\<Nom>\README.md` au format de MAINTENANCE.md (1re ligne = description, en-tête Groupe · Disposition · Exemple, puis règles) et `preview.html` (copier un existant). Ajouter la ligne au catalogue du README du design system et, si besoin, aux ordres types.
- **Composant retiré** : proposer la suppression de la fiche et de ses mentions ; attendre l'accord d'Arnaud.
- **Fondation, section ou masque modifiés** : mettre à jour `design-system\README.md`, `tokens.json` ou `design-system\masque\README.md` à partir des notes.
- Garder `skill\SKILL.md` cohérent avec le README (une règle du README se retrouve dans le skill).
- Montrer à Arnaud le diff des fiches avant de continuer.

## 5. Générer et tester
- `.venv\Scripts\python.exe scripts\mettre_a_jour.py --generer`. En cas de « phrase introuvable », mettre à jour la table de `generer_skill.py` (jamais contourner).
- Pour un changement visuel, construire une présentation de test avec le skill généré et en vérifier le rendu.

## 6. Publier, livrer, consigner
- Publier le design system (étape 6 de MAINTENANCE.md) : fichiers texte modifiés, vignettes en assets du groupe *Exemples* (ids repris dans les `preview.html`), `design-system.json` en dernier. Si les outils d'artefacts ne le permettent pas depuis Claude Code, le dire et indiquer les fichiers à publier depuis Cowork.
- Indiquer à Arnaud le zip à réimporter : `build\eulidia-powerpoint.zip`.
- Ajouter une ligne au *Journal* de MAINTENANCE.md, committer (message en français : quoi et pourquoi), puis proposer le merge sur `main` et le push : les faire seulement après son accord.
