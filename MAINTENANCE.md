# Maintenance du design system « Eulidia 2026 · PowerPoint » et du skill `eulidia-powerpoint`

Mémoire de travail pour faire évoluer les deux sans refaire les erreurs passées. À lire en entier avant toute modification, et à compléter à la fin de chaque session de maintenance (section *Journal*, en plus du message de commit).

## 1. Ce qui existe et qui fait foi

| Élément | Où | Rôle |
| --- | --- | --- |
| PowerPoint de référence | `reference/galerie-composants-eulidia.pptx` (remplacer ce fichier quand la galerie évolue, dans un commit dédié) | Source du masque, de la plaquette et des exemples. Ses notes de slide portent : Composant / Groupe / Description / Disposition. |
| Design system | artefact « Eulidia 2026 · PowerPoint » (type Design System), https://claude.ai/artifact/EfURVfbSGkTwtbcbYwvuw1, slug `6ea5f727-2ccd-4264-992c-5a1d6c288d94`, partagé (`read: users`) | Référence lisible par les humains. Fichiers texte sous `project/` + polices + assets (images). |
| Copie de travail du design system | `design-system/` de ce dépôt | Copie **exacte** des fichiers texte de l'artefact (README, construction.md, tokens.json, fiches, previews). C'est elle qu'on modifie, puis on publie. |
| Skill de l'organisation | `eulidia-powerpoint` (zip importé dans les skills de l'orga par Arnaud) | Copie embarquée du design system + socle + module python + assets. **Toujours régénéré** par `scripts/generer_skill.py`, jamais modifié à la main. |
| Source du skill | `skill/SKILL.md` (rédigé à la main) + `skill/assets/` + `pptx/eulidia-socle-2026.pptx` | Ce qui n'est pas dérivé du design system. |
| Hors périmètre | skill perso `presentation-eulidia` (méthode d'Arnaud en 5 phases, appelée seulement sur demande explicite ; elle appelle `eulidia-powerpoint`) | Ne pas le modifier depuis ce dépôt. Signaler seulement si une évolution le rend incohérent (ex. nouvelle fonction du module, règle de plaquette). |

Chaîne : PowerPoint de référence → (socle) → design system (artefact) → skill généré → utilisateurs.

## 2. Procédure de mise à jour

**Nouvelle version de la galerie** : dans Claude Code, `/maj-galerie <fichier.pptx>` (`.claude/commands/maj-galerie.md`) déroule toute la procédure ci-dessous. Sa partie mécanique est `scripts/mettre_a_jour.py` : comparaison slide par slide avec `reference/` (rapport `build/maj-galerie/rapport.json`), confirmation, copie, socle reconstruit si le masque ou la plaquette changent (slides de la plaquette détectées par la disposition « Slide vide claire »), vignettes 1200×675 exportées par PowerPoint dans `build/maj-galerie/vignettes/`, puis génération et test si aucune fiche n'est à rédiger (sinon `--generer` une fois les fiches à jour). Utilisable seul : `--analyse` pour le rapport sans rien modifier.

1. **Vérifier que la copie de travail est à jour** : `list_files` sur l'artefact et comparer les sha256 avec `sha256sum design-system/...`. Si un fichier diffère, le relire (`read_file`, un fichier par appel) et le remplacer dans le kit avant toute modification.
2. **Modifier** sur une branche les fichiers de `design-system/` (et `skill/SKILL.md` si la règle touche le démarrage ou les règles « qui ne se discutent pas »). Garder les deux cohérents : une règle du README se retrouve dans le skill.
3. **Si le masque ou la plaquette change** dans le PowerPoint de référence : `python3 scripts/construire_socle.py reference/galerie-composants-eulidia.pptx pptx/eulidia-socle-2026.pptx 39,40,41,42,43,44,45 "Eulidia 2026 — socle (masque + plaquette)"` (vérifier les numéros : ce sont les slides de disposition « Slide vide claire »). Environ 1 min avec l'optimisation PNG. pyoxipng n'a pas de wheel pour Python 3.14 : sous Windows, lancer avec `py -3.13`.
4. **Générer le skill** : `python3 scripts/generer_skill.py . build`. S'il échoue sur une « phrase introuvable », c'est qu'une phrase traduite (vocabulaire artefact → chemins de fichiers) a changé : mettre à jour la table du script, pas contourner.
5. **Tester** : `python3 scripts/tester_skill.py build/eulidia-powerpoint` (socle, exemple de SKILL.md avec et sans plaquette, audit). Pour un changement visuel, construire aussi un document de test et le rendre (`soffice --headless --convert-to pdf`, `pdftoppm -r 50`), en se rappelant que LibreOffice ne dessine pas les liserés dégradés.
6. **Publier le design system** (outils d'artefacts claude.ai ; s'ils ne sont pas disponibles dans Claude Code, faire cette étape depuis Cowork avec le dépôt cloné) : `publish_files` avec `slug`, `base_version` (de `get_artifact`), `shared_live: "users"`, `mode` patch, un `as: "project/<chemin>"` par fichier modifié. **Puis**, en dernier, mettre à jour `lastChange` de `project/design-system.json` avec `edit_version` (lire d'abord la fin du fichier avec `read_file`) : `at` (UTC), `by: "Arnaud"`, `via: "Claude (Cowork)"`, `note` d'une ligne.
7. **Livrer le skill** : `build/eulidia-powerpoint.zip` (non versionné), qu'Arnaud réimporte dans les skills de l'organisation (remplace l'ancien).
8. **Committer** : la copie `design-system/` reflète l'artefact publié ; ajouter une ligne au *Journal* ; commit (message en français : quoi et pourquoi), merge sur `main`, push.

Nouveau composant : fiche `design-system/components/<Nom>/README.md` (1re ligne = description, puis `**Groupe** : … · **Disposition** : … · **Exemple** : \`slide-NN.jpg\` (groupe d'assets *Exemples*)`, puis règles) + `preview.html` (copier un existant ; l'image est `/_blob/<id>` de l'asset) ; vignette 1200×675 dans `skill/assets/exemples/` et, côté artefact, `upload_asset` dans le groupe *Exemples* puis entrée dans `assetGroups.Exemples` de `design-system.json` ; ligne au catalogue du README (tableau du groupe) et, si besoin, dans les ordres types.

## 3. Décisions prises (ne pas les remettre en cause sans Arnaud)

- **Design system autonome** : aucune mention de l'ancien design system Eulidia d'origine ni de l'artefact « compagnon » (supprimé). Nom : « Eulidia 2026 · PowerPoint ».
- **Retours de revue** (commentaires de l'ancien système) : ignorés.
- **Masque** : 5 dispositions seulement, utilisées telles quelles par leur nom : `Première page`, `Sommaire`, `Section`, `Slide de contenu + barre horizontale`, `Slide vide claire` (réservée à la plaquette).
- **Plaquette commerciale** (7 slides : ChiffresCles, MurLogosClients, OffreEndToEnd, PartenairesTechno, FocusPartenaire, EquipeDirection, EulidiaInstitute) : **systématique en proposition commerciale**, ailleurs **seulement sur demande explicite** (sinon `retirer_plaquette()`). Toujours telle quelle, juste avant la SlideFin (`placer_plaquette()`). Ne sert jamais de modèle.
- **Carte vs panneau** : `fond-carte` #eef1f4 = carte, **liseré coloré obligatoire** ; `fond-panneau` #f2f2f2 = panneau, **jamais de liseré** (option retenue après un audit qui signalait les panneaux).
- **Liseré** : dégradé de 1,5 pt de #eef1f4 vers la couleur d'accent, orienté : l'accent pointe vers ce qui relie les cartes. Angles (`lin ang` en 60000e de degré) : bas-droite 45, bas 90, bas-gauche 135, haut-gauche 225, haut 270, haut-droite 315. Rangée de 3 au-dessus d'une citation : 45/90/135, départs 62/26/62 %. Grille 2×2 : vers le centre. Rangée sous un élément commun : 270. Défaut : 225.
- **Numéro de slide** sur toutes les slides de contenu, de section et de sommaire.
- **Skill dédié** `eulidia-powerpoint` partagé à l'orga, sans méthode imposée ; le lien vers l'artefact a été retiré du skill (les collègues n'en ont pas besoin, et ne peuvent pas lire ses fichiers binaires).

## 4. Pièges techniques déjà rencontrés

python-pptx / PowerPoint :
- Ordre des enfants de `p:spPr` imposé par PowerPoint : xfrm, géométrie, remplissage, `a:ln`, `a:effectLst`… Un `effectLst` avant `ln` fait ignorer le liseré **sans erreur**. `ordonner_sppr()` corrige, `auditer()` contrôle.
- Un `a:xfrm` sans `a:off` (ou avec une hauteur nulle) place la forme en 0,0 : toujours fixer les quatre cotes.
- `prs.slides.add_slide()` ne recopie pas l'espace réservé du numéro (`type="ftr"`, champ `slidenum` ; idx 3 dans la disposition 04, idx 4294967295 pour Section et Sommaire) : d'où `nouvelle_slide()` + `finaliser_numeros()` (après tout déplacement de slides).
- Sous-titre de section : la zone du masque ne fait que 104,5 pt de large. Cotes à fixer : sous-titre x 129,7 / y 384,1 / l 700 / h 23,5 ; titre l 820 / h 53,8. L'audit refuse < 400 pt.
- LibreOffice ne dessine pas les liserés dégradés : ne pas « corriger » ce qui n'est qu'un défaut d'aperçu.
- Retirer des slides : `drop_rel` + retrait de `sldIdLst` suffit, les parties orphelines ne sont pas écrites (socle 10,5 Mo → document sans plaquette ~0,4 Mo).
- Le socle est optimisé sans perte (pyoxipng) : 16,5 → 10,5 Mo, pixels identiques.

Windows / OneDrive :
- Le dépôt est dans OneDrive, qui marque les dossiers en lecture seule : `shutil.rmtree` échoue (« Accès refusé ») sur un dossier vidé. Utiliser `supprimer_dossier()` (dans `generer_skill.py` et `mettre_a_jour.py`).
- Vignettes : les exporter avec PowerPoint (COM, `Slide.Export`), qui dessine les liserés dégradés, contrairement à LibreOffice. L'export ne ferme PowerPoint que s'il n'avait aucune présentation ouverte.

Artefacts :
- Un artefact créé à partir d'un type (Design System) ne sert que du **texte**, un fichier par `read_file` ; `paths` est refusé. Les binaires (pptx, images) ne se relisent pas : d'où le socle et les assets conservés dans ce kit et dans le skill.
- Extensions `.wdp` / `.emf` refusées par `publish_files` (contournement historique : base64 dans un `.b64.txt`). Le socle n'est plus publié dans l'artefact.
- Artefact partagé : `publish_files` et `edit_version` exigent `shared_live: "users"`, sinon erreur 422.
- `design-system.json` (index) s'écrit **en dernier** et se modifie par `edit_version` ; ne jamais republier une vieille copie locale (elle écraserait les assets).
- Les assets d'image s'ajoutent par `upload_asset` ; leurs id (`blob`) sont référencés dans `assetGroups` et dans les `preview.html`.

Skills :
- `propose_skills` (amélioration) exige que le SKILL.md courant ait été lu dans la conversation ; les `old_string` doivent être uniques. Le skill de l'orga, lui, passe par un zip réimporté par Arnaud.

## 5. Contenu du dépôt

```
CLAUDE.md                 consignes pour Claude Code (importe ce fichier)
MAINTENANCE.md            ce fichier
reference/                PowerPoint de référence (galerie des composants)
design-system/            copie de travail des fichiers texte de l'artefact (sans design-system.json ni polices)
skill/SKILL.md            SKILL.md du skill eulidia-powerpoint (source rédigée à la main)
skill/assets/             logos, fonds, exemples (1200×675), icones.json
pptx/eulidia-socle-2026.pptx   socle optimisé (masque + plaquette)
scripts/construire_socle.py    galerie -> socle
scripts/generer_skill.py       dépôt -> build/eulidia-powerpoint/ + zip
scripts/tester_skill.py        contrôle du skill généré
scripts/mettre_a_jour.py       nouvelle galerie -> rapport des différences, socle, vignettes, skill
.claude/commands/maj-galerie.md   commande Claude Code /maj-galerie (procédure complète)
build/                    sortie de génération (ignorée par git)
```

Prérequis : `pip install -r requirements.txt` (sous Windows, `python` au lieu de `python3`) ; LibreOffice et poppler-utils pour les rendus ; polices Poppins pour des rendus fidèles.

## 6. Journal

- 2026-11-17 : création du design system à partir de la galerie (masque 5 dispositions, socle + plaquette, 40 composants).
- 2026-11-17 : corrections liserés (carte grise toujours avec liseré orienté), sous-titres de section, numéros de slide ; module `auditer()`.
- 2026-11-17 : panneaux en #f2f2f2 sans liseré (option 2) ; skill `eulidia-powerpoint` créé et importé dans l'orga ; artefact compagnon abandonné.
- 2026-11-18 : plaquette conditionnelle (proposition commerciale ou demande explicite) : `retirer_plaquette()`, `placer_plaquette()`, `auditer(plaquette=…)`.
- 2026-11-19 : kit de maintenance et scripts de génération reproductibles.
- 2026-11-20 : passage en dépôt Git (Claude Code) avec CLAUDE.md ; la galerie de référence est versionnée dans `reference/`.
- 2026-11-20 : `tester_skill.py` compatible Windows (retour au dossier de départ avant suppression du dossier temporaire, chemin du skill injecté avec des « / »). Design system et skill inchangés. Sous Python 3.14, pyoxipng n'a pas de wheel : reconstruire le socle avec un Python 3.12/3.13.
- 2026-11-20 : titre du socle « Eulidia 2026 — socle (masque + plaquette) » dans la commande de l'étape 3 et par défaut dans `construire_socle.py` (la commande passait « Eulidia 2026 » et aurait changé le titre). Reconstruction testée sous Python 3.13 : pixels et métadonnées identiques au socle versionné. Socle, design system et skill inchangés.
- 2026-11-24 : mise à jour simplifiée depuis une nouvelle galerie : `scripts/mettre_a_jour.py` + commande `/maj-galerie`. Testé sur une variante (visuel, notes, déplacement, retrait, plaquette modifiés) : détection exacte, socle reconstruit (10,6 Mo), vignettes PowerPoint fidèles, skill généré et testé. `generer_skill.py` : suppression de `build/` compatible avec les dossiers OneDrive en lecture seule. `.venv/` ignoré (Python 3.12 via uv). Design system, socle et skill inchangés.
