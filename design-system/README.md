Eulidia — *Data Innovation Partner*. Ce système sert à produire des **présentations PowerPoint 16:9** fidèles au template Eulidia 2026 : documents génériques, propositions commerciales, supports de kick-off et comités projet (COPROJ). Il se suffit à lui-même : masque, slides types, plaquette, tokens, logos, icônes et polices sont tous ici.

## Règles d'or

1. **Toujours partir du socle** `eulidia-socle-2026.pptx` : le masque à 5 dispositions + la plaquette commerciale, distribué avec tout le nécessaire de génération dans le skill `eulidia-powerpoint` (voir *Masque PowerPoint*). Ne jamais redessiner un titre, une barre ou un pied de page à la main : utiliser la disposition qui les porte.
2. **Seules les 5 dispositions du masque sont autorisées**, appelées par leur nom (composant *Masques*). Toute slide de contenu utilise `Slide de contenu + barre horizontale`.
3. **Seuls les composants de ce système sont autorisés** comme slides types. Choisir le composant par sa **description** (première phrase de chaque fiche) et par son **groupe**, qui dit dans quel type de document il s'emploie.
4. **La plaquette commerciale (7 slides)** termine **systématiquement les propositions commerciales**, juste avant la SlideFin, telle quelle. Dans les autres documents, elle n'est ajoutée **que si l'utilisateur la demande explicitement** ; sinon on la retire du socle. Elle ne sert jamais de modèle pour autre chose.
5. **Format** : 16:9, 960 × 540 pt (12 192 000 × 6 858 000 EMU). Toutes les cotes sont en points ; dans les tokens, 1px = 1pt.
6. **Une seule police : Poppins** (Light 300, Regular 400, Medium 500, Bold 700, italiques).
7. **Une idée par slide**, titre sur une ligne, sous-titre sur une ligne (deux au maximum). Rien sous 505 pt.
8. `rouge-ppt`, `jaune-ppt`, `vert-ppt` sont réservés aux graphiques, pastilles, météos et données — jamais à la charte.
9. Toujours en **formes PowerPoint natives**, éditables : jamais d'image de texte ni de capture de slide.
10. **Construire avec le module de la section *Construire en python-pptx*** (`nouvelle_slide()` pour le numéro de slide, `carte()` pour les liserés) et lancer son `auditer()` avant de livrer : un rendu LibreOffice correct ne garantit pas l'affichage dans PowerPoint.

## Lire une fiche de composant

Chaque composant a une fiche (`components/<Nom>/README.md`) et une vignette (un exemple réel, groupe d'assets *Exemples*) :

- **Composant** : un bloc réutilisable avec une forme et une fonction précises. La fiche donne les règles de construction : disposition, cotes, polices, couleurs, ce qu'il faut fournir (*À fournir*) et ce qu'il ne faut pas faire (*Ne pas*).
- **Groupe** : la famille d'usage — Structure, Slides génériques, Proposition commerciale, Kick-off, Coproj, Plaquette commerciale. Il dit dans quel type de document le composant s'emploie.
- **Description** : la première phrase de la fiche ; elle dit en une ligne à quoi sert le composant et quand l'utiliser. C'est elle qui départage deux composants proches et qui associe un besoin (« une slide prochaines étapes ») au bon composant (*PlanActions*).
- **Disposition** : le masque à utiliser dans le thème PowerPoint.

## Quand utiliser quoi

| Groupe | S'utilise dans | Règle |
| --- | --- | --- |
| Structure | tous les documents | Couverture, sommaire, sections, gabarit de contenu, pied de page, fin. |
| Slides génériques | tous les documents | Blocs utilisables dans toute situation. |
| Proposition commerciale | propositions commerciales **uniquement** | Contexte, prérequis, démarche, phases, calendrier, chiffrage. |
| Kick-off | supports de lancement de mission | Déroulé, enjeux, démarche, plannings niveaux 1–3, parties prenantes, charge, pilotage, plan d'actions. |
| Coproj | comités de pilotage projet | Planning de suivi, statut par chantier, reste à faire. |
| Plaquette commerciale | **fin des propositions commerciales** (systématique) ; ailleurs seulement sur demande explicite | Telle quelle, juste avant la SlideFin. Jamais réutilisée pour générer autre chose. |

Les groupes Proposition commerciale, Kick-off et Coproj ne se mélangent pas : un kick-off n'utilise pas *Chiffrage*, une proposition n'utilise pas *StatutChantier*. Les slides génériques et de structure complètent n'importe lequel des trois.

## Ordres types

- **Document générique** : Couverture → Sommaire (ou SommaireSombre) → pour chaque partie : Section puis slides de contenu (génériques) → SlideFin (plaquette seulement si elle est demandée).
- **Proposition commerciale** : Couverture (sur-titre « PROPOSITION COMMERCIALE ») → BlocTexte (contexte) → ColonnesPrerequis → ChevronsDemarche (vision macro) → démarche détaillée (ChevronsDemarche, volet actif en couleur, autres grisés) → FichePhase / VisionSynthetique (avec IndicateurProgression) → Calendrier → Chiffrage → plaquette (toujours) → SlideFin.
- **Kick-off** : CouvertureKickoff → DerouleSeance → ContexteEnjeux → DemarchePhases → TimelineNiveau1 → TimelineNiveau2 → TimelineNiveau3 → PartiesPrenantes → ChargeClient → DispositifPilotage → PlanActions (prochaines étapes) → Section « ANNEXES » → TableauAteliers → SlideFin (plaquette seulement si elle est demandée).
- **Comité projet (COPROJ)** : CouvertureCoproj → PlanningCoproj → une StatutChantier par chantier → ResteAFaire → SlideFin (plaquette seulement si elle est demandée). Ne pas reprendre les notes ou le compte rendu du comité précédent.


## Masques (dispositions)

| # | Nom dans PowerPoint | Usage | Espaces réservés |
| --- | --- | --- | --- |
| 01 | `Première page` | Couvertures ; SlideFin (fond posé par-dessus) | aucun : zones de texte libres |
| 02 | `Sommaire` | SommaireSombre : fond dégradé, filet et logo blancs | pied de page |
| 03 | `Section` | Changement de partie | Titre (48 pt), sous-titre idx 10 (14 pt menthe-vif), sur-titre idx 11 (12 pt) |
| 04 | `Slide de contenu + barre horizontale` | **Disposition par défaut** de toute slide de contenu | Titre idx 10, sous-titre idx 11, pied de page |
| 05 | `Slide vide claire` | Fond blanc nu, réservé à la plaquette commerciale | aucun |

## Structure d'une slide de contenu (disposition 04)

| Élément | Position (pt) | Typo | Couleur |
| --- | --- | --- | --- |
| Titre | zone x 24,5 · y 25,8 · l 847,7 · h 27 | Poppins Bold 16 (`slide-titre`) | `titre` ; un mot-clé en `violet-flash` ou un préfixe (« Planning : ») en `bleu-data` |
| Sous-titre | zone x 24,5 · y 42,6 · l 455,5 | Poppins Bold 9 (`slide-soustitre`) | `sous-titre` #808080 |
| Barre horizontale | y 72,8, pleine largeur | — | dégradé #808080 → `filet-titre` → transparent |
| Zone de contenu | x 32 → 928 · y 90 → 505 | — | — |
| Pied de page | bande 505–540 : numéro, mention légale, logo et slogan | 8 / 7 / 6 pt | `pied-de-page`, `slogan-logo` |

La mention légale et le logo sont dessinés dans le masque : ils s'affichent seuls, ne pas les recréer. Le **numéro de slide**, lui, est un espace réservé de pied de page que python-pptx ne recopie pas : il faut l'ajouter sur chaque slide de contenu, de section et de sommaire avec `nouvelle_slide()` (section *Construire en python-pptx*). Ne rien poser d'autre dans la bande 505–540 pt.

## Ton et contenu

- **Français**, vouvoiement du client (« vos équipes »), « nous » pour Eulidia.
- Titres en casse de phrase, affirmatifs, qui donnent la conclusion : « Peu d'acteurs se démarquent réellement », « Planning : Sprint #1 lancé en S47, cadrage hors proposition à arbitrer ».
- Sous-titres : le périmètre, la promesse ou le contexte en une phrase.
- CAPITALES réservées aux sur-titres, chevrons, titres de section et de couverture.
- Corps : phrases courtes, mots-clés en **gras**. Pas d'emoji. Montants `34.500€`, charges en `JH`.
- Comités et kick-offs diffusables : rôles ou « Client » / « Eulidia » plutôt que des noms ; pas de TJM ni de montants en euros dans un COPROJ.

## Couleurs

**Principales « flashy »** — `bleu-data` #4c7bec, `rose-lumineux` #cb1870, `violet-flash` #6627a4, `corail-vif` #ff5a3c : signature visuelle, phases 1 → 4 d'une démarche, mots-clés.

**Structurantes** — `midnight-blue` #132048 (éléments clés, logo, dk2 du thème), `bleu-nuit` #000036 (aplats sombres, en-têtes kick-off), `gris-fonce` #636363, `gris-clair` #eef0f3 ; accents `vert-menthe` #58e1c9 et `menthe-vif` #08e5c8 (texte sur fond sombre uniquement).

**Graphiques et évaluation uniquement** — `rouge-ppt` #e2394a, `jaune-ppt` #f4d93d, `vert-ppt` #47c56d.

Thème Office du masque : dk1 #424242 · lt1 #ffffff · dk2 #132048 · lt2 #58e1c9 · accent1 #4c7bec · accent2 #cb1870 · accent3 #6627a4 · accent4 #e2394a · accent5 #f4d93d · accent6 #47c56d · lien #636363 · lien visité #eef0f3. Dans PowerPoint, préférer les couleurs du thème aux valeurs RVB ; `corail-vif` n'est pas encore dans le thème : le saisir en RVB.

Usage : fond des slides de contenu `fond-slide` (crème) ; cartes `blanc` ; cartes signature `fond-carte` #eef1f4 (toujours avec liseré coloré) ; panneaux et bandes `fond-panneau` #f2f2f2 (jamais de liseré) ; texte `texte` #424242 ; titres `titre` #000000. Teintes de fond à ~10 % de la couleur d'accent (`teinte-bleu-data`, `teinte-rose-lumineux`, `teinte-violet-flash`). Un volet ou une phase = une couleur, gardée sur tout le document. La couleur ne porte jamais seule le sens : toujours un libellé.

Contraste : `sous-titre` (3,7:1), `pied-de-page` (1,7:1), `gris-annotation` (2,4:1), `bleu-data` et `corail-vif` sous texte blanc (3,9:1 et 3:1) sont sous 4,5:1 ; valeurs du template conservées, à réserver au texte gras court.

## Typographie

Poppins partout (fichiers dans `fonts/`). Sur un poste sans Poppins, PowerPoint substitue : installer la police avant d'éditer.

| Style | Taille / graisse | Usage |
| --- | --- | --- |
| `accroche` | 54 Bold, interligne 51,3 pt | Message clé géant |
| `section-titre` | 48 Bold CAPITALES | Section |
| `couverture-titre` / `couverture-soustitre` | 40 Bold CAPITALES / 40 Regular | Couvertures |
| `numero-sommaire` | 28 Bold | 01–04 |
| `slide-titre` | 16 Bold | Titre de slide |
| `carte-titre` / `chevron-titre` / `corps-italique` | 14 | Titres de carte, chevrons, citations |
| `surtitre` / `corps` | 12 Bold CAPITALES / 12 Regular | Sur-titres, texte courant |
| `corps-dense` / `intertitre-fiche` | 11 | Fiches, chiffrage |
| `corps-petit` / `pastille` / `chevron-etape` | 10 | Cartes compactes, pastilles, étapes |
| `slide-soustitre` | 9 Bold | Sous-titre |
| `legende` / `pied-numero` | 8 | Légendes ; taille minimale de lecture |

Les plannings (kick-off, COPROJ) descendent ponctuellement à 6,5–7,5 pt pour les libellés de barres et de jalons, comme dans les exemples. Interligne « simple » ; 90 % pour les blocs denses.

## Formes et composition

- **Carte signature** : rectangle arrondi gris `fond-carte` #eef1f4 (`rayon-carte` 16 pt), liseré 1,5 pt **toujours coloré**, en dégradé de #eef1f4 vers la couleur d'accent, orienté selon l'agencement (l'accent pointe vers ce qui relie les cartes ; à défaut vers le haut-gauche). Cartes blanches seulement dans les composants qui le prévoient. Pas d'ombre. Règles et code : section *Construire en python-pptx*.
- Seule ombre autorisée : `ombre-fiche` (VisionSynthetique).
- Pastilles et barres de période : arrondi 50 %, sans contour.
- Chevrons : « Pentagone » pour la première étape puis « Chevron », hauteur 39 pt (27 pt compacte).
- Tableaux toujours **en formes** (bandes arrondies 4 pt en alternance), jamais en tableau Office.
- Angles vifs pour les en-têtes de chiffrage, barres de phase et bandeau de total.
- Fonds sombres : images du groupe **Fonds** sur `bleu-nuit`.

## Logo

- `eulidia-logo-bleu-nuit.svg` (`midnight-blue`) sur fond clair, `eulidia-logo-blanc.svg` sur fond sombre, `eulidia-logo-noir.svg` en monochrome. Logos du masque et de la SlideFin dans le groupe **Logos**.
- Ne pas déformer, recolorer, ni reconstruire le logo en texte ; police du logo (Uniwars) jamais pour du texte.
- Les logos tiers (clients, éditeurs) sont posés tels quels.

## Iconographie

Une seule famille : **Phosphor Fill** (72 icônes, groupe **Icones**), monochromes, recolorées dans la couleur de la colonne. Guillemets pleins `violet-flash` pour les citations. Pas d'autre famille, pas d'emoji.

## Composants par groupe


**Structure**

| Composant | Description | Disposition |
| --- | --- | --- |
| Masques | Les cinq dispositions du masque « Eulidia 2026 » : ce sont les seules autorisées. | toutes |
| Couverture | Slide de couverture (disposition 01) : type de document, titre, sous-titre, filet, destinataire et date sur fond dégradé. | 01 Première page |
| SommaireSombre | Sommaire sur fond sombre (disposition 02) : sur-titre « SOMMAIRE » et parties numérotées en colonnes. | 02 Sommaire |
| Section | Slide de changement de partie (disposition 03) : grand titre en capitales, filet blanc, sous-titre menthe. | 03 Section |
| SlideContenu | Gabarit de toute slide de contenu : titre, sous-titre, barre horizontale, zone de contenu, pied de page (disposition 04). | 04 Slide de contenu + barre horizontale |
| EnTeteSlide | En-tête de slide : titre 16 pt gras, sous-titre 9 pt gras gris, barre horizontale dégradée à 72,8 pt. | 04 Slide de contenu + barre horizontale |
| PiedDePage | Pied de page du masque : « slide N » à gauche, mention légale centrée, logo et slogan à droite. | masque (toutes dispositions) |
| SlideFin | Dernière slide : logo blanc, slogan, contacts Eulidia et adresse sur le fond dégradé. | 01 Première page + fond pleine page |

**Slides génériques**

| Composant | Description | Disposition |
| --- | --- | --- |
| Sommaire | Sommaire en quatre tuiles numérotées 01–04 (cartes signature 400 × 150 pt en grille 2 × 2). | 04 Slide de contenu + barre horizontale |
| Carte | Carte signature Eulidia : fond gris #eef1f4, coins arrondis 16 pt, liseré dégradé qui s'allume dans la couleur d'accent, orienté selon l'agencement des cartes. | 04 Slide de contenu + barre horizontale |
| Citation | Définition ou verbatim mis en exergue entre deux guillemets pleins `violet-flash`. | 04 Slide de contenu + barre horizontale |
| Pastille | Pastilles arrondies pour exprimer criticité, intensité et statut. | 04 Slide de contenu + barre horizontale |
| Accroche | Slide « direction » : message clé en grand, options positionnées sur un axe, recommandation en bandeau. | 04 Slide de contenu + barre horizontale |
| TableauComparatif | Tableau critères × colonnes (Atouts / Points de vigilance) fait de formes, pas d'un tableau Office. | 04 Slide de contenu + barre horizontale |
| SyntheseActeurs | Synthèse de benchmark : un acteur par colonne avec logo, filet, verdict, arguments et action. | 04 Slide de contenu + barre horizontale |

**Proposition commerciale**

| Composant | Description | Disposition |
| --- | --- | --- |
| BlocTexte | Slide de compréhension du contexte : intertitres bleus Contexte / Besoin / Approche et texte justifié. | 04 Slide de contenu + barre horizontale |
| ColonnesPrerequis | Quatre axes d'audit (Use cases, Data management, Technologie, Organisation) : icône, intitulé couleur, questions sur fond teinté. | 04 Slide de contenu + barre horizontale |
| ChevronsDemarche | Vision macro de la démarche : deux grands volets en chevrons contour, étapes en chevrons pleins, descriptions dessous. | 04 Slide de contenu + barre horizontale |
| IndicateurProgression | Petit repère de trois chevrons en haut à droite des slides de détail pour situer l'étape courante. | 04 Slide de contenu + barre horizontale |
| FichePhase | Fiche détaillée d'une phase : cartes Objectifs / Activités / Livrables à gauche, panneau d'ateliers gris à droite. | 04 Slide de contenu + barre horizontale |
| VisionSynthetique | Vue synthétique d'un cadrage : trois étapes en chevrons pleins, chacune avec trois fiches ombrées. | 04 Slide de contenu + barre horizontale |
| Calendrier | Calendrier prévisionnel en Gantt hebdomadaire : mois, semaines, phases colorées et jalons. | 04 Slide de contenu + barre horizontale |
| Chiffrage | Coûts d'intervention : une colonne par volet (staffing, livrables, prix) et un bandeau de total. | 04 Slide de contenu + barre horizontale |

**Kick-off**

| Composant | Description | Disposition |
| --- | --- | --- |
| CouvertureKickoff | Couverture d'un kick-off de mission, déclinaison du composant Couverture. | 01 Première page |
| DerouleSeance | Ordre du jour minuté à gauche, décisions attendues en sortie de séance à droite. | 04 Slide de contenu + barre horizontale |
| ContexteEnjeux | Quatre colonnes de questions à poser au Client, chacune avec son impact sur la mission. | 04 Slide de contenu + barre horizontale |
| DemarchePhases | Rappel de la démarche en quatre phases : en-têtes en chevrons pleins et une fiche par phase (objectif, livrables, modalités, période). | 04 Slide de contenu + barre horizontale |
| TimelineNiveau1 | Planning macro (niveau 1) : phases, rendez-vous et livrables semaine par semaine, sur x semaines. (nombre de semaines variables en fonction du projet) | 04 Slide de contenu + barre horizontale |
| TimelineNiveau2 | Planning méso (niveau 2) : sous-phases numérotées, une barre par sous-phase, livrable en fin de phase, repère « Aujourd'hui ». (nombre de semaines variables en fonction du projet) | 04 Slide de contenu + barre horizontale |
| TimelineNiveau3 | Planning micro (niveau 3) : une phase au jour près, avec un statut par tâche. | 04 Slide de contenu + barre horizontale |
| PartiesPrenantes | Matrice rôles × temps collectifs : qui associer à quel moment. | 04 Slide de contenu + barre horizontale |
| ChargeClient | Temps à mobiliser par profil côté Client (barres de fourchette) et leviers pour limiter la charge. | 04 Slide de contenu + barre horizontale |
| DispositifPilotage | Interlocuteurs Client et Eulidia, et rituels de suivi de la mission, en trois colonnes. | 04 Slide de contenu + barre horizontale |
| PlanActions | Plan d'actions de fin de séance : action, responsable, échéance. | 04 Slide de contenu + barre horizontale |
| TableauAteliers | Détail des temps collectifs : un atelier par ligne, positionné par phase et par semaine. | 04 Slide de contenu + barre horizontale |

**Coproj**

| Composant | Description | Disposition |
| --- | --- | --- |
| CouvertureCoproj | Couverture d'un comité projet (COPROJ), déclinaison du composant Couverture. | 01 Première page |
| PlanningCoproj | Gantt hebdomadaire de suivi : où en est le projet par rapport au plan, à la date du comité. | 04 Slide de contenu + barre horizontale |
| StatutChantier | Une slide par chantier (ex. Modélisation & données, UX/UI & Dataviz) : météo et quatre cartes signature en grille 2 × 2. | 04 Slide de contenu + barre horizontale |
| ResteAFaire | Suivi de charge du projet : consommé / prévu par profil et par phase, synthèse budgétaire et trois chiffres clés. | 04 Slide de contenu + barre horizontale |

**Plaquette commerciale**

| Composant | Description | Disposition |
| --- | --- | --- |
| PlaquetteCommerciale | Sept slides figées à insérer telles quelles juste avant la SlideFin : toujours dans une proposition commerciale, ailleurs seulement sur demande explicite. | 05 Slide vide claire |
