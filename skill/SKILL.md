---
name: eulidia-powerpoint
description: "Design system PowerPoint Eulidia 2026 : socle (masque + plaquette), 40 composants, règles, icônes et module python-pptx embarqués. À utiliser pour produire ou modifier tout .pptx Eulidia, seul ou appelé par un skill de méthode."
---

# Eulidia 2026 · PowerPoint

Ce skill est le **design system** des présentations Eulidia, embarqué avec tout ce qu'il faut pour construire un `.pptx` conforme : le socle PowerPoint, les fiches des composants, les règles, les icônes, les logos, les fonds et un module python-pptx. Il dit **avec quoi** et **comment** construire, pas **quoi** dire : la méthode (cadrage, plan, validation avec l'utilisateur) appartient à l'utilisateur ou au skill qui l'appelle. Sans méthode imposée, construire directement à partir de la demande.

## Contenu du skill

| Chemin | Contenu | Quand le lire |
| --- | --- | --- |
| `references/design-system.md` | Règles d'or, quand utiliser quel groupe, ordres types, structure d'une slide, ton, couleurs, typographie, formes, logo, iconographie, catalogue des composants | **Toujours**, au démarrage |
| `references/construction.md` | Règles techniques python-pptx : ordre XML, espaces réservés, numéro de slide, liserés et agencements, audit | **Toujours**, avant d'écrire le script |
| `references/composants/<Nom>.md` | Fiche d'un composant : cotes, polices, couleurs, liseré, à fournir, à ne pas faire | Seulement pour les composants utilisés |
| `scripts/eulidia.py` | Module : `nouvelle_slide`, `carte`, `liseré`, `ordonner_sppr`, `finaliser_numeros`, `auditer`, `ANGLES` | À importer dans le script de génération |
| `assets/eulidia-socle-2026.pptx` | Les 5 dispositions du masque + les 7 slides de la plaquette commerciale | Point de départ de **tout** document (on en retire la plaquette si elle n'est pas voulue) |
| `assets/icones.json` | 72 icônes Phosphor Fill et logos SVG, indexés par nom (`{"icones": {...}, "logos": {...}}`) | Pour poser une icône ou un logo vectoriel |
| `assets/logos/`, `assets/fonds/` | Logos PNG (blanc des dispositions, blanc de la SlideFin) et fonds JPG (couverture/fin, section) | Couvertures, sections, SlideFin |
| `assets/exemples/slide-NN.jpg` | Rendu de l'exemple de chaque composant (le fichier est cité dans sa fiche) | Pour voir le composant avant de le construire |

Lire les fichiers sur place, dans le dossier du skill : rien n'est à télécharger.

## Démarrer

1. Lire `references/design-system.md` et `references/construction.md`.
2. Choisir les composants : seuls ceux du catalogue sont autorisés, choisis par leur **description** et leur **groupe** (Structure et Slides génériques partout ; Proposition commerciale, Kick-off et Coproj chacun dans son type de document, sans les mélanger). Lire ensuite leurs fiches et regarder leurs vignettes d'exemple.
3. Copier `assets/eulidia-socle-2026.pptx` dans le dossier de travail. Vérifier : 5 dispositions (`Première page`, `Sommaire`, `Section`, `Slide de contenu + barre horizontale`, `Slide vide claire`) et 7 slides.
4. Vérifier python-pptx, LibreOffice et la police Poppins (l'installer si besoin).

## Construire

Tout le document dans **un seul script**, reconstruit entièrement à chaque modification :

```python
import sys; sys.path.insert(0, "<dossier du skill>/scripts")
from eulidia import *                     # nouvelle_slide, carte, liseré, retirer/placer_plaquette, finaliser_numeros, auditer
from pptx import Presentation
from pptx.util import Pt

prs = Presentation("eulidia-socle-2026.pptx")
avec_plaquette = True        # True : proposition commerciale, ou plaquette demandée explicitement ; sinon False
if not avec_plaquette:
    retirer_plaquette(prs)     # le socle contient la plaquette : on la retire tout de suite

couv = nouvelle_slide(prs, "Première page")                       # Couverture (zones de texte libres)
s = nouvelle_slide(prs, "Slide de contenu + barre horizontale")   # numéro de slide inclus
s.placeholders[10].text = "Le cloud souverain : une maîtrise sur trois niveaux"   # titre
s.placeholders[11].text = "Au-delà de la localisation des données"                # sous-titre
for i, (vers, pos) in enumerate([("bas-droite", 62), ("bas", 26), ("bas-gauche", 62)]):
    carte(s, 60 + i * 290, 110, 270, 232, accent="accent3", accent_vers=vers, pos_depart=pos)
# … autres slides …
fin = nouvelle_slide(prs, "Première page")                        # SlideFin, toujours en dernier

placer_plaquette(prs)                                             # contenu → plaquette → SlideFin (sans effet si retirée)
finaliser_numeros(prs)
prs.save("document.pptx")
defauts = auditer("document.pptx", plaquette=avec_plaquette)
assert not defauts, defauts
```

Règles qui ne se discutent pas (le détail est dans les références) :

- **Créer chaque slide avec `nouvelle_slide()`**, jamais `prs.slides.add_slide()` : c'est elle qui ajoute le numéro de slide, que python-pptx ne recopie pas.
- **Ne jamais redessiner** titre, barre, pied de page, mention légale ou logo : ils sont portés par les dispositions. Ne pas toucher à la géométrie des espaces réservés, sauf la Section, dont on fixe toujours les quatre cotes (fiche Section).
- **Cartes et liserés uniquement via `carte()` / `liseré()`**, avec l'angle de l'agencement (`references/construction.md`). `fond-carte` #eef1f4 = carte, liseré coloré obligatoire ; `fond-panneau` #f2f2f2 = panneau, jamais de liseré.
- **Plaquette commerciale** : **toujours dans une proposition commerciale** ; dans tout autre document, **seulement si l'utilisateur la demande explicitement** — sinon `retirer_plaquette(prs)`. Quand elle est gardée : les 7 slides du socle, telles quelles, **juste avant la SlideFin** (`placer_plaquette()`) ; ne jamais en modifier ou en retirer une partie, ni s'en inspirer pour d'autres slides. La disposition `Slide vide claire` ne sert qu'à elles.
- **SlideFin** : `Première page` + `assets/fonds/fond-couverture-fin.jpg` en pleine page (non retourné), logo `assets/logos/eulidia-logo-blanc-fin.png`, un ou deux contacts, adresse fixe « 20, rue Thérèse / 75001 Paris ».
- Formes natives et texte éditable, jamais d'image de texte ni de capture de slide ; Poppins partout ; couleurs en RGB exact des tokens ; icônes depuis `icones.json`, une seule couleur par icône.
- Contenu dans x 32 → 928 pt, y 90 → 505 pt ; rien dans la bande 505–540 pt.

## Vérifier

1. `auditer()` doit renvoyer une liste vide : ordre XML des formes, géométrie des espaces réservés, numéros de slide, cartes grises sans liseré, sous-titres de section, présence et position de la plaquette (`plaquette=True` pour une proposition commerciale, `False` sinon, sauf demande explicite).
2. Un **seul** rendu visuel de l'ensemble, à la fin ou après chaque lot de modifications : `soffice --headless --convert-to pdf`, `pdftoppm -r 50`, planches d'images. Comparer chaque slide à la vignette de son composant.
3. **LibreOffice ne dessine pas les liserés dégradés** : les cartes y paraissent sans contour. Ce n'est pas un défaut du fichier ; ne pas le « corriger ». Les liserés sont garantis par `auditer()`.

## Maintenance

Les fichiers de ce skill sont générés à partir du design system de référence, maintenu par son propriétaire : ne pas les modifier à la main. Toute évolution (fiche, règle, icône, masque, plaquette) passe par le design system, puis par une nouvelle version du skill.
