# Masque PowerPoint

Les fichiers PowerPoint de ce système ne sont pas servis par cette page : un design system ne livre à un agent que des fichiers texte. Ils sont distribués dans le skill **`eulidia-powerpoint`** (skills de l'organisation), qui embarque tout ce qu'il faut pour générer un document conforme.

| Fichier du skill | Contenu |
| --- | --- |
| `assets/eulidia-socle-2026.pptx` | Le masque (5 dispositions) + les 7 slides de la plaquette commerciale : point de départ de tout document |
| `references/design-system.md`, `references/construction.md` | Ce README et la section *Construire en python-pptx* |
| `references/composants/<Nom>.md` | Les fiches des composants |
| `scripts/eulidia.py` | Le module `nouvelle_slide`, `carte`, `liseré`, `finaliser_numeros`, `auditer` |
| `assets/icones.json`, `assets/logos/`, `assets/fonds/`, `assets/exemples/` | Icônes, logos, fonds et vignettes d'exemple |

Dispositions du masque, dans l'ordre : `Première page`, `Sommaire`, `Section`, `Slide de contenu + barre horizontale`, `Slide vide claire`. Les images du socle sont optimisées sans perte (pixels identiques au template d'origine).

Pour générer un document : utiliser le skill `eulidia-powerpoint` (seul, ou appelé par un skill de méthode). Ne jamais reconstruire le masque à la main ni partir d'un autre fichier PowerPoint ; si le skill n'est pas disponible, s'arrêter et le signaler.

## Maintenance

Le skill est **dérivé** de ce design system : toute modification d'une fiche, d'une règle, d'une icône, du masque ou de la plaquette se fait ici (le socle PowerPoint, lui, se modifie dans PowerPoint), puis le skill est régénéré et réimporté dans les skills de l'organisation. Sinon les agents travailleront sur une version périmée.
