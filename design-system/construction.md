# Construire en python-pptx

Règles techniques pour produire un document conforme **qui s'affiche correctement dans PowerPoint**. LibreOffice et les aperçus tolèrent des XML que PowerPoint ignore en silence : un rendu LibreOffice correct ne prouve rien. Copier le module ci-dessous dans le script de génération, l'utiliser pour toute forme à liseré, et lancer `auditer()` avant de livrer.

## Quatre règles qui évitent les défauts connus

1. **Ordre XML des propriétés de forme.** Dans `p:spPr`, PowerPoint impose : `a:xfrm` → géométrie → remplissage → `a:ln` → `a:effectLst`. Un `a:ln` placé après `a:effectLst` est ignoré : la forme prend alors le contour par défaut du thème (trait fin bleu uni) au lieu du liseré dégradé. Ne jamais ajouter un élément avec `append` sans réordonner ; utiliser `carte()` ou `ordonner_sppr()`.
2. **Géométrie des espaces réservés : tout ou rien.** Régler une seule cote (la largeur, par exemple) crée un `a:xfrm` sans position, souvent avec une hauteur nulle : PowerPoint place alors la zone en 0,0, en haut à gauche de la slide. Donc soit on n'y touche pas (titre et sous-titre des slides de contenu), soit on fixe `left`, `top`, `width` et `height` ensemble. Cas obligatoire : la **Section**, dont le sous-titre ne fait que 104,5 pt de large dans le masque :

```python
from pptx.util import Pt
for ph in section.placeholders:
    i = ph.placeholder_format.idx
    if i == 0:   # titre
        ph.left, ph.top, ph.width, ph.height = Pt(129.7), Pt(315.1), Pt(820), Pt(53.8)
    if i == 10:  # sous-titre
        ph.left, ph.top, ph.width, ph.height = Pt(129.7), Pt(384.1), Pt(700), Pt(23.5)
```
3. **Numéro de slide : à recopier soi-même.** Le numéro est un espace réservé de pied de page (`type="ftr"`) avec un champ `slidenum`. `add_slide` de python-pptx ne recopie jamais les pieds de page : sans intervention, aucune slide n'a de numéro (la mention légale et le logo, dessinés dans le masque, apparaissent, eux). Toujours créer les slides avec `nouvelle_slide()`, puis appeler `finaliser_numeros()` une fois l'ordre définitif (plaquette déplacée) et juste avant d'enregistrer. Slides concernées : contenu, section, sommaire ; jamais les couvertures, la plaquette ni la SlideFin.
4. **Pas de `p:style` sur les formes dessinées.** `add_shape` en ajoute un qui réintroduit contour, ombre et police du thème dès qu'une propriété manque. `carte()` le retire.

## Le liseré dégradé

La carte « s'allume » : un liseré de 1,5 pt (`bordure-carte`) en dégradé linéaire qui part de la couleur du fond et finit dans la couleur d'accent. **L'accent pointe vers ce qui relie les cartes entre elles** (le centre d'une grille, un élément commun au-dessus ou en dessous) ; à défaut, vers le haut-gauche, là où commence la lecture.

| Accent vers | `a:lin ang` | Degrés |
| --- | --- | --- |
| bas-droite | 2700000 | 45 |
| bas | 5400000 | 90 |
| bas-gauche | 8100000 | 135 |
| haut-gauche | 13500000 | 225 |
| haut | 16200000 | 270 |
| haut-droite | 18900000 | 315 |

Agencements relevés dans les exemples de référence :

| Agencement | Accent vers | Départ du dégradé | Composants |
| --- | --- | --- | --- |
| Carte seule, colonne de cartes, cartes côte à côte indépendantes | haut-gauche (225°), toutes les cartes | couleur du fond à 35 % | ContexteEnjeux, StatutChantier, chiffres clés de ResteAFaire, FichePhase |
| Rangée de cartes **sous** un élément commun (axe, chevrons, en-têtes) | haut (270°), toutes les cartes | couleur du fond à 35 % | Accroche (options), VisionSynthetique, TableauComparatif (colonnes) |
| Rangée de 3 cartes **au-dessus** d'un élément commun (citation, bandeau) | bas-droite / bas / bas-gauche (45° / 90° / 135°) : les accents convergent vers le bas-centre | 62 % / 26 % / 62 % | Carte |
| Grille 2 × 2 autour d'un centre | bas-droite / bas-gauche / haut-droite / haut-gauche pour les tuiles haut-gauche / haut-droite / bas-gauche / bas-droite : convergence vers le centre | 35 % | Sommaire |

- **La couleur de fond dit ce qu'est la forme.** `fond-carte` #eef1f4 = carte : liseré coloré obligatoire. `fond-panneau` #f2f2f2 = panneau, bande de tableau ou fond secondaire : jamais de liseré. L'audit s'appuie sur cette règle.
- **Carte grise (cas par défaut)** : fond `fond-carte` #eef1f4, liseré **toujours coloré** et toujours orienté selon le tableau. Départ du dégradé = #eef1f4 (le liseré naît du fond de la carte).
- **Carte blanche** : seulement là où le composant la prévoit (FichePhase, ligne Objectifs de VisionSynthetique, StatutChantier, chiffres clés de ResteAFaire). Départ noir 10 % (FichePhase, VisionSynthetique, TableauComparatif sans fond) ou #eef1f4 (StatutChantier, ResteAFaire), comme indiqué dans chaque fiche.
- **Ne sont pas des cartes** et restent sans liseré : bandes de tableaux et panneaux en `fond-panneau` #f2f2f2 (DerouleSeance, ChargeClient, ateliers de FichePhase, colonnes de Chiffrage), blocs teintés à 10 % (ColonnesPrerequis), pastilles, bandeaux pleins.
- Liseré plein (non dégradé) : uniquement DemarchePhases et DispositifPilotage.
- Couleur d'accent : celle de la colonne, de la phase ou du volet ; `midnight-blue` (tx2) pour les tuiles de Sommaire.
- Cartes d'une même rangée : même départ et même largeur de liseré ; seuls l'angle (selon l'agencement) et la couleur (selon le contenu) varient.
- Pas d'ombre, pas de liseré gauche seul.

## Module de construction

```python
from lxml import etree
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

ANGLES = {"bas-droite": 45, "bas": 90, "bas-gauche": 135, "haut-gauche": 225, "haut": 270, "haut-droite": 315}
ORDRE_SPPR = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill",
              "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]
SCHEMA = ("accent", "tx", "bg", "dk", "lt")

def _couleur(parent, c):
    """c : 'EEF1F4', 'accent1', 'tx2' ou 'noir10' (noir à 10 %)."""
    if c == "noir10":
        e = etree.SubElement(parent, qn("a:srgbClr"), val="000000")
        etree.SubElement(e, qn("a:alpha"), val="10000")
        return e
    tag = "a:schemeClr" if c.startswith(SCHEMA) else "a:srgbClr"
    return etree.SubElement(parent, qn(tag), val=c.lstrip("#").upper() if tag == "a:srgbClr" else c)

def ordonner_sppr(spPr):
    """Remet les enfants de p:spPr dans l'ordre exigé par PowerPoint."""
    enfants = list(spPr)
    rang = lambda e: ORDRE_SPPR.index(etree.QName(e).localname) if etree.QName(e).localname in ORDRE_SPPR else 99
    for e in enfants:
        spPr.remove(e)
    for e in sorted(enfants, key=rang):
        spPr.append(e)

def liseré(shape, accent, accent_vers="haut-gauche", depart="EEF1F4", pos_depart=35, largeur_pt=1.5):
    """Remplace le contour de la forme par le liseré dégradé Eulidia."""
    sp = shape._element.spPr
    for e in sp.findall(qn("a:ln")) + sp.findall(qn("a:effectLst")):
        sp.remove(e)
    ln = etree.SubElement(sp, qn("a:ln"), w=str(int(largeur_pt * 12700)))
    gf = etree.SubElement(ln, qn("a:gradFill"), flip="none", rotWithShape="1")
    gs_lst = etree.SubElement(gf, qn("a:gsLst"))
    for pos, c in ((pos_depart, depart), (100, accent)):
        _couleur(etree.SubElement(gs_lst, qn("a:gs"), pos=str(pos * 1000)), c)
    etree.SubElement(gf, qn("a:lin"), ang=str(ANGLES[accent_vers] * 60000), scaled="1")
    etree.SubElement(gf, qn("a:tileRect"))
    etree.SubElement(sp, qn("a:effectLst"))
    ordonner_sppr(sp)

def carte(slide, x, y, l, h, accent, accent_vers="haut-gauche", fond="EEF1F4",
          depart=None, pos_depart=35, rayon_pt=16, nom="carte"):
    """Carte signature. x, y, l, h en points. fond=None : sans remplissage.
    depart : couleur de départ du liseré (par défaut celle du fond, noir10 si pas de fond)."""
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(l), Pt(h))
    sh.name = nom
    sh.adjustments[0] = min(0.5, rayon_pt / min(l, h))
    style = sh._element.find(qn("p:style"))
    if style is not None:
        sh._element.remove(style)
    sp = sh._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for e in sp.findall(qn(tag)):
            sp.remove(e)
    if fond:
        _couleur(etree.SubElement(sp, qn("a:solidFill")), fond)
    else:
        etree.SubElement(sp, qn("a:noFill"))
    liseré(sh, accent, accent_vers, depart or (fond or "noir10"), pos_depart)
    return sh

NUMEROTEES = ("Slide de contenu + barre horizontale", "Section", "Sommaire")

def _pied_modele(prs):
    """Copie de l'espace réservé « slide N » de la disposition de contenu."""
    import copy
    lay = next(l for l in prs.slide_layouts if l.name == NUMEROTEES[0])
    sp = next(ph for ph in lay.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.FOOTER)
    return copy.deepcopy(sp._element)

def nouvelle_slide(prs, disposition):
    """add_slide + numéro de slide pour les dispositions de contenu, section et sommaire."""
    lay = next(l for l in prs.slide_layouts if l.name == disposition)
    slide = prs.slides.add_slide(lay)
    if disposition in NUMEROTEES:
        pied = _pied_modele(prs)
        ph = pied.find(f".//{qn('p:ph')}")
        if disposition != NUMEROTEES[0]:
            ph.set("idx", "4294967295")       # pas d'espace équivalent dans la disposition : forme autonome
        ids = [int(e.get("id")) for e in slide._element.iter(qn("p:cNvPr"))]
        pied.find(f".//{qn('p:cNvPr')}").set("id", str(max(ids + [1]) + 1))
        slide.shapes._spTree.append(pied)
    return slide

def finaliser_numeros(prs):
    """Écrit le bon numéro dans chaque champ slidenum (PowerPoint le recalcule, mais les aperçus lisent ce texte)."""
    for n, s in enumerate(prs.slides, 1):
        for fld in s._element.iter(qn("a:fld")):
            if fld.get("type") == "slidenum":
                fld.find(qn("a:t")).text = str(n)

PLAQUETTE = "Slide vide claire"

def retirer_plaquette(prs):
    """Retire les 7 slides de la plaquette du socle (document sans plaquette)."""
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        if prs.part.related_slide(sid.rId).slide_layout.name == PLAQUETTE:
            prs.part.drop_rel(sid.rId)
            lst.remove(sid)

def placer_plaquette(prs):
    """Place la plaquette juste avant la dernière slide (la SlideFin). Sans effet si elle a été retirée."""
    lst = prs.slides._sldIdLst
    ids = list(lst)
    plaq = [i for i in ids if prs.part.related_slide(i.rId).slide_layout.name == PLAQUETTE]
    autres = [i for i in ids if i not in plaq]
    for e in ids:
        lst.remove(e)
    for e in autres[:-1] + plaq + autres[-1:]:
        lst.append(e)

def auditer(chemin, plaquette=None, n_plaquette=7):
    """Retourne la liste des défauts ; liste vide = conforme. À lancer avant toute livraison.
    plaquette=True : elle doit être présente (proposition commerciale) ; False : absente ; None : les deux sont admis."""
    from pptx import Presentation
    A, P = "{http://schemas.openxmlformats.org/drawingml/2006/main}", "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    prs, defauts = Presentation(chemin), []
    slides = list(prs.slides)
    for n, s in enumerate(slides, 1):
        for sp in s._element.iter(P + "spPr"):
            nom = sp.getparent().find(f".//{P}cNvPr").get("name")
            noms = [etree.QName(e).localname for e in sp]
            rangs = [ORDRE_SPPR.index(t) for t in noms if t in ORDRE_SPPR]
            if rangs != sorted(rangs):
                defauts.append(f"slide {n} · {nom} : ordre spPr {noms}")
            x = sp.find(A + "xfrm")
            if x is not None and (x.find(A + "off") is None or x.find(A + "ext") is None):
                defauts.append(f"slide {n} · {nom} : position ou taille manquante")
            elif x is not None and (int(x.find(A + "ext").get("cx")) == 0 or int(x.find(A + "ext").get("cy")) == 0) \
                    and sp.find(A + "prstGeom") is None:
                defauts.append(f"slide {n} · {nom} : largeur ou hauteur nulle")
        for sh in s.shapes:
            sp = sh._element.find(P + "spPr")
            if sp is None or sp.find(A + "prstGeom") is None or sp.find(A + "prstGeom").get("prst") != "roundRect":
                continue
            fond = sp.find(f"{A}solidFill/{A}srgbClr")
            if fond is not None and fond.get("val").upper() == "EEF1F4" and sp.find(f"{A}ln/{A}gradFill") is None:
                defauts.append(f"slide {n} · {sh.name} : carte grise sans liseré coloré")
        if s.slide_layout.name in NUMEROTEES and not any(
                f.get("type") == "slidenum" for f in s._element.iter(A + "fld")):
            defauts.append(f"slide {n} : numéro de slide absent (créer la slide avec nouvelle_slide())")
        if s.slide_layout.name == "Section":
            for ph in s.placeholders:
                if ph.placeholder_format.idx == 10 and ph.width < 400 * 12700:
                    defauts.append(f"slide {n} · sous-titre de section trop étroit (fixer les 4 cotes)")
    if not slides or slides[-1].slide_layout.name != "Première page":
        defauts.append("la dernière slide doit être la SlideFin (disposition Première page)")
    plaq = [i for i, sl in enumerate(slides) if sl.slide_layout.name == PLAQUETTE]
    attendu = list(range(len(slides) - 1 - n_plaquette, len(slides) - 1))
    if plaq and plaq != attendu:
        defauts.append("plaquette incomplète ou mal placée : ses 7 slides vont juste avant la SlideFin (placer_plaquette())")
    if plaquette is True and not plaq:
        defauts.append("plaquette absente : obligatoire dans une proposition commerciale")
    if plaquette is False and plaq:
        defauts.append("plaquette présente alors qu'elle n'est pas demandée (retirer_plaquette())")
    return defauts
```

Exemples :

```python
slide = nouvelle_slide(prs, "Slide de contenu + barre horizontale")
slide.placeholders[10].text = "Le cloud souverain : une maîtrise sur trois niveaux"
# Rangée de 3 cartes grises au-dessus d'une citation (composant Carte)
for i, (vers, pos) in enumerate([("bas-droite", 62), ("bas", 26), ("bas-gauche", 62)]):
    carte(slide, 60 + i * 290, 110, 270, 232, accent="accent3", accent_vers=vers, pos_depart=pos)

# Colonne de fiches blanches (FichePhase)
carte(slide, 32, 100, 317, 111, accent="accent1", fond="FFFFFF", depart="noir10")

placer_plaquette(prs)
finaliser_numeros(prs)
prs.save("document.pptx")
defauts = auditer("document.pptx", plaquette=False)   # True pour une proposition commerciale
assert not defauts, "\n".join(defauts)
```
