Carte signature Eulidia : fond gris #eef1f4, coins arrondis 16 pt, liseré dégradé qui s'allume dans la couleur d'accent, orienté selon l'agencement des cartes.

**Groupe** : Slides génériques · **Disposition** : 04 Slide de contenu + barre horizontale · **Exemple** : `slide-23.jpg` (groupe d'assets *Exemples*)

- Fond `fond-carte` #eef1f4 (jamais blanc pour ce composant). Liseré **toujours coloré** : 1,5 pt, dégradé linéaire de #eef1f4 vers l'accent (`bleu-data`, `rose-lumineux`, `violet-flash`, `corail-vif` ; une couleur par carte si les cartes portent des thèmes distincts, sinon la même).
- Angle selon l'agencement : rangée de 3 au-dessus d'une citation → accents vers le bas-centre (45° départ 62 % / 90° départ 26 % / 135° départ 62 %, comme l'exemple) ; rangée sous un élément commun → 270° (haut) ; carte seule ou cartes indépendantes → 225° (haut-gauche) ; grille 2 × 2 → vers le centre. Liseré : voir *Construire en python-pptx*, section « Le liseré dégradé » (fonction `carte()`).
- Contenu : sur-titre CAPITALES 12 couleur (`surtitre`), titre 14 gras noir (`carte-titre`), texte 12 `texte` avec mots-clés en gras.
- Variante compacte centrée (3 colonnes de 202 pt) : sur-titre gris 8, titre 14 CAPITALES sur deux lignes, texte 10 noir ; fond transparent possible.
- Rangées de 3 cartes de même hauteur, gouttière 8–14 pt. Peut être suivie d'une *Citation*.
- Ne pas : laisser une carte grise sans liseré, ajouter une ombre, une bordure colorée pleine ou un liseré gauche seul, donner le même angle à toutes les cartes d'une rangée qui converge.
