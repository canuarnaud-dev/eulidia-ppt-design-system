Pied de page du masque : « slide N » à gauche, mention légale centrée, logo et slogan à droite.

**Groupe** : Structure · **Disposition** : masque (toutes dispositions) · **Exemple** : `slide-16.jpg` (groupe d'assets *Exemples*)

- Numéro : x 13,6 pt, « slide » Poppins Regular 8 + numéro Poppins Bold 8, `pied-de-page`.
- Mention : `Eulidia © – copyright – diffusion  interdite sans l'accord explicite`, Poppins 7, centrée sur 480 pt, `pied-de-page`.
- Logo : 49,5 × 6,7 pt à x 874,9 / y 513,5, teinte `logo-pied-de-page` ; slogan `DATA INNOVATION PARTNER` Poppins Bold Italic 6 `slogan-logo`.
- Mention légale, logo et slogan : dessinés dans le masque, ils apparaissent seuls ; ne jamais les recréer ni les masquer.
- Numéro de slide : espace réservé de **pied de page** (`type="ftr"`, champ `slidenum`), présent dans la disposition 04 mais **pas recopié par python-pptx** sur les nouvelles slides. Il doit figurer sur chaque slide de contenu, de section et de sommaire : créer les slides avec `nouvelle_slide()` puis appeler `finaliser_numeros()` avant d'enregistrer (section *Construire en python-pptx*). Pas de numéro sur les couvertures, la plaquette ni la SlideFin.
- Laisser la bande 505–540 pt vide de tout autre contenu.
