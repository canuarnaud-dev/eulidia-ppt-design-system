# Socle et plaquette

`eulidia-socle-2026.pptx` = le masque (5 dispositions) **+ les 7 slides de la plaquette commerciale**, prêtes à l'emploi. C'est le fichier de départ de **tout** document généré ; on en retire la plaquette quand le document n'en a pas besoin. Il est fourni par le skill `eulidia-powerpoint` (`assets/eulidia-socle-2026.pptx`), avec le module de construction (`scripts/eulidia.py`).

Les 7 slides, dans l'ordre : ChiffresCles, MurLogosClients, OffreEndToEnd, PartenairesTechno, FocusPartenaire, EquipeDirection, EulidiaInstitute (composant *PlaquetteCommerciale*). Elles n'ont pas de notes et ne doivent pas être modifiées.

## Construire un document à partir du socle

1. Copier le socle depuis le skill `eulidia-powerpoint`, l'ouvrir avec python-pptx : les 7 slides de la plaquette sont déjà là.
2. **Décider de la plaquette** : elle est **obligatoire dans une proposition commerciale** ; dans tout autre document, elle n'est gardée **que si l'utilisateur la demande explicitement**. Sinon, la retirer tout de suite avec `retirer_plaquette(prs)`.
3. Ajouter les slides du document, en choisissant les dispositions **par leur nom**, et terminer par la SlideFin (disposition `Première page` + fond `fond-couverture-fin.jpg` pleine page, voir composant *SlideFin*).
4. `placer_plaquette(prs)` remet la plaquette, si elle est gardée, **juste avant la SlideFin** ; puis numéroter, enregistrer et auditer :

```python
from pptx import Presentation
# + le module scripts/eulidia.py du skill eulidia-powerpoint (nouvelle_slide, carte, finaliser_numeros, auditer)
prs = Presentation("eulidia-socle-2026.pptx")
avec_plaquette = True        # True : proposition commerciale, ou plaquette demandée explicitement
if not avec_plaquette:
    retirer_plaquette(prs)

couv = nouvelle_slide(prs, "Première page")       # composant Couverture (zones de texte libres)
s = nouvelle_slide(prs, "Slide de contenu + barre horizontale")   # avec le numéro de slide
s.placeholders[10].text = "Peu d'acteurs se démarquent réellement"   # titre
s.placeholders[11].text = "Une première synthèse du benchmark"        # sous-titre
# … autres slides …
fin = nouvelle_slide(prs, "Première page")        # composant SlideFin, toujours en dernier

placer_plaquette(prs)                              # contenu → plaquette → SlideFin (sans effet si retirée)
finaliser_numeros(prs)                             # numéros écrits après le déplacement
prs.save("document.pptx")
assert not auditer("document.pptx", plaquette=avec_plaquette)
```

- Ne jamais supprimer, dupliquer ou réordonner les slides de la plaquette entre elles.
- Toujours créer les slides avec `nouvelle_slide()`, jamais `prs.slides.add_slide()` : c'est elle qui ajoute le numéro de slide, que python-pptx ne recopie pas.
- Vérifier avant livraison (`auditer()` le fait) : la dernière slide est la SlideFin ; si la plaquette est gardée, ses 7 slides la précèdent immédiatement, dans l'ordre ci-dessus.
