# eulidia-ppt-design-system

Sources et outils de maintenance du design system PowerPoint « Eulidia 2026 · PowerPoint » et du skill Claude `eulidia-powerpoint` partagé à l'organisation Eulidia.

- Procédure, décisions et pièges : [MAINTENANCE.md](MAINTENANCE.md)
- Nouvelle version de la galerie : dans Claude Code, `/maj-galerie <fichier.pptx>` (ou `python scripts/mettre_a_jour.py <fichier.pptx>`)
- Générer et tester le skill :

```
pip install -r requirements.txt
python3 scripts/generer_skill.py . build
python3 scripts/tester_skill.py build/eulidia-powerpoint
```

Le zip produit (`build/eulidia-powerpoint.zip`) s'importe dans les skills de l'organisation Claude.
