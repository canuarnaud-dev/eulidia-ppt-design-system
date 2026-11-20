# Dépôt de maintenance : design system « Eulidia 2026 · PowerPoint » et skill `eulidia-powerpoint`

Ce dépôt sert uniquement à faire évoluer le design system PowerPoint Eulidia 2026 (artefact claude.ai) et le skill de l'organisation `eulidia-powerpoint` qui en est généré. Langue de travail : français.

Avant toute modification, suivre la procédure de `MAINTENANCE.md` (importé ci-dessous) :
1. vérifier que `design-system/` correspond à l'artefact publié ;
2. modifier le dépôt sur une branche — jamais le skill généré à la main ;
3. `python3 scripts/generer_skill.py . build` puis `python3 scripts/tester_skill.py build/eulidia-powerpoint` ;
4. publier le design system (shared_live "users", `design-system.json` en dernier) ;
5. livrer `build/eulidia-powerpoint.zip` à réimporter dans l'orga, compléter le *Journal*, committer.

Les décisions de la section 3 ne se remettent pas en cause sans l'accord d'Arnaud. Le skill perso `presentation-eulidia` est hors périmètre : ne pas le modifier, seulement signaler une incohérence.

@MAINTENANCE.md
