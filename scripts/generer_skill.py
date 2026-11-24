"""Génère le skill eulidia-powerpoint à partir des sources du kit.

Usage : python3 scripts/generer_skill.py [dossier_kit] [dossier_sortie]
  dossier_kit    : racine du kit (par défaut : le parent de scripts/)
  dossier_sortie : où écrire eulidia-powerpoint/ et eulidia-powerpoint.zip (par défaut : ./build)

Sources lues :
  design-system/README.md          -> references/design-system.md
  design-system/construction.md    -> references/construction.md (module remplacé par un renvoi)
                                      + scripts/eulidia.py (le bloc de code qui contient `def auditer`)
  design-system/components/*/README.md -> references/composants/<Nom>.md
  skill/SKILL.md                   -> SKILL.md (rédigé à la main, copié tel quel)
  skill/assets/                    -> assets/ (logos, fonds, exemples, icones.json)
  pptx/eulidia-socle-2026.pptx     -> assets/eulidia-socle-2026.pptx

Le design system parle de « groupes d'assets » et de « sections » (vocabulaire de l'artefact) ;
le skill parle de chemins de fichiers. Les tables REMPLACEMENTS_* font la traduction.
Si une phrase source change, le script échoue en nommant la phrase introuvable : mettre la table à jour.
"""
import os, pathlib, re, shutil, stat, sys, zipfile


def supprimer_dossier(chemin):
    """rmtree qui passe les dossiers en lecture seule (OneDrive les marque ainsi sous Windows)."""
    def forcer(fonction, p, _):
        os.chmod(p, stat.S_IWRITE)
        fonction(p)
    if sys.version_info >= (3, 12):
        shutil.rmtree(chemin, onexc=forcer)
    else:
        shutil.rmtree(chemin, onerror=forcer)


KIT = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else pathlib.Path(__file__).resolve().parent.parent
OUT = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else pathlib.Path("build")
DS = KIT / "design-system"
SK = OUT / "eulidia-powerpoint"

# (ancien, nouveau, obligatoire) : obligatoire=True -> échec si la phrase est absente
REMPLACEMENTS_README = [
    ("1. **Toujours partir du socle** `eulidia-socle-2026.pptx` : le masque à 5 dispositions + la plaquette commerciale, distribué avec tout le nécessaire de génération dans le skill `eulidia-powerpoint` (voir *Masque PowerPoint*).",
     "1. **Toujours partir du socle** `assets/eulidia-socle-2026.pptx` : le masque à 5 dispositions + la plaquette commerciale.", True),
    ("sinon on la retire du socle.", "sinon on la retire du socle (`retirer_plaquette()`).", True),
    ("10. **Construire avec le module de la section *Construire en python-pptx***",
     "10. **Construire avec le module `scripts/eulidia.py`** (règles dans `references/construction.md`)", True),
    ("(fichiers dans `fonts/`)", "(à installer sur le poste)", True),
    ("images du groupe **Fonds**", "images du `assets/fonds/`", True),
    ("dans le groupe **Logos**", "dans le `assets/logos/`", True),
    ("groupe **Icones**", "`assets/icones.json`", True),
]
REMPLACEMENTS_COMMUNS = [  # README et fiches
    ("groupe d'assets *Exemples*", "`assets/exemples/`", False),
    ("section *Construire en python-pptx*", "`references/construction.md`", False),
    ("voir *Construire en python-pptx*", "voir `references/construction.md`", False),
]
REMPLACEMENTS_FICHES = {
    "Masques": [
        ("depuis le socle (skill `eulidia-powerpoint`, section *Masque PowerPoint*).", "depuis le socle récupéré (`SKILL.md`).", True),
        ("**Vignettes** : groupe d'assets *Masques*", "**Vignettes** : `assets/exemples/slide-13.jpg`", True),
    ],
    "PlaquetteCommerciale": [
        ("(section *Socle et plaquette*)", "(`SKILL.md`)", True),
        ("présentes dans le socle (skill `eulidia-powerpoint`)", "présentes dans le socle récupéré", True),
    ],
}
RENVOI_MODULE = ("Le module complet est fourni dans `scripts/eulidia.py` (à importer, ne pas le recopier) : "
                 "`ANGLES`, `ordonner_sppr()`, `liseré()`, `carte()`, `nouvelle_slide()`, `finaliser_numeros()`, "
                 "`retirer_plaquette()`, `placer_plaquette()`, `auditer()`.")


def remplacer(texte, table, origine):
    for ancien, nouveau, obligatoire in table:
        if ancien in texte:
            texte = texte.replace(ancien, nouveau)
        elif obligatoire:
            raise SystemExit(f"[{origine}] phrase introuvable, mettre à jour generer_skill.py :\n  {ancien[:120]}")
    return texte


def fiche_markdown(nom, texte):
    texte = remplacer(texte, REMPLACEMENTS_FICHES.get(nom, []), nom)
    texte = remplacer(texte, REMPLACEMENTS_COMMUNS, nom)
    if not texte.startswith("# "):
        texte = f"# {nom}\n\n" + texte
    return texte


def main():
    if SK.exists():
        supprimer_dossier(SK)
    (SK / "references" / "composants").mkdir(parents=True)
    (SK / "scripts").mkdir()

    # SKILL.md et assets
    shutil.copy(KIT / "skill" / "SKILL.md", SK / "SKILL.md")
    shutil.copytree(KIT / "skill" / "assets", SK / "assets")
    shutil.copy(KIT / "pptx" / "eulidia-socle-2026.pptx", SK / "assets" / "eulidia-socle-2026.pptx")

    # design-system.md
    readme = (DS / "README.md").read_text(encoding="utf-8")
    readme = remplacer(readme, REMPLACEMENTS_README, "README")
    readme = remplacer(readme, REMPLACEMENTS_COMMUNS, "README")
    if not readme.startswith("# "):
        readme = "# Design system Eulidia 2026 · PowerPoint\n\n" + readme
    (SK / "references" / "design-system.md").write_text(readme, encoding="utf-8")

    # construction.md + module
    cons = (DS / "construction.md").read_text(encoding="utf-8")
    blocs = list(re.finditer(r"```python\n(.*?)```", cons, re.S))
    module = [b for b in blocs if "def auditer" in b.group(1)]
    if len(module) != 1:
        raise SystemExit("construction.md : il faut exactement un bloc python contenant `def auditer`")
    m = module[0]
    (SK / "scripts" / "eulidia.py").write_text(
        '"""Module de construction Eulidia 2026 (généré depuis construction.md du design system)."""\n'
        + m.group(1), encoding="utf-8")
    cons = cons[:m.start()] + RENVOI_MODULE + cons[m.end():]
    cons = remplacer(cons, [("Copier le module ci-dessous dans le script de génération",
                             "Importer `scripts/eulidia.py` dans le script de génération", True)], "construction")
    cons = remplacer(cons, REMPLACEMENTS_COMMUNS, "construction")
    (SK / "references" / "construction.md").write_text(cons, encoding="utf-8")

    # fiches
    n = 0
    for f in sorted((DS / "components").glob("*/README.md")):
        nom = f.parent.name
        (SK / "references" / "composants" / f"{nom}.md").write_text(
            fiche_markdown(nom, f.read_text(encoding="utf-8")), encoding="utf-8")
        n += 1

    # contrôles : plus aucune référence au vocabulaire de l'artefact
    restes = []
    for f in SK.rglob("*.md"):
        t = f.read_text(encoding="utf-8")
        for motif in ("groupe d'assets", "*Construire en python-pptx*", "*Masque PowerPoint*", "*Socle et plaquette*", "claude.ai/artifact"):
            if motif in t:
                restes.append(f"{f.relative_to(SK)} : {motif}")
    if restes:
        raise SystemExit("Références à l'artefact restantes :\n  " + "\n  ".join(restes))

    # zip
    z = OUT / "eulidia-powerpoint.zip"
    if z.exists():
        z.unlink()
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(SK.rglob("*")):
            if f.is_file() and "__pycache__" not in f.parts:
                zf.write(f, f.relative_to(OUT))
    print(f"{n} fiches, skill dans {SK}, zip {z} ({z.stat().st_size // 1024} Ko)")


if __name__ == "__main__":
    main()
