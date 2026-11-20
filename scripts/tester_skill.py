"""Teste un skill eulidia-powerpoint généré : exécute l'exemple de SKILL.md avec et sans plaquette,
vérifie que auditer() est vide et qu'il détecte bien l'erreur inverse, puis contrôle le socle.

Usage : python3 scripts/tester_skill.py [build/eulidia-powerpoint]
"""
import os, pathlib, re, shutil, sys, tempfile

SK = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "build/eulidia-powerpoint").resolve()
sys.path.insert(0, str(SK / "scripts"))
from pptx import Presentation  # noqa: E402

DISPOSITIONS = ["Première page", "Sommaire", "Section", "Slide de contenu + barre horizontale", "Slide vide claire"]
socle = Presentation(SK / "assets" / "eulidia-socle-2026.pptx")
assert [l.name for l in socle.slide_layouts] == DISPOSITIONS, [l.name for l in socle.slide_layouts]
assert len(socle.slides) == 7 and all(s.slide_layout.name == "Slide vide claire" for s in socle.slides)

code = re.search(r"```python\n(.*?)```", (SK / "SKILL.md").read_text(encoding="utf-8"), re.S).group(1)
code = code.replace("<dossier du skill>", SK.as_posix())   # « / » : pas de séquence d'échappement sous Windows
assert "avec_plaquette = True" in code, "l'exemple de SKILL.md doit contenir « avec_plaquette = True »"

depart = os.getcwd()
with tempfile.TemporaryDirectory() as d:
    shutil.copy(SK / "assets" / "eulidia-socle-2026.pptx", d)
    os.chdir(d)
    try:
        for avec in (True, False):
            exec(compile(code.replace("avec_plaquette = True", f"avec_plaquette = {avec}"), "SKILL.md", "exec"), {})
            import eulidia
            prs = Presentation("document.pptx")
            noms = [s.slide_layout.name for s in prs.slides]
            assert eulidia.auditer("document.pptx", plaquette=avec) == []
            assert eulidia.auditer("document.pptx", plaquette=not avec), "l'audit ne détecte pas la plaquette absente / en trop"
            print(f"plaquette={avec} : {len(noms)} slides, audit OK, taille {os.path.getsize('document.pptx') // 1024} Ko")
    finally:
        os.chdir(depart)   # Windows refuse de supprimer le dossier courant
print("Skill OK")
