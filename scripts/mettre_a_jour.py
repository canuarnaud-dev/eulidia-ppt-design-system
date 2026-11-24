"""Met à jour le kit à partir d'une nouvelle version de la galerie des composants.

Usage (depuis la racine du dépôt, avec le .venv Python 3.12/3.13 activé) :
  python scripts/mettre_a_jour.py NOUVELLE_GALERIE.pptx            analyse, confirmation, application
  python scripts/mettre_a_jour.py NOUVELLE_GALERIE.pptx --analyse  rapport seul, rien n'est modifié
  python scripts/mettre_a_jour.py NOUVELLE_GALERIE.pptx --oui      sans demander de confirmation
  python scripts/mettre_a_jour.py --generer                        génère puis teste le skill

Étapes :
  1. Compare la nouvelle galerie à reference/galerie-composants-eulidia.pptx, slide par slide.
     Les slides sont appariées par la 1re ligne de leurs notes (« COMPOSANT : Carte », « FONDATION : … »),
     puis par ordre d'apparition quand une même ligne revient (SlideFin, Section…).
     Contenu comparé : XML de la slide + images et objets liés (pas les notes, comparées à part).
     Masque comparé : masques, dispositions, thèmes et leurs images.
  2. Croise avec design-system/components/ : fiches manquantes ou orphelines, groupe / disposition
     différents des notes, vignettes `slide-NN.jpg` à renuméroter.
  3. Écrit le rapport dans build/maj-galerie/rapport.json (lu par la commande Claude Code /maj-galerie).
  4. Après confirmation : copie la galerie dans reference/, reconstruit le socle si le masque ou la
     plaquette ont changé (slides de la plaquette = disposition « Slide vide claire », détectées seules),
     exporte via PowerPoint les vignettes 1200×675 des slides nouvelles, modifiées ou déplacées
     dans build/maj-galerie/vignettes/ (rien n'est écrasé dans skill/assets/).
  5. Si aucune fiche ni règle n'est à revoir : génère et teste le skill. Sinon s'arrête et liste
     ce qu'il reste à rédiger (c'est le rôle de /maj-galerie), puis relancer avec --generer.
"""
import argparse, hashlib, json, pathlib, re, shutil, subprocess, sys
from collections import Counter

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT



def supprimer_dossier(chemin):
    """rmtree qui passe les dossiers en lecture seule (OneDrive les marque ainsi sous Windows)."""
    import os, stat
    def forcer(fonction, p, _):
        os.chmod(p, stat.S_IWRITE)
        fonction(p)
    if sys.version_info >= (3, 12):
        shutil.rmtree(chemin, onexc=forcer)
    else:
        shutil.rmtree(chemin, onerror=forcer)

sys.stdout.reconfigure(errors="replace")
KIT = pathlib.Path(__file__).resolve().parent.parent
REF = KIT / "reference" / "galerie-composants-eulidia.pptx"
SOCLE = KIT / "pptx" / "eulidia-socle-2026.pptx"
TITRE_SOCLE = "Eulidia 2026 — socle (masque + plaquette)"
COMPOSANTS = KIT / "design-system" / "components"
SORTIE = KIT / "build" / "maj-galerie"
DISPO_PLAQUETTE = "Slide vide claire"
FICHE_PLAQUETTE = "PlaquetteCommerciale"


# ---------- analyse d'une galerie ----------

def _h(data):
    return hashlib.sha256(data).hexdigest()[:16]


def empreinte_slide(slide):
    """Empreinte du contenu : XML où chaque rId est remplacé par l'empreinte de sa cible."""
    xml = slide.part.blob.decode("utf-8")
    jetons = {}
    for rid, rel in slide.part.rels.items():
        if rel.reltype == RT.NOTES_SLIDE:
            jetons[rid] = "notes"
        elif rel.is_external:
            jetons[rid] = "ext:" + rel.target_ref
        elif rel.reltype == RT.SLIDE_LAYOUT:
            jetons[rid] = "disposition:" + slide.slide_layout.name
        else:
            jetons[rid] = _h(rel.target_part.blob)
    xml = re.sub(r'"(rId\d+)"', lambda m: '"%s"' % jetons.get(m.group(1), m.group(1)), xml)
    return _h(xml.encode("utf-8"))


def empreinte_masque(prs):
    """Masques + dispositions + thèmes + images qu'ils utilisent."""
    h = hashlib.sha256()
    parts = {}
    for master in prs.slide_masters:
        for part in [master.part] + [l.part for l in master.slide_layouts]:
            parts[str(part.partname)] = part
            for rel in part.rels.values():
                if not rel.is_external and rel.reltype not in (RT.SLIDE_LAYOUT, RT.SLIDE_MASTER):
                    parts[str(rel.target_part.partname)] = rel.target_part
    for nom in sorted(parts):
        h.update(nom.encode() + b"\0" + parts[nom].blob)
    return h.hexdigest()[:16]


def analyser_galerie(chemin):
    prs = Presentation(chemin)
    slides, vus = [], Counter()
    for num, s in enumerate(prs.slides, 1):
        notes = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
        lignes = [l.strip() for l in notes.splitlines()]
        cle = lignes[0] if lignes and lignes[0] else "(slide sans note)"
        vus[cle] += 1
        champs = {}
        for l in lignes[1:4]:
            m = re.match(r"(Groupe|Disposition|Modèle)\s*:\s*(.*)", l)
            if m:
                champs[m.group(1).lower()] = m.group(2).strip()
        m = re.match(r"(COMPOSANT|FONDATION|SECTION)\s*:\s*(.+)", cle)
        slides.append({
            "num": num,
            "cle": cle,
            "occurrence": vus[cle],
            "type": m.group(1).lower() if m else "autre",
            "nom": m.group(2).strip() if m else cle,
            "groupe": champs.get("groupe"),
            "disposition_notes": champs.get("disposition"),
            "disposition": s.slide_layout.name,
            "contenu": empreinte_slide(s),
            "notes": _h(notes.encode("utf-8")),
            "texte_notes": notes,
        })
    return {
        "chemin": str(chemin),
        "dispositions": [l.name for l in prs.slide_layouts],
        "masque": empreinte_masque(prs),
        "slides": slides,
        "plaquette": [s["num"] for s in slides if s["disposition"] == DISPO_PLAQUETTE],
    }


# ---------- comparaison ----------

def lire_fiches():
    fiches = {}
    for d in sorted(p for p in COMPOSANTS.iterdir() if p.is_dir()):
        f = d / "README.md"
        texte = f.read_text(encoding="utf-8") if f.exists() else ""
        entete = next((l for l in texte.splitlines() if l.startswith("**Groupe**")), "")
        champ = lambda n: (re.search(r"\*\*%s\*\* : ([^·]+)" % n, entete) or [None, None])[1]
        fiches[d.name] = {
            "groupe": (champ("Groupe") or "").strip() or None,
            "disposition": (champ("Disposition") or "").strip() or None,
            "exemples": [int(n) for n in re.findall(r"slide-(\d+)\.jpg", entete)],
        }
    return fiches


def comparer(ancien, nouveau):
    idx = lambda g: {(s["cle"], s["occurrence"]): s for s in g["slides"]}
    A, N = idx(ancien), idx(nouveau)
    r = {"ajoutees": [], "retirees": [], "modifiees": [], "notes_modifiees": [], "deplacees": []}
    for k, s in N.items():
        if k not in A:
            r["ajoutees"].append(s)
            continue
        a = A[k]
        if a["contenu"] != s["contenu"] or a["disposition"] != s["disposition"]:
            r["modifiees"].append(s)
        if a["notes"] != s["notes"]:
            r["notes_modifiees"].append(s)
        if a["num"] != s["num"]:
            r["deplacees"].append(dict(s, ancien_num=a["num"]))
    r["retirees"] = [s for k, s in A.items() if k not in N]

    num_ancien_vers_nouveau = {a["num"]: N[k]["num"] for k, a in A.items() if k in N}
    r["renumerotation"] = {a: n for a, n in num_ancien_vers_nouveau.items() if a != n}
    r["masque_modifie"] = ancien["masque"] != nouveau["masque"] or ancien["dispositions"] != nouveau["dispositions"]
    plaq = lambda g: [(s["cle"], s["contenu"]) for s in g["slides"] if s["disposition"] == DISPO_PLAQUETTE]
    r["plaquette_modifiee"] = plaq(ancien) != plaq(nouveau)
    r["socle_a_reconstruire"] = r["masque_modifie"] or r["plaquette_modifiee"]

    # Croisement avec les fiches du design system
    fiches = lire_fiches()
    comp_nouveaux = {}
    for s in nouveau["slides"]:
        if s["type"] == "composant":
            comp_nouveaux.setdefault(s["nom"], []).append(s)
    a_revoir = []
    for nom, ss in comp_nouveaux.items():
        fiche = nom if nom in fiches else (FICHE_PLAQUETTE if ss[0]["disposition"] == DISPO_PLAQUETTE else None)
        if fiche is None:
            a_revoir.append({"fiche": nom, "raison": "nouveau composant sans fiche", "slides": [s["num"] for s in ss]})
            continue
        if fiche == nom:
            f = fiches[nom]
            if ss[0]["groupe"] and f["groupe"] and ss[0]["groupe"] != f["groupe"]:
                a_revoir.append({"fiche": nom, "raison": f"groupe : fiche « {f['groupe']} », notes « {ss[0]['groupe']} »"})
            if ss[0]["disposition_notes"] and f["disposition"] and ss[0]["disposition_notes"] != f["disposition"]:
                a_revoir.append({"fiche": nom, "raison": f"disposition : fiche « {f['disposition']} », notes « {ss[0]['disposition_notes']} »"})
            for n in f["exemples"]:
                if n in r["renumerotation"]:
                    a_revoir.append({"fiche": nom, "raison": f"vignette slide-{n:02d}.jpg devient slide-{r['renumerotation'][n]:02d}.jpg"})
    for s in r["modifiees"] + r["ajoutees"]:
        if s["type"] == "composant" and s["nom"] in fiches:
            a_revoir.append({"fiche": s["nom"], "raison": f"visuel modifié (slide {s['num']}) : vignette et règles à revoir"})
    for s in r["notes_modifiees"]:
        if s["type"] == "composant":
            fiche = s["nom"] if s["nom"] in fiches else FICHE_PLAQUETTE
            a_revoir.append({"fiche": fiche, "raison": f"notes modifiées (slide {s['num']}) : description et règles à revoir"})
    for s in r["retirees"]:
        if s["type"] == "composant" and s["nom"] not in comp_nouveaux and s["nom"] in fiches:
            a_revoir.append({"fiche": s["nom"], "raison": "composant retiré de la galerie : fiche orpheline"})
    for s in r["modifiees"] + r["notes_modifiees"] + r["ajoutees"] + r["retirees"]:
        if s["type"] == "fondation":
            a_revoir.append({"fiche": "README.md / tokens.json", "raison": f"fondation « {s['nom']} » modifiée (slide {s['num']})"})
        elif s["type"] == "section":
            a_revoir.append({"fiche": "README.md", "raison": f"section « {s['nom']} » modifiée (slide {s['num']}) : groupes et ordres types"})
    if r["masque_modifie"]:
        a_revoir.append({"fiche": "masque/README.md, Masques", "raison": "masque modifié : vignettes et règles du masque à revoir"})
    vus, r["a_revoir"] = set(), []
    for x in a_revoir:
        k = (x["fiche"], x["raison"])
        if k not in vus:
            vus.add(k); r["a_revoir"].append(x)

    exp = {s["num"] for s in r["ajoutees"] + r["modifiees"]} | {s["num"] for s in r["deplacees"]}
    r["vignettes_a_exporter"] = sorted(n for n in exp if next(s for s in nouveau["slides"] if s["num"] == n)["type"] in ("composant", "fondation"))
    return r


def afficher(rapport, ancien, nouveau):
    r = rapport
    print(f"Galerie actuelle : {len(ancien['slides'])} slides · nouvelle : {len(nouveau['slides'])} slides\n")
    rien = not any(r[k] for k in ("ajoutees", "retirees", "modifiees", "notes_modifiees", "deplacees")) and not r["masque_modifie"]
    if rien:
        print("Aucun changement détecté (contenu, notes, ordre et masque identiques).")
        return
    lig = lambda s: f"  - slide {s['num']:>2} · {s['cle']}" + (f" (n°{s['occurrence']})" if s["occurrence"] > 1 else "")
    for titre, cle in (("Slides ajoutées", "ajoutees"), ("Slides retirées (numéro dans l'ancienne galerie)", "retirees"),
                       ("Visuel modifié", "modifiees"), ("Notes modifiées", "notes_modifiees")):
        if r[cle]:
            print(f"{titre} :"); [print(lig(s)) for s in r[cle]]; print()
    if r["deplacees"]:
        print("Slides déplacées :")
        for s in r["deplacees"]:
            print(f"  - {s['cle']} : slide {s['ancien_num']} -> {s['num']}")
        print()
    print(f"Masque : {'MODIFIÉ' if r['masque_modifie'] else 'inchangé'} · "
          f"Plaquette : {'MODIFIÉE' if r['plaquette_modifiee'] else 'inchangée'} (slides {nouveau['plaquette']})")
    if len(nouveau["plaquette"]) != 7:
        print(f"  ATTENTION : la plaquette compte {len(nouveau['plaquette'])} slides au lieu de 7 (décision du MAINTENANCE.md, et tester_skill.py l'exige).")
    print(f"Socle : {'à reconstruire' if r['socle_a_reconstruire'] else 'inchangé'}\n")
    if r["a_revoir"]:
        print("À revoir dans le design system (rédaction, via /maj-galerie) :")
        for x in r["a_revoir"]:
            print(f"  - {x['fiche']} : {x['raison']}")
        print()
    if r["vignettes_a_exporter"]:
        print(f"Vignettes à exporter : slides {r['vignettes_a_exporter']}\n")


# ---------- actions ----------

def lancer(*args):
    print("\n> " + " ".join(str(a) for a in args), flush=True)
    subprocess.run([sys.executable, *map(str, args)], check=True, cwd=KIT)


def reconstruire_socle(plaquette):
    try:
        import oxipng  # noqa: F401
    except ImportError:
        sys.exit("pyoxipng absent : le socle ne serait pas optimisé (16 Mo au lieu de 10). "
                 "Lancer ce script avec le Python du .venv (3.12/3.13).")
    lancer(KIT / "scripts" / "construire_socle.py", REF, SOCLE, ",".join(map(str, plaquette)), TITRE_SOCLE)


PS_EXPORT = r"""
param($src, $dst, $nums)
$ppt = New-Object -ComObject PowerPoint.Application
$deja = $ppt.Presentations.Count
$p = $ppt.Presentations.Open($src, -1, 0, 0)
foreach ($n in $nums.Split(',')) {
  $f = Join-Path $dst ('slide-{0:D2}.jpg' -f [int]$n)
  $p.Slides.Item([int]$n).Export($f, 'JPG', 1200, 675)
  Write-Output $f
}
$p.Close()
if ($deja -eq 0 -and $ppt.Presentations.Count -eq 0) { $ppt.Quit() }
"""


def exporter_vignettes(nums):
    if not nums:
        return
    dst = SORTIE / "vignettes"
    if dst.exists():
        supprimer_dossier(dst)
    dst.mkdir(parents=True)
    ps1 = SORTIE / "exporter_vignettes.ps1"
    ps1.write_text(PS_EXPORT, encoding="utf-8-sig")
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1),
                        str(REF), str(dst), ",".join(map(str, nums))], check=True)
        print(f"Vignettes exportées dans {dst.relative_to(KIT)} (à reprendre dans skill/assets/exemples/ après relecture).")
    except (OSError, subprocess.CalledProcessError) as e:
        print(f"Export des vignettes impossible ({e}) : PowerPoint est-il installé ? Étape à faire à la main.")


def generer():
    lancer(KIT / "scripts" / "generer_skill.py", KIT, KIT / "build")
    lancer(KIT / "scripts" / "tester_skill.py", KIT / "build" / "eulidia-powerpoint")
    print(f"\nSkill prêt : {KIT / 'build' / 'eulidia-powerpoint.zip'} (à réimporter dans les skills de l'organisation).")


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("galerie", nargs="?", help="nouvelle version de la galerie (.pptx)")
    ap.add_argument("--analyse", action="store_true", help="rapport seul, rien n'est modifié")
    ap.add_argument("--oui", action="store_true", help="ne pas demander de confirmation")
    ap.add_argument("--generer", action="store_true", help="générer et tester le skill")
    a = ap.parse_args()
    if a.generer and not a.galerie:
        return generer()
    if not a.galerie:
        ap.error("indiquer la nouvelle galerie (.pptx) ou --generer")
    nouvelle = pathlib.Path(a.galerie).resolve()
    if not nouvelle.exists():
        sys.exit(f"Fichier introuvable : {nouvelle}")

    ancien, nouveau = analyser_galerie(REF), analyser_galerie(nouvelle)
    rapport = comparer(ancien, nouveau)
    SORTIE.mkdir(parents=True, exist_ok=True)
    (SORTIE / "rapport.json").write_text(json.dumps(
        {"ancienne": ancien, "nouvelle": nouveau, "differences": rapport}, ensure_ascii=False, indent=1), encoding="utf-8")
    afficher(rapport, ancien, nouveau)
    print(f"Rapport détaillé : {(SORTIE / 'rapport.json').relative_to(KIT)}")
    if nouveau["dispositions"] != ancien["dispositions"]:
        print(f"ATTENTION : dispositions du masque {nouveau['dispositions']} (attendu : {ancien['dispositions']}).")
    if a.analyse:
        return

    meme = nouvelle == REF.resolve()
    if not meme and not a.oui and input("\nAppliquer cette nouvelle galerie ? [o/N] ").strip().lower() not in ("o", "oui", "y"):
        return print("Rien n'a été modifié.")
    if not meme:
        shutil.copy2(nouvelle, REF)
        print(f"Galerie copiée dans {REF.relative_to(KIT)}")
    if rapport["socle_a_reconstruire"]:
        reconstruire_socle(nouveau["plaquette"])
    exporter_vignettes(rapport["vignettes_a_exporter"])
    if rapport["a_revoir"]:
        print("\nDes fiches ou règles sont à mettre à jour avant de générer le skill (voir la liste ci-dessus).\n"
              "Dans Claude Code : /maj-galerie s'en charge. Ensuite : python scripts/mettre_a_jour.py --generer")
        return
    generer()


if __name__ == "__main__":
    main()
