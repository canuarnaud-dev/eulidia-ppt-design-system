"""Reconstruit le socle (masque + plaquette) à partir du PowerPoint de référence (galerie des composants).

Usage :
  python3 scripts/construire_socle.py galerie.pptx pptx/eulidia-socle-2026.pptx 39,40,41,42,43,44,45 "Eulidia 2026"
  (39 à 45 = les 7 slides de la plaquette dans la galerie de novembre 2026 : vérifier les numéros
   si la galerie change ; ce sont les slides de disposition « Slide vide claire »)

Étapes : garde le masque complet et les slides listées (notes retirées), supprime les autres slides,
les parties SharePoint / révisions / volets web, la liste des sections et la miniature, réécrit les
propriétés du document, puis optimise les PNG sans perte (pyoxipng) : pixels identiques, ~16 -> ~10 Mo.
Prérequis : pip install python-pptx pyoxipng --break-system-packages
Après : python3 scripts/generer_skill.py puis python3 scripts/tester_skill.py
"""
import sys, copy, zipfile, io, hashlib, json, re
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

SRC = sys.argv[1]; OUT = sys.argv[2]; KEEP = [int(x) for x in sys.argv[3].split(',')] if len(sys.argv) > 3 and sys.argv[3] else []
TITLE = sys.argv[4] if len(sys.argv) > 4 else "Eulidia 2026"

prs = Presentation(SRC)
pp = prs.part
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
for i, sid in enumerate(ids, 1):
    rid = sid.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    if i in KEEP:
        slide_part = pp.related_part(rid)
        for r in list(slide_part.rels.values()):
            if r.reltype == RT.NOTES_SLIDE:
                slide_part.drop_rel(r.rId)
    else:
        sldIdLst.remove(sid)
        pp.drop_rel(rid)
# drop SharePoint / revision parts
for r in list(pp.rels.values()):
    if r.reltype.endswith(('/customXml', '/changesInfo', '/revisionInfo')):
        pp.drop_rel(r.rId)
pkg = pp.package
for r in list(pkg._rels.values()):
    if r.reltype.endswith(('webextensiontaskpanes', '/custom-properties')):
        pkg.drop_rel(r.rId)
# drop section list (refers to removed slides)
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
for ext in pp._element.findall(f'{P}extLst/{P}ext'):
    if ext.get('uri') == '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}':
        ext.getparent().remove(ext)
cp = prs.core_properties
cp.title = TITLE; cp.author = "Eulidia"; cp.last_modified_by = "Eulidia"; cp.revision = 1
cp.subject = ""; cp.keywords = ""; cp.comments = ""
buf = io.BytesIO(); prs.save(buf)

# rewrite app.xml to a minimal one
APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
       'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
       f'<Application>Microsoft Office PowerPoint</Application><PresentationFormat>Grand écran</PresentationFormat><Slides>{len(KEEP)}</Slides>'
       '<Company>Eulidia</Company><AppVersion>16.0000</AppVersion></Properties>').encode('utf-8')
zin = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
names = zin.namelist()
names.sort(key=lambda n: (n != '[Content_Types].xml', n))
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names:
        if n.startswith('docProps/thumbnail'):
            continue
        data = APP if n == 'docProps/app.xml' else zin.read(n)
        if n == '[Content_Types].xml':
            data = re.sub(rb'<Override PartName="/docProps/thumbnail[^>]*/>', b'', data)
            data = re.sub(rb'<Default Extension="jpeg"[^>]*/>', lambda m: m.group(0), data)
        if n == '_rels/.rels':
            data = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b'', data)
        z.writestr(n, data)
print(OUT, len(names))

# Optimisation sans perte des PNG (le reste est déjà compressé)
try:
    import oxipng
    zin = zipfile.ZipFile(OUT); items = [(i, zin.read(i.filename)) for i in zin.infolist()]; zin.close()
    avant = sum(len(d) for _, d in items); apres = 0
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename.lower().endswith('.png'):
                data = oxipng.optimize_from_memory(data, level=2)   # ~1 min ; level=4 gagne peu et prend > 2 min
            apres += len(data)
            z.writestr(info.filename, data)
    print(f"PNG optimisés sans perte : {avant // 1024} -> {apres // 1024} Ko (non compressé)")
except ImportError:
    print("pyoxipng absent : socle non optimisé (pip install pyoxipng --break-system-packages)")
