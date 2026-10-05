# Insereix el joc de plànols A3 (../plànols) darrere de l'índex de documentació gràfica del PDF de la memòria.
from pypdf import PdfReader, PdfWriter
plans = PdfReader('../plànols/20260921 - Tanques.pdf'); mem = PdfReader('mem.pdf')
dg = [i for i, p in enumerate(mem.pages) if i > 3 and 'DG IN. ÍNDEX DE LA DOCUMENTACIÓ GRÀFICA' in (p.extract_text() or '')][0]
w = PdfWriter()
for i in range(dg + 1): w.add_page(mem.pages[i])
for p in plans.pages:
    p.transfer_rotation_to_content(); w.add_page(p)
for i in range(dg + 1, len(mem.pages)): w.add_page(mem.pages[i])
w.write('out.pdf')
