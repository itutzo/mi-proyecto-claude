from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, black
import io
plans=PdfReader('plans_new.pdf')
mem=PdfReader('mem.pdf')
# page with DG index
dg=None
for i,p in enumerate(mem.pages):
    t=p.extract_text() or ''
    if 'DG IN. ÍNDEX DE LA DOCUMENTACIÓ GRÀFICA' in t and 'Escala' in t and i>3: dg=i
print('dg page', dg)
w=PdfWriter()
for i in range(dg+1): w.add_page(mem.pages[i])
for k,p in enumerate(plans.pages):
    p.transfer_rotation_to_content()
    if k==3:
        W=float(p.mediabox.width); H=float(p.mediabox.height); print('A3',W,H)
        buf=io.BytesIO(); c=canvas.Canvas(buf,pagesize=(W,H))
        c.setFillColor(white); c.rect(1121,42,23,18,stroke=0,fill=1)
        c.setFillColor(black); c.setFont('Helvetica',18.5); c.drawRightString(1141.5,44.6,'03')
        c.save(); buf.seek(0); p.merge_page(PdfReader(buf).pages[0])
    w.add_page(p)
for i in range(dg+1,len(mem.pages)): w.add_page(mem.pages[i])
w.add_metadata({'/Title':'Projecte de substitució de portes d’accés – Nau de Turbines del Besòs','/Author':'Ignasi Tutzó Seró'})
w.write('out.pdf'); print('pages',len(w.pages))
