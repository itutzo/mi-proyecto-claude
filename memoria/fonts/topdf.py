import uno, time, subprocess, sys
from com.sun.star.beans import PropertyValue
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','-env:UserInstallation=file:///tmp/lo/profile','--headless','--invisible','--norestore','--accept=socket,host=127.0.0.1,port=2202;urp;'])
ctx=None
for _ in range(60):
    try:
        local=uno.getComponentContext()
        res=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=res.resolve('uno:socket,host=127.0.0.1,port=2202;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
smgr=ctx.ServiceManager
desk=smgr.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
doc=desk.loadComponentFromURL('file:///tmp/lo/build/mem.docx','_blank',0,(pv('Hidden',True),))
for _ in range(2):
    idx=doc.getDocumentIndexes()
    for i in range(idx.getCount()): idx.getByIndex(i).update()
    doc.refresh()
doc.storeToURL('file:///tmp/lo/build/mem.pdf',(pv('FilterName','writer_pdf_Export'),))
doc.close(True)
try: desk.terminate()
except Exception: pass
proc.wait(timeout=30)
print('done')
