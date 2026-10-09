import sys, zipfile, collections
from lxml import etree
K = sys.argv[1]
schema = etree.XMLSchema(etree.parse(K + '/wml.xsd'))
KEEP = {'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'http://schemas.openxmlformats.org/drawingml/2006/picture', 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'http://www.w3.org/XML/1998/namespace'}
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
def prep(root):
    for ac in list(root.iter('{%s}AlternateContent' % MC)):
        fb = ac.find('{%s}Fallback' % MC); par = ac.getparent(); i = par.index(ac)
        if fb is not None:
            for c in list(fb): par.insert(i, c); i += 1
        par.remove(ac)
    for el in list(root.iter()):
        if not isinstance(el.tag, str): continue
        ns = etree.QName(el).namespace
        if ns not in KEEP and el.getparent() is not None:
            el.getparent().remove(el); continue
        for a in list(el.attrib):
            ans = etree.QName(a).namespace
            if ans and ans not in KEEP: del el.attrib[a]
    return root
for f in sys.argv[2:]:
    root = prep(etree.fromstring(zipfile.ZipFile(f).read('word/document.xml')))
    ok = schema.validate(etree.ElementTree(root))
    errs = collections.Counter(e.message[:150] for e in schema.error_log)
    print('==', f.split('/')[-1], 'valid' if ok else f'{sum(errs.values())} errors')
    for m, n in errs.most_common(12): print('  ', n, m)
