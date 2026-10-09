import sys, zipfile
from lxml import etree
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def ptext(p):
    out=[]
    for el in p.iter():
        if el.tag==W+'t':
            # skip if inside w:del
            anc=el.getparent()
            dele=False
            while anc is not None and anc is not p:
                if anc.tag in (W+'del',W+'moveFrom'): dele=True;break
                anc=anc.getparent()
            if not dele: out.append(el.text or '')
        elif el.tag==W+'tab': out.append('\t')
        elif el.tag==W+'drawing': out.append('[IMG]')
    return ''.join(out)
def main(path):
    z=zipfile.ZipFile(path)
    root=etree.fromstring(z.read('word/document.xml'))
    body=root.find(W+'body')
    lines=[]
    def walk(node, depth=0):
        for ch in node:
            if ch.tag==W+'p':
                lines.append(ptext(ch))
            elif ch.tag==W+'tbl':
                lines.append('<<TABLE>>')
                for tr in ch.iter(W+'tr'):
                    cells=[]
                    for tc in tr.findall(W+'tc'):
                        cells.append(' / '.join(ptext(p) for p in tc.findall('.//'+W+'p')))
                    lines.append(' | '.join(cells))
                lines.append('<<END TABLE>>')
            elif ch.tag==W+'sdt':
                walk(ch.find(W+'sdtContent') if ch.find(W+'sdtContent') is not None else ch)
    walk(body)
    print('\n'.join(lines))
main(sys.argv[1])
