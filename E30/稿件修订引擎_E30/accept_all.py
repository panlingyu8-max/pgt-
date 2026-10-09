# -*- coding: utf-8 -*-
"""Accept all tracked changes in a .docx (document.xml only; headers/footers untouched unless they carry revisions).
Rules: run-level deletions and moveFrom removed; insertions and moveTo unwrapped; deleted paragraph marks merge the paragraph
into the following one (an emptied paragraph is removed); deleted table rows removed; property-change records removed.
Usage: python accept_all.py in.docx out.docx"""
import sys, zipfile, copy
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
q = lambda t: '{%s}%s' % (W, t)
PROP_CHANGES = ['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'tblGridChange', 'numberingChange', 'tblPrExChange']


def unwrap(el):
    parent = el.getparent(); idx = parent.index(el)
    for c in list(el):
        parent.insert(idx, c); idx += 1
    parent.remove(el)


def accept(root):
    # 1. deleted table rows
    for tr in list(root.iter(q('tr'))):
        trpr = tr.find(q('trPr'))
        if trpr is not None and trpr.find(q('del')) is not None:
            tr.getparent().remove(tr)
    for tbl in list(root.iter(q('tbl'))):          # E30: a table whose rows were all deleted disappears, as in Word
        if tbl.find(q('tr')) is None:
            tbl.getparent().remove(tbl)
    # 2. run-level deletions / moved-from content
    for tag in ('del', 'moveFrom'):
        for el in list(root.iter(q(tag))):
            par = el.getparent()
            if par is None: continue
            if par.tag in (q('rPr'), q('trPr')):     # paragraph-mark / row markers handled below
                continue
            par.remove(el)
    # 3. deleted paragraph marks: merge into next paragraph
    for p in list(root.iter(q('p'))):
        ppr = p.find(q('pPr'))
        rpr = ppr.find(q('rPr')) if ppr is not None else None
        mk = rpr.find(q('del')) if rpr is not None else None
        if mk is None: continue
        rpr.remove(mk)
        content = [c for c in p if c.tag != q('pPr')]
        has_content = any(c.tag in (q('r'), q('ins'), q('hyperlink'), q('moveTo'), q('fldSimple'), q('smartTag')) for c in content)
        nxt = p.getnext()
        if nxt is not None and nxt.tag == q('p'):
            if has_content:
                nppr = nxt.find(q('pPr'))
                pos = 1 if nppr is not None else 0
                for c in content:
                    nxt.insert(pos, c); pos += 1
            p.getparent().remove(p)
        elif not has_content:
            # last paragraph in a cell or before a table: drop it only if the container keeps another paragraph
            par = p.getparent()
            if par.tag != q('tc') or len(par.findall(q('p'))) > 1:
                par.remove(p)
    # 4. insertions: unwrap content, drop paragraph-mark / row markers
    for tag in ('ins', 'moveTo'):
        for el in list(root.iter(q(tag))):
            par = el.getparent()
            if par is None: continue
            if par.tag in (q('rPr'), q('trPr')):
                par.remove(el)
            else:
                unwrap(el)
    for tag in ('moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd'):
        for el in list(root.iter(q(tag))): el.getparent().remove(el)
    for tag in PROP_CHANGES:
        for el in list(root.iter(q(tag))): el.getparent().remove(el)
    # leftover deleted-text containers anywhere else
    for el in list(root.iter(q('delText'))):
        r = el.getparent(); r.remove(el)
    return root


def accept_docx(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('word/') and item.filename.endswith('.xml') and (b'<w:ins ' in data or b'<w:del ' in data or b'Change ' in data or b'Change>' in data):
                root = etree.fromstring(data)
                accept(root)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(item, data)


if __name__ == '__main__':
    accept_docx(sys.argv[1], sys.argv[2])
    print('accepted ->', sys.argv[2])
