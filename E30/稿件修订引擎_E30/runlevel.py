# -*- coding: utf-8 -*-
"""Run-level helpers that leave existing tracked revisions intact (E29)."""
import copy, re
from docx.oxml.ns import qn
import tc

def _inside(el, tag):
    a = el.getparent()
    while a is not None:
        if a.tag == qn(tag): return True
        a = a.getparent()
    return False

def _split_run(r, segs):
    """replace run r by consecutive copies holding the given texts; returns the new runs"""
    out = []
    for s in segs:
        n = copy.deepcopy(r)
        for t in n.findall(qn('w:t')): n.remove(t)
        t = tc.etree.SubElement(n, qn('w:t')); t.text = s; t.set(tc.XMLSPACE, 'preserve')
        out.append(n)
    for n in out: r.addprevious(n)
    r.getparent().remove(r)
    return out

def tracked_sub(p_el, pattern, repl):
    """tracked regex substitution inside single text nodes; inserted text is edited directly; deleted text skipped"""
    rx = re.compile(pattern); n = 0
    for t in list(p_el.iter(qn('w:t'))):
        r = t.getparent()
        if r.tag != qn('w:r') or _inside(r, 'w:del') or not rx.search(t.text or ''): continue
        if len(r.findall(qn('w:t'))) != 1: continue
        if _inside(r, 'w:ins'):
            t.text = rx.sub(repl, t.text); n += 1; continue
        s = t.text; pos = 0; segs = []
        for m in rx.finditer(s):
            segs += [('k', s[pos:m.start()]), ('x', m.group(0), m.expand(repl) if isinstance(repl, str) else repl(m))]; pos = m.end()
        segs.append(('k', s[pos:]))
        texts = [x[1] for x in segs if x[1] != '' or x[0] == 'x']
        runs = _split_run(r, [x[1] for x in segs])
        for x, nr in zip(segs, runs):
            if x[0] == 'k':
                if x[1] == '': nr.getparent().remove(nr)
                continue
            new = copy.deepcopy(nr); new.find(qn('w:t')).text = x[2]
            nr.find(qn('w:t')).tag = qn('w:delText')
            dl = tc._stamp('w:del'); nr.addprevious(dl); dl.append(nr)
            ins = tc._stamp('w:ins'); dl.addnext(ins); ins.append(new); n += 1
    return n

def edit_inserted(p_el, old, new):
    """text replacement in a paragraph whose runs are all tracked insertions (e.g. a moved copy); runs are rebuilt as one inserted run"""
    runs = [r for r in p_el.iter(qn('w:r'))]
    assert all(_inside(r, 'w:ins') for r in runs)
    text = ''.join(t.text or '' for r in runs for t in r.findall(qn('w:t')))
    assert text.count(old) == 1, old
    keep = copy.deepcopy(runs[0])
    ins = [c for c in p_el if c.tag == qn('w:ins')]
    for c in ins: p_el.remove(c)
    for t in keep.findall(qn('w:t')): keep.remove(t)
    t = tc.etree.SubElement(keep, qn('w:t')); t.text = text.replace(old, new); t.set(tc.XMLSPACE, 'preserve')
    w = tc._stamp('w:ins'); w.append(keep); p_el.append(w)

def append_inserted(p_el, text):
    ins = [c for c in p_el if c.tag == qn('w:ins')][-1]
    r = copy.deepcopy([c for c in ins if c.tag == qn('w:r')][-1])
    for t in r.findall(qn('w:t')): r.remove(t)
    t = tc.etree.SubElement(r, qn('w:t')); t.text = text; t.set(tc.XMLSPACE, 'preserve'); ins.append(r)
