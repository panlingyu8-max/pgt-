# Tracked-change helpers for python-docx paragraphs (word-level diff, formatting preserved per character)
import copy, re, difflib, itertools
from docx.oxml.ns import qn
from lxml import etree

AUTHOR = 'Claude (E16 修订)'
DATE = '2026-10-02T00:00:00Z'
_ids = itertools.count(91000)
WHOLE_RATIO = 0.6
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


def _chars(p_el):
    out = []
    for r in p_el.iter(qn('w:r')):
        # skip runs already inside a deletion
        if r.getparent().tag == qn('w:del'):
            continue
        rpr = r.find(qn('w:rPr'))
        for c in r:
            if c.tag == qn('w:t'):
                out += [(ch, rpr) for ch in (c.text or '')]
            elif c.tag == qn('w:tab'):
                out.append(('\t', rpr))
    return out


def _rkey(rpr):
    return b'' if rpr is None else etree.tostring(rpr)


def _is_sup(rpr):
    if rpr is None:
        return False
    v = rpr.find(qn('w:vertAlign'))
    return v is not None and v.get(qn('w:val')) == 'superscript'


def _mk_run(text, rpr, deleted=False):
    r = etree.Element(qn('w:r'))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    parts = re.split(r'(\t)', text)
    for part in parts:
        if part == '':
            continue
        if part == '\t':
            etree.SubElement(r, qn('w:tab'))
        else:
            t = etree.SubElement(r, qn('w:delText' if deleted else 'w:t'))
            t.text = part
            t.set(XMLSPACE, 'preserve')
    return r


def _group(chars):
    """[(ch, rpr)] -> [(text, rpr)] grouped by identical formatting"""
    res = []
    for k, g in itertools.groupby(chars, key=lambda x: _rkey(x[1])):
        g = list(g)
        res.append((''.join(c for c, _ in g), g[0][1]))
    return res


def _wrap(tag, runs):
    w = etree.Element(qn(tag))
    w.set(qn('w:id'), str(next(_ids)))
    w.set(qn('w:author'), AUTHOR)
    w.set(qn('w:date'), DATE)
    for r in runs:
        w.append(r)
    return w


def _base_rpr(chars):
    cnt = {}
    for _, r in chars:
        if not _is_sup(r):
            cnt.setdefault(_rkey(r), [0, r])[0] += 1
    if not cnt:
        return None
    return max(cnt.values(), key=lambda x: x[0])[1]


def _parse_sup(text):
    plain, mask, i = [], [], 0
    for m in re.finditer(r'\^\{([^}]*)\}', text):
        seg = text[i:m.start()]; plain.append(seg); mask += [False] * len(seg)
        plain.append(m.group(1)); mask += [True] * len(m.group(1)); i = m.end()
    plain.append(text[i:]); mask += [False] * len(text[i:])
    return ''.join(plain), mask


def _sup_rpr(base):
    r = copy.deepcopy(base) if base is not None else etree.Element(qn('w:rPr'))
    for v in r.findall(qn('w:vertAlign')): r.remove(v)
    v = etree.SubElement(r, qn('w:vertAlign')); v.set(qn('w:val'), 'superscript')
    return r


def revise(paragraph, new_text, track=True):
    p = paragraph._p
    old = _chars(p)
    old_text = ''.join(c for c, _ in old)
    new_text, supmask = _parse_sup(new_text)
    if old_text == new_text:
        return False
    base = _base_rpr(old)
    tok = lambda s: re.findall(r'\s+|[\w≥≤κ%./–\-]+|[^\s\w]', s)
    ot, nt = tok(old_text), tok(new_text)
    # char offsets of old tokens
    off = [0]
    for t in ot:
        off.append(off[-1] + len(t))
    sm = difflib.SequenceMatcher(None, ot, nt, autojunk=False)
    ops = sm.get_opcodes()
    if track and sm.ratio() < WHOLE_RATIO:   # heavy rewrite: show as one deletion + one insertion
        ops = [('replace', 0, len(ot), 0, len(nt))]
    new_children = []
    for op, i1, i2, j1, j2 in ops:
        if op == 'equal':
            for text, rpr in _group(old[off[i1]:off[i2]]):
                new_children.append(_mk_run(text, rpr))
            continue
        if op in ('delete', 'replace') and track:
            runs = [_mk_run(text, rpr, deleted=True) for text, rpr in _group(old[off[i1]:off[i2]])]
            new_children.append(_wrap('w:del', runs))
        if op in ('insert', 'replace'):
            ins = ''.join(nt[j1:j2])
            prev = old[off[i1] - 1][1] if off[i1] > 0 else base
            rpr = prev if (prev is not None and not _is_sup(prev)) else base
            st = sum(len(t) for t in nt[:j1]); m = supmask[st:st + len(ins)]
            runs = []
            for flag, g in itertools.groupby(zip(ins, m), key=lambda x: x[1]):
                runs.append(_mk_run(''.join(c for c, _ in g), _sup_rpr(rpr) if flag else rpr))
            if track:
                new_children.append(_wrap('w:ins', runs))
            else:
                new_children += runs
    # remove old content runs (keep pPr, bookmarks, etc.)
    for c in list(p):
        if c.tag in (qn('w:r'), qn('w:proofErr'), qn('w:ins'), qn('w:del')):
            p.remove(c)
    for c in new_children:
        p.append(c)
    return True


def insert_paragraph_after(paragraph, text, track=True):
    """Clone paragraph formatting, insert a new paragraph after it, with text (tracked as inserted)."""
    src = paragraph._p
    new = copy.deepcopy(src)
    base = _base_rpr(_chars(src))
    for c in list(new):
        if c.tag != qn('w:pPr'):
            new.remove(c)
    ppr = new.find(qn('w:pPr'))
    if ppr is None:
        ppr = etree.SubElement(new, qn('w:pPr'))
        new.insert(0, ppr)
    if track:
        rpr = ppr.find(qn('w:rPr'))
        if rpr is None:
            rpr = etree.SubElement(ppr, qn('w:rPr'))
        rpr.insert(0, _stamp('w:ins'))
    r = _mk_run(text, base)
    new.append(_wrap('w:ins', [r]) if track else r)
    src.addnext(new)
    return new


def accepted_text(p_el):
    """text with insertions kept and deletions dropped"""
    return ''.join(c for c, _ in _chars(p_el))


def _stamp(tag):
    e = etree.Element(qn(tag)); e.set(qn('w:id'), str(next(_ids))); e.set(qn('w:author'), AUTHOR); e.set(qn('w:date'), DATE); return e


def mark(el, kind):
    """Mark a whole paragraph or table element as tracked insertion ('ins') or deletion ('del')."""
    ps = [el] if el.tag == qn('w:p') else list(el.iter(qn('w:p')))
    for p in ps:
        ppr = p.find(qn('w:pPr'))
        if ppr is None:
            ppr = etree.Element(qn('w:pPr')); p.insert(0, ppr)
        rpr = ppr.find(qn('w:rPr'))
        if rpr is None:
            rpr = etree.SubElement(ppr, qn('w:rPr'))
        rpr.insert(0, _stamp('w:' + kind))   # E28: ins/del must be the first children of the paragraph-mark rPr
        for r in [r for r in p if r.tag == qn('w:r')]:
            if kind == 'del':
                for t in r.findall(qn('w:t')): t.tag = qn('w:delText')
            w = _stamp('w:' + kind); r.addprevious(w); w.append(r)
    if el.tag == qn('w:tbl'):
        for tr in el.iter(qn('w:tr')):
            trpr = tr.find(qn('w:trPr'))
            if trpr is None:
                trpr = etree.Element(qn('w:trPr'))
                tr.insert(1 if tr.find(qn('w:tblPrEx')) is not None else 0, trpr)
            trpr.append(_stamp('w:' + kind))


def move_block(els, anchor, track=True):
    """Move consecutive body elements els to after anchor. Tracked: copies inserted, originals deleted."""
    if not track:
        for e in els:
            anchor.addnext(e); anchor = e
        return els
    out = []
    for e in els:
        c = copy.deepcopy(e); mark(c, 'ins'); anchor.addnext(c); anchor = c; out.append(c)
    for e in els:
        mark(e, 'del')
    return out


def delete_paragraph(p_el, track=True):
    if track:
        mark(p_el, 'del')
    else:
        p_el.getparent().remove(p_el)


def swap_image(paragraph, doc, path, width_emu, track=True, docpr_id=None):
    """Replace the picture in a paragraph by a new image (tracked: old drawing run deleted, new one inserted)."""
    from PIL import Image
    p = paragraph._p
    run = [r for r in p.iter(qn('w:r')) if r.find('.//' + qn('w:drawing')) is not None][0]
    rid, _ = doc.part.get_or_add_image(path)
    w, h = Image.open(path).size
    cx, cy = int(width_emu), int(width_emu * h / w)
    new = copy.deepcopy(run)
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    for blip in new.iter('{%s}blip' % ns['a']):
        blip.set('{%s}embed' % ns['r'], rid)
    for ext in new.iter('{%s}extent' % ns['wp']):
        ext.set('cx', str(cx)); ext.set('cy', str(cy))
    for ext in new.iter('{%s}ext' % ns['a']):
        if ext.get('cx') is not None:
            ext.set('cx', str(cx)); ext.set('cy', str(cy))
    for dp in new.iter('{%s}docPr' % ns['wp']):
        dp.set('id', str(docpr_id or next(_ids)))
    if track:
        run.addnext(new)
        w1 = _stamp('w:del'); run.addprevious(w1); w1.append(run)
        w2 = _stamp('w:ins'); new.addprevious(w2); w2.append(new)
    else:
        run.addnext(new); run.getparent().remove(run)
    return rid
