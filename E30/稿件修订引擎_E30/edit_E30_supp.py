# -*- coding: utf-8 -*-
"""E30 supplement (base: E29 supplement, all revisions accepted).
1. Introductory paragraph, contents, and analysis overview: single study population with two analysis sets; report-availability weighting removed;
   new overview row for the post hoc joint-improvement factor analysis.
2. Table S1A: part A replaced by the unified flow (R56_A/B); part B (with vs without report) deleted; phenotype table rebuilt without the weighted
   column; remaining parts relettered B-G and updated to the 3,045-patient values (E30_values_汇总.txt); note rewritten.
3. Table S1B: 601 admissions without 1-month TTE (R56_C/D); note: 1,445 admissions / 1,419 patients, AF type missing 48, delta coefficient 0.416.
4. Table S11 part A: R50_A values.
5. New Table S13 (R57) and Figure S2 heading/legend (image to be inserted from the redrawn figure files).
Usage: python edit_E30_supp.py <base.docx> <out_trackchanges.docx> <R57_joint_12factor.csv>"""
import sys, copy, csv
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import tc
from runlevel import tracked_sub
tc.AUTHOR = 'Claude (E30 修订)'
tc.DATE = '2026-10-09T12:00:00Z'
BASE, OUT, R57 = sys.argv[1:4]
d = Document(BASE); body = d.element.body
TXT = lambda p: tc.accepted_text(p._p)
P = lambda el: Paragraph(el, d._body)
def allp(): return [P(p) for p in body.iter(qn('w:p'))]
def hits(prefix, exact=False): return [p for p in allp() if (TXT(p).strip() == prefix if exact else TXT(p).startswith(prefix))]
def one(prefix, exact=False):
    h = hits(prefix, exact); assert len(h) == 1, (prefix, len(h)); return h[0]
def rev(p, pairs):
    t = TXT(p)
    for o, n in pairs:
        assert o in t, (o, t[:100]); t = t.replace(o, n)
    tc.revise(p, t)
def rows(tbl): return tbl.findall(qn('w:tr'))
def cells(tr): return tr.findall(qn('w:tc'))
def cell_par(tr, i): return P(cells(tr)[i].find(qn('w:p')))
def set_cell(tr, i, new):
    p = cell_par(tr, i)
    if TXT(p) != new: tc.revise(p, new)
def settext(p_el, text):
    """untracked: replace the text of a paragraph (used inside copies that are marked as inserted afterwards)"""
    rs = list(p_el.iter(qn('w:r'))); rpr = None
    for r in rs:
        if r.find(qn('w:t')) is not None:
            rpr = r.find(qn('w:rPr')); break
    for c in list(p_el):
        if c.tag in (qn('w:r'), qn('w:ins'), qn('w:del'), qn('w:proofErr'), qn('w:hyperlink')): p_el.remove(c)
    p_el.append(tc._mk_run(text, copy.deepcopy(rpr) if rpr is not None else None))
def set_row(tr, texts):
    for c, t in zip(cells(tr), texts):
        ps = c.findall(qn('w:p'))
        for extra in ps[1:]: c.remove(extra)
        settext(ps[0], t)
def set_widths(tbl, W):
    for g, w in zip(tbl.find(qn('w:tblGrid')).findall(qn('w:gridCol')), W): g.set(qn('w:w'), str(w))
    for r in rows(tbl):
        for c, w in zip(cells(r), W):
            pr = c.find(qn('w:tcPr'))
            tw = pr.find(qn('w:tcW')) if pr is not None else None
            if tw is not None: tw.set(qn('w:w'), str(w)); tw.set(qn('w:type'), 'dxa')
def table_after(p): t = p._p.getnext(); assert t.tag == qn('w:tbl'); return t
def mark_row_ins(tr):
    for p in tr.iter(qn('w:p')): tc.mark(p, 'ins')
    trpr = tr.find(qn('w:trPr'))
    if trpr is None:
        trpr = tc.etree.Element(qn('w:trPr')); tr.insert(1 if tr.find(qn('w:tblPrEx')) is not None else 0, trpr)
    trpr.append(tc._stamp('w:ins'))
def rebuild(old, spec, tmpl, drop_last_col=False):
    """tracked replacement of a whole table: old marked deleted, a rebuilt copy inserted after it.
    spec: [(kind, [texts])]; tmpl: {kind: row index in old table to clone}"""
    new = copy.deepcopy(old); R = rows(new); T = {k: copy.deepcopy(R[i]) for k, i in tmpl.items()}; last_tcpr = [copy.deepcopy(c.find(qn('w:tcPr'))) for c in cells(R[-1])]
    for r in R: new.remove(r)
    out = []
    for kind, texts in spec:
        r = copy.deepcopy(T[kind]); set_row(r, texts); new.append(r); out.append(r)
    for c, pr in zip(cells(out[-1]), last_tcpr):            # bottom border of the original last row
        if pr is not None:
            old_pr = c.find(qn('w:tcPr'))
            if old_pr is not None: c.remove(old_pr)
            c.insert(0, pr)
    if drop_last_col:
        gc = new.find(qn('w:tblGrid')).findall(qn('w:gridCol')); W = [int(g.get(qn('w:w'))) for g in gc]
        for r in rows(new): r.remove(cells(r)[-1])
        new.find(qn('w:tblGrid')).remove(gc[-1]); extra = W[-1] // (len(W) - 2)
        set_widths(new, [W[0]] + [x + extra for x in W[1:-1]])
    tc.mark(new, 'ins'); old.addnext(new); tc.mark(old, 'del')
    return new

# ================================================================ intro paragraph, contents
rev(one('Baseline regurgitation phenotypes were described in 2,322 patients'),
    [('Baseline regurgitation phenotypes were described in 2,322 patients with a TTE report confirmed as pre-procedural (baseline phenotype cohort; Table S1A). '
      'One-month responses were assessed in a separately assembled population of 844 ablation procedures in 837 patients (paired response cohort; 7 patients underwent 2 ablations, with each procedure analyzed separately). '
      'These sampling populations were partially overlapping, with 357 shared patients; the paired response cohort was not nested within the baseline phenotype cohort. '
      'For shared patients, the baseline grades of the paired analysis were used in both populations.',
      'Baseline regurgitation phenotypes were described in 3,045 patients with an in-hospital baseline TTE report (baseline phenotype set; Table S1A). '
      'One-month responses were assessed in 844 ablation procedures in 837 patients (paired response set; 7 patients underwent 2 ablations, with each procedure analyzed separately). '
      'Both analysis sets were drawn from one study population of 3,551 patients and shared 372 patients; for the 366 whose baseline-set admission was a paired procedure, the baseline grades of the paired analysis were used in both sets.'),
     ('analyzed in parallel within these sampling populations', 'analyzed in parallel within these analysis sets')])
toc12 = hits('Table S12. Aortic regurgitation as a negative control outcome')[0]
tc.insert_paragraph_after(toc12, 'Table S13. Baseline factors associated with improvement of both valves in combined MR/TR (post hoc)')
tocF1 = hits('Figure S1. Weighted comparisons of improvement between isolated and combined regurgitation')[0]
tc.insert_paragraph_after(tocF1, 'Figure S2. Standardized probabilities of TR improvement across RA and LV end-diastolic diameters')

# ---------------------------------------------------------------- analysis overview
ov = [tr for tr in body.iter(qn('w:tr')) if cells(tr) and TXT(cell_par(tr, 0)).startswith('Baseline regurgitation phenotypes (baseline phenotype cohort)')]
assert len(ov) == 1
set_cell(ov[0], 3, 'Patients without an in-hospital baseline TTE report not included and not assumed free of regurgitation; registry grades (recorded only with regurgitation) not used for prevalence; grades of the paired analysis used when the baseline-set admission was a paired procedure')
jr = [tr for tr in body.iter(qn('w:tr')) if cells(tr) and TXT(cell_par(tr, 0)).startswith('Joint response by extent of biatrial reduction')]
assert len(jr) == 1
nr = copy.deepcopy(jr[0])
set_row(nr, ['Baseline factors and improvement of both valves in combined MR/TR', 'Post hoc',
             'The 12 baseline factors with both baseline grades in place of phenotype: logistic regression of improvement of both valves, proportional-odds regression of the number of improved valves, and logistic regression of extension to both valves among procedures with ≥1 improved valve; Benjamini–Hochberg false discovery rate across the 12 factors (Table S13)',
             'PMM, m=20, as for the 12-factor model', 'Patient-clustered robust standard errors (logistic models); Rubin’s rules'])
jr[0].addnext(nr); mark_row_ins(nr)

# ================================================================ Table S1A
for p in hits('Table S1A. Baseline regurgitation phenotypes in patients undergoing a first AF ablation'):
    rev(p, [('in patients undergoing a first AF ablation', 'in patients undergoing AF ablation')])
s1a_notes = 'Baseline phenotype cohort: patients with a first AF ablation admission'
rev(one('A. Cohort assembly', exact=True), [('A. Cohort assembly', 'A. Study population and analysis sets')])
tA = table_after(one('A. Study population and analysis sets', exact=True)) if False else table_after(hits('A. Cohort assembly')[0] if hits('A. Cohort assembly') else one('A. Study population and analysis sets'))
rebuild(tA, [('h', ['Step', 'n']),
    ('m', ['AF ablation patients, January 2020 to September 2024, with an in-hospital baseline TTE report', '3,682']),
    ('m', ['Excluded: prior or concurrent valve surgery (71 from records; 3 identified on review)', '74']),
    ('m', ['Excluded on review: organic valve disease, congenital heart disease, or obstructive hypertrophic cardiomyopathy', '57']),
    ('s', ['Rheumatic heart disease', '12']), ('s', ['Mitral stenosis, mild or moderate', '12']), ('s', ['Aortic stenosis, mild-to-moderate or worse', '1']),
    ('s', ['Valve prolapse', '5']), ('s', ['Obstructive hypertrophic cardiomyopathy', '8']), ('s', ['Unrepaired atrial septal defect with shunt', '16']),
    ('s', ['Residual defect or shunt after repair or closure', '2']), ('s', ['Complex congenital heart disease', '1']),
    ('m', ['Study population', '3,551']),
    ('m', ['Baseline phenotype set: first ablation admission in the study period, baseline TTE report graded to estimate prevalence', '3,045']),
    ('s', ['Paroxysmal AF / persistent AF', '1,848 / 1,197']),
    ('m', ['Registry admissions with MR or TR of analysis grade ≥1', '1,445 admissions; 1,419 patients']),
    ('s', ['No 1-month TTE (Table S1B)', '601 admissions; 592 patients']),
    ('m', ['Paired response set', '844 procedures; 837 patients']),
    ('m', ['Patients in both analysis sets', '372']),
    ('s', ['Baseline-set admission was a paired procedure (baseline grades of the paired analysis used)', '366'])],
    {'h': 0, 'm': 1, 's': 8})
pB = one('B. Patients with and without a confirmed pre-procedural TTE report', exact=True)
tB = table_after(pB); tc.delete_paragraph(pB._p); tc.mark(tB, 'del')
pC = one('C. MR/TR phenotypes, n (%; Wilson 95% CI)', exact=True)
tC = table_after(pC); rev(pC, [('C. MR/TR', 'B. MR/TR')])
rebuild(tC, [('h', ['Phenotype', 'All (n=3,045)', 'Paroxysmal AF (n=1,848)', 'Persistent AF (n=1,197)']),
    ('m', ['Neither', '1,474 (48.4; 46.6–50.2)', '1,154 (62.4; 60.2–64.6)', '320 (26.7; 24.3–29.3)']),
    ('m', ['MR only', '303 (10.0; 8.9–11.1)', '188 (10.2; 8.9–11.6)', '115 (9.6; 8.1–11.4)']),
    ('m', ['TR only', '605 (19.9; 18.5–21.3)', '291 (15.7; 14.2–17.5)', '314 (26.2; 23.8–28.8)']),
    ('m', ['MR and TR', '663 (21.8; 20.3–23.3)', '215 (11.6; 10.3–13.2)', '448 (37.4; 34.7–40.2)'])],
    {'h': 0, 'm': 1}, drop_last_col=True)
pD = one('D. Study-defined moderate or greater (analysis grade ≥2), n/N (%)', exact=True)
tD = table_after(pD); rev(pD, [('D. Study', 'C. Study')])
VD = {'All': ['319/3,045 (10.5)', '519/3,045 (17.0)'], 'Paroxysmal AF': ['120/1,848 (6.5)', '154/1,848 (8.3)'], 'Persistent AF': ['199/1,197 (16.6)', '365/1,197 (30.5)'],
      'MR only': ['69/303 (22.8)', '—'], 'TR only': ['—', '169/605 (27.9)'], 'MR and TR': ['250/663 (37.7)', '350/663 (52.8)']}
for tr in rows(tD)[1:]:
    k = TXT(cell_par(tr, 0)).strip(); set_cell(tr, 1, VD[k][0]); set_cell(tr, 2, VD[k][1])
pE = one('E. Persistent vs paroxysmal AF', exact=True)
tE = table_after(pE); rev(pE, [('E. Persistent', 'D. Persistent')])
VE = [['MR and TR (n=3,045)', '11.6', '37.4', '3.22 (2.78–3.72)', '3.32 (2.88–3.82)'], ['Any MR (n=3,045)', '21.8', '47.0', '2.16 (1.94–2.40)', '2.23 (2.01–2.47)'],
      ['Any TR (n=3,045)', '27.4', '63.7', '2.32 (2.13–2.53)', '2.34 (2.15–2.54)'], ['MR and TR, among patients with MR or TR (n=1,571)', '31.0', '51.1', '1.65 (1.45–1.88)', '1.79 (1.57–2.03)']]
for tr, v in zip(rows(tE)[1:], VE):
    for i, x in enumerate(v): set_cell(tr, i, x)
pF = one('F. Co-occurrence of MR and TR', exact=True)
tF = table_after(pF); rev(pF, [('F. Co-occurrence', 'E. Co-occurrence')])
VF = [['Observed MR and TR, n (%)', '663 (21.8)'], ['Observed/expected, marginal: all', '1.65 (1.58–1.72)'], ['Observed/expected, marginal: paroxysmal AF', '1.95 (1.80–2.11)'],
      ['Observed/expected, marginal: persistent AF', '1.25 (1.20–1.30)'], ['Observed/expected, conditional on age, sex, and AF type', '1.27 (1.22–1.32)'],
      ['OR for TR with vs without MR, unadjusted', '5.33 (4.52–6.29)'], ['OR, adjusted for age, sex, and AF type', '3.31 (2.76–3.97)'],
      ['OR, further adjusted for LA, RA, and LV end-diastolic diameters and LVEF (n=2,525)', '2.85 (2.27–3.59)']]
for tr, v in zip(rows(tF)[1:], VF):
    for i, x in enumerate(v): set_cell(tr, i, x)
pG = one('G. Chamber dimensions by phenotype', exact=True)
tG = table_after(pG); rev(pG, [('G. Chamber', 'F. Chamber')])
VG = [['37 (33–41)', '41 (38–45)', '40 (37–44)', '43 (40–47)'], ['Reference', '+3.6 (2.9 to 4.3)', '+1.4 (0.8 to 2.0)', '+4.2 (3.6 to 4.8)'],
      ['35 (32–38)', '37 (33–42)', '41 (36–46)', '44 (40–48)'], ['Reference', '+1.8 (1.1 to 2.4)', '+3.6 (3.0 to 4.1)', '+6.1 (5.5 to 6.7)'],
      ['47 (45–50)', '50 (46–53)', '47 (44–50)', '49 (46–52)'], ['Reference', '+2.8 (2.1 to 3.5)', '−0.4 (−0.8 to 0.1)', '+2.4 (1.8 to 3.0)'],
      ['22 (20–23)', '22 (20–23)', '22 (21–24)', '23 (21–25)'], ['Reference', '+0.2 (−0.1 to 0.5)', '+0.4 (0.1 to 0.7)', '+1.2 (0.9 to 1.5)'],
      ['66 (62–70)', '64 (58–68)', '65 (60–69)', '61 (53–68)'], ['Reference', '−3.5 (−4.6 to −2.3)', '−0.3 (−1.0 to 0.4)', '−5.2 (−6.2 to −4.1)']]
for tr, v in zip(rows(tG)[1:], VG):
    for i, x in enumerate(v): set_cell(tr, i + 1, x)
pH = one('H. Patients with baseline MR or TR: in vs not in the paired cohort', exact=True)
tH = table_after(pH); rev(pH, [('H. Patients with baseline MR or TR: in vs not in the paired cohort', 'G. Patients with baseline MR or TR: in vs not in the paired response set')])
VH = {'Variable': ['In paired response set (n=366), %', 'Not in paired response set (n=1,205), %'], 'Age, y, mean': ['66.5', '65.7'], 'Female sex': ['53.8', '46.7'],
      'Persistent AF': ['57.4', '55.4'], 'MR and TR': ['42.3', '42.2'], 'MR only': ['20.5', '18.9'], 'TR only': ['37.2', '38.9'], 'MR grade ≥2': ['24.9', '18.9'],
      'TR grade ≥2': ['35.2', '32.4'], 'LA diameter, mm, median': ['42', '42'], 'RA diameter, mm, median': ['41', '42'], 'Concomitant LAAO': ['37.7', '12.0']}
for tr in rows(tH):
    k = TXT(cell_par(tr, 0)).strip(); set_cell(tr, 1, VH[k][0]); set_cell(tr, 2, VH[k][1])
tc.revise(one(s1a_notes),
    'Study population: patients undergoing AF ablation between January 2020 and September 2024 with an in-hospital baseline TTE report. '
    'Prior or concurrent valve replacement or repair was identified from reports, discharge diagnoses, and procedure codes, and organic valve disease, congenital heart disease, and obstructive '
    'hypertrophic cardiomyopathy were excluded after review by the investigating team; the same criteria applied to both analysis sets. Each patient contributes one admission, the first ablation admission in the study period, to the baseline '
    'phenotype set, whereas the paired response set counts procedures. MR and TR were graded on the institutional eight-level scale from the report conclusion (for example, "mild mitral '
    'regurgitation") or, if not stated there, from the reported regurgitant volume in the findings; trace regurgitation and regurgitation not mentioned were recorded as none. Against manual '
    'abstraction of the same 692 reports in the study registry, exact agreement on the eight-level scale was 97.4% (MR) and 97.8% (TR), with linear weighted κ of 0.96 and 0.97. A valve was '
    'classified as affected at analysis grade ≥1 (very mild or worse), as in the paired response set. For the 366 patients whose baseline-set admission was also a paired procedure, the baseline grades used '
    'in the paired analysis were applied; they differed from the extracted grades in 26. The other 6 patients in both sets contributed a different admission to each set. Registry grades were recorded only for patients with regurgitation and were therefore not used to estimate prevalence. AF type was taken from the '
    'registry or, if unavailable, from clinical records. Prevalence ratios are from modified Poisson regression with robust standard errors. Expected joint prevalence was the product of the '
    'marginal prevalences (marginal) or the sum over patients of the product of individual probabilities from logistic models of MR and of TR on age, sex, and AF type (conditional); 95% CIs '
    'from 1,000 (marginal) and 500 (conditional) bootstrap resamples. Adjusted differences in chamber dimensions are from linear regression on phenotype, age, sex, and AF type with robust '
    'standard errors; chamber dimensions were available from the same report in 2,755–3,009 patients. Part G compares patients of the baseline phenotype set with MR or TR who were or were '
    'not included in the paired response set. All analyses of this table were added post hoc and are descriptive. AF indicates atrial fibrillation; CI, confidence interval; LA, left atrial; '
    'LAAO, left atrial appendage occlusion; LV, left ventricular; LVEF, left ventricular ejection fraction; MR, mitral regurgitation; OR, odds ratio; RA, right atrial; RV, right ventricular; '
    'TR, tricuspid regurgitation; TTE, transthoracic echocardiography.')

# ================================================================ Table S1B
tA1 = table_after(hits('A. Baseline characteristics', exact=True)[0])
assert TXT(cell_par(rows(tA1)[0], 1)).startswith('Included (844 procedures)')
set_cell(rows(tA1)[0], 2, 'No 1-month TTE (601 admissions)')
with open(R57.replace('R57_joint_12factor.csv', 'R56_C_S1B_included_vs_not.csv'), encoding='utf-8') as f:
    C = {r['variable']: r for r in csv.DictReader(f)}
LAB = {'Age, y': 'Age, y', 'Female sex': 'Female sex', 'Persistent AF': 'Persistent AF', 'LA diameter, mm': 'LA diameter, mm', 'RA diameter, mm': 'RA diameter, mm',
       'RV diameter, mm': 'RV diameter, mm', 'LVEF, %': 'LVEF, %', 'MR grade ≥2': 'MR grade >=2', 'TR grade ≥2': 'TR grade >=2', 'isolated MR': 'isolated MR',
       'isolated TR': 'isolated TR', 'combined MR/TR': 'combined MR/TR'}
smd = lambda s: s.replace('-', '−') if s.startswith('-') else s
fmt_smd = lambda v: smd(f'{float(v):.2f}')
for tr in rows(tA1)[1:]:
    k = TXT(cell_par(tr, 0)).strip(); r = C[LAB[k]]
    assert TXT(cell_par(tr, 1)) == r['included'].replace('-', '−') or TXT(cell_par(tr, 1)) == r['included'], (k, TXT(cell_par(tr, 1)), r['included'])
    set_cell(tr, 2, r['not_included']); set_cell(tr, 3, fmt_smd(r['SMD']))
tB1 = table_after(hits('B. ≥1-grade improvement at 1 month, unweighted and weighted', )[0])
VW = {'Isolated MR': '48.2 (39.9–57.1)', 'Isolated TR': '51.3 (46.0–56.7)', 'Combined disease: MR': '47.0 (41.9–52.2)', 'Combined disease: TR': '53.6 (48.5–58.5)'}
for tr in rows(tB1)[1:]:
    set_cell(tr, 3, VW[TXT(cell_par(tr, 0)).strip()])
rev(one('Eligible admissions: 1,471 admissions in 1,446 patients'),
    [('Eligible admissions: 1,471 admissions in 1,446 patients;', 'Eligible admissions: 1,445 admissions in 1,419 patients of the study population with MR or TR of analysis grade ≥1;'),
     ('for 553 of 627 admissions without a 1-month TTE (not recorded or unspecified in 74)', 'for 553 of 601 admissions without a 1-month TTE (not recorded or unspecified in 48)'),
     ('in all 1,471 eligible admissions', 'in all 1,445 eligible admissions'),
     ('(AF type, n=74; LA, RA, and LVEF, ≤70 each)', '(AF type, n=48; LA, RA, and LVEF, ≤64 each)'),
     ('(C statistic, 0.68)', '(C statistic, 0.66)'), ('(range, 0.76–1.12)', '(range, 0.80–1.06)'),
     ('observed rate − 0.426 × δ, where 0.426 is', 'observed rate − 0.416 × δ, where 0.416 is'),
     ('43.5% and 39.2% (isolated MR, MR), 47.1% and 42.9% (isolated TR, TR), 42.8% and 38.5% (combined MR/TR, MR), and 49.5% and 45.2% (combined MR/TR, TR)',
      '43.6% and 39.4% (isolated MR, MR), 47.2% and 43.1% (isolated TR, TR), 42.9% and 38.7% (combined MR/TR, MR), and 49.6% and 45.4% (combined MR/TR, TR)')])

# ================================================================ Table S11 part A
s11 = table_after(hits('Table S11. Paroxysmal and persistent AF analyzed separately')[1])
V11 = {'Patients, n': ['1,848', '1,197', ''], 'Both valves affected, n (%)': ['215 (11.6)', '448 (37.4)', ''],
       'Observed/expected co-occurrence, conditional on age and sex*': ['1.56 (1.43–1.67)', '1.18 (1.14–1.22)', '<0.001'],
       'OR for TR given MR, adjusted for age, sex, LA, RA, LV end-diastolic diameter, and LVEF': ['2.91 (2.14–3.96)', '2.83 (2.00–4.01)', '0.91'],
       'LA diameter, both valves vs neither, adjusted difference, mm (95% CI)': ['+5.3 (4.4 to 6.2)', '+3.2 (2.4 to 4.1)', '<0.001'],
       'RA diameter, both valves vs neither, adjusted difference, mm (95% CI)': ['+6.4 (5.4 to 7.3)', '+6.3 (5.4 to 7.3)', '0.97']}
seen = set()
for tr in rows(s11):
    k = TXT(cell_par(tr, 0)).strip()
    if k.startswith('B. '): break
    if k in V11:
        for i, x in enumerate(V11[k]): set_cell(tr, i + 1, x)
        seen.add(k)
assert seen == set(V11), set(V11) - seen
rev(cell_par([tr for tr in rows(s11) if TXT(cell_par(tr, 0)).startswith('B. Paired cohort')][0], 0), [('B. Paired cohort:', 'B. Paired response set:')])

# ================================================================ new Table S13 (after the Table S12 note)
s12t = hits('Table S12. Aortic regurgitation as a negative control outcome')[1]
s12note = one('AR was abstracted from the text of the baseline TTE report')
title = copy.deepcopy(s12t._p); settext(title, 'Table S13. Baseline factors associated with improvement of both valves in combined MR/TR (post hoc)')
src = s11  # 4-column template
new = copy.deepcopy(src); R = rows(new); hdr, dat = copy.deepcopy(R[0]), copy.deepcopy(R[2]); last_tcpr = [copy.deepcopy(c.find(qn('w:tcPr'))) for c in cells(R[-1])]
for r in R: new.remove(r)
rows57 = list(csv.DictReader(open(R57, encoding='utf-8')))
def qf(q):
    q = float(q); return '<0.001' if q < 0.001 else (f'{q:.3f}' if q < 0.01 else f'{q:.2f}')
def est(r, k): return f"{float(r[k+'_OR']):.2f} ({float(r[k+'_lo']):.2f}–{float(r[k+'_hi']):.2f}); q={qf(r[k+'_q'])}"
CON = {'Age, y': 'Age, 75 vs 61 years', 'Female sex': 'Female sex', 'BMI, kg/m²': 'BMI, 25.9 vs 22.0 kg/m²', 'Persistent AF': 'Persistent AF',
       'TR peak velocity, m/s': 'TRV, 3.0 vs 2.4 m/s', 'LA diameter, mm': 'LA diameter, 47 vs 40 mm', 'RA diameter, mm': 'RA diameter, 48 vs 40 mm',
       'LV end-diastolic diameter, mm': 'LV end-diastolic diameter, 50 vs 45 mm', 'RV diameter, mm': 'RV diameter, 24 vs 21 mm', 'LVEF, %': 'LVEF, 68 vs 56%',
       'Baseline MR grade ≥2 (vs 1)': 'Baseline MR grade ≥2 vs 1', 'Baseline TR grade ≥2 (vs 1)': 'Baseline TR grade ≥2 vs 1'}
r0 = copy.deepcopy(hdr); set_row(r0, ['Factor (contrast)', 'Both valves improved vs one or neither (n=387; 133 events), OR (95% CI); q',
                                      'Number of improved valves, 0/1/2 (n=387; 130/124/133), OR (95% CI); q',
                                      'Extension to both valves among procedures with ≥1 improved valve (n=257; 133 events), OR (95% CI); q']); new.append(r0)
for r in rows57:
    x = copy.deepcopy(dat); set_row(x, [CON[r['factor']], est(r, 'joint'), est(r, 'ordinal'), est(r, 'extension')]); new.append(x)
for c, pr in zip(cells(rows(new)[-1]), last_tcpr):
    if pr is not None:
        o = c.find(qn('w:tcPr'))
        if o is not None: c.remove(o)
        c.insert(0, pr)
TW = sum(int(g.get(qn('w:w'))) for g in new.find(qn('w:tblGrid')).findall(qn('w:gridCol'))); c0 = int(TW * 0.31); c1 = (TW - c0) // 3
set_widths(new, [c0, c1, c1, TW - c0 - 2 * c1])
note = copy.deepcopy(s12note._p); settext(note,
    'Combined MR/TR procedures (n=387). Each outcome was modeled with the 12 baseline factors of the main model (Table S6), with baseline MR and TR grades (≥2 vs 1) in place of phenotype, '
    'which is constant in combined disease; missing values were imputed by predictive mean matching (m=20) as for the main model and estimates pooled by Rubin’s rules. Continuous factors '
    'are compared at the 75th versus 25th percentile within combined MR/TR. Both valves improved: logistic regression of improvement of both valves versus one or neither. Number of improved '
    'valves: proportional-odds logistic regression (0, 1, or 2 valves improved); ORs >1 indicate more improved valves. Extension to both valves: logistic regression of both versus one valve '
    'improved among the 257 procedures with at least one improved valve. Logistic models used patient-clustered robust standard errors. q, Benjamini–Hochberg false discovery rate across '
    'the 12 factors within each model. These analyses were post hoc. AF indicates atrial fibrillation; BMI, body mass index; LA, left atrial; LV, left ventricular; LVEF, left ventricular '
    'ejection fraction; MR, mitral regurgitation; OR, odds ratio; RA, right atrial; RV, right ventricular; TR, tricuspid regurgitation; TRV, peak tricuspid regurgitant velocity.')
s12note._p.addnext(title); title.addnext(new); new.addnext(note)
for e in (title, note): tc.mark(e, 'ins')
tc.mark(new, 'ins')

# ================================================================ Figure S2 heading and legend (after the Figure S1 legend)
f1h = hits('Figure S1. Weighted comparisons of improvement between isolated and combined regurgitation')[1]
f1leg = one('(A) Odds ratios (95% CIs) for ≥1-grade improvement in isolated versus combined disease under IPTW-ATT')
h2 = copy.deepcopy(f1h._p); settext(h2, 'Figure S2. Standardized probabilities of TR improvement across RA and LV end-diastolic diameters')
l2 = copy.deepcopy(f1leg._p); settext(l2,
    'Standardized probabilities (95% CIs) of ≥1-grade TR improvement at 1 month derived by g-computation from the 12-factor valve-level model (main Figure 4A; Table S6) across RA (A) and '
    'LV end-diastolic (B) diameters; chamber diameters are unindexed. These curves were previously shown in main Figure 4C–D. LV, left ventricular; RA, right atrial; TR, tricuspid regurgitation.')
f1leg._p.addnext(h2); h2.addnext(l2)
for e in (h2, l2): tc.mark(e, 'ins')

# ================================================================ global wording
N = 0
for p in allp():
    for a, b in (('baseline phenotype cohort', 'baseline phenotype set'), ('Baseline phenotype cohort', 'Baseline phenotype set'),
                 ('paired response cohort', 'paired response set'), ('Paired response cohort', 'Paired response set'),
                 ('the paired cohort', 'the paired response set'), ('Paired cohort', 'Paired response set'), ('all paired-cohort analyses', 'all paired-set analyses')):
        N += tracked_sub(p._p, a, b)
d.save(OUT); print('saved', OUT, '| cohort->set substitutions', N)
