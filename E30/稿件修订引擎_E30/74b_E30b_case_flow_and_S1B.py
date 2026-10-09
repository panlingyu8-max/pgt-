# -*- coding: utf-8 -*-
"""74b. (E30b, 2026-10-09) 病例（消融住院）为单位的统一流程与 Table S1B。取代 74 号脚本。在项目根目录运行，先运行 63 号脚本无需单独执行（本脚本执行其前段）。
流程：AF 消融住院 → 有同次住院基线 TTE 分级（电子档案文本解析，或研究登记补录且有 1 个月心超）→ 统一排除 → 基线表型集
      → MR 和/或 TR ≥1 级 → 无 1 个月 TTE（表 S1B 比较对象）→ 配对反应分析集（844 次）。
S1B：配对集 vs 基线集中 MR/TR ≥1 级但无 1 个月 TTE 的住院；随访逆概率加权改善率；delta 系数 = 未随访住院比例。
输出（汇总，无标识符）：11_顶刊修订分析/tables/R59_A_case_flow.csv、R59_B_exclusions.csv、R59_C_S1B_included_vs_not.csv、R59_D_S1B_ipw_rates.csv"""
import numpy as np, pandas as pd
SRC = '05_统计代码/可移植版/顶刊修订/'; TAB = '11_顶刊修订分析/tables/'; RAW = '09_原始数据_含直接标识符/'
src = open(SRC + '63_baseline_phenotype_analysis.py', encoding='utf-8').read().split('# ---------- 1. availability')[0]
g = {'__name__': 'r63'}; exec(src, g)
IA, I, REV = g['I_ALL'], g['I'], g['REV']
T = I[I.text].copy()
T['anyreg'] = (T.MRp + T.TRp) > 0

# ---------------- A. flow (cases = ablation admissions)
ex_mask = IA.inset & (IA.excl_valve_surgery | IA.excl_team_review) & ~IA.reinstated
R1 = T[T.anyreg]; paired = R1[R1.in844]; notp = R1[~R1.in844]
flow = [('AF ablation admissions, Jan 2020 to Sep 2024', len(IA), IA.uid.nunique()),
        ('  with a baseline TTE grade from the same admission', int(IA.inset.sum()), IA[IA.inset].uid.nunique()),
        ('    of which electronic report text', int((IA.inset & ~IA.reg_src).sum()), np.nan),
        ('    of which registry abstraction (manual entry)', int(IA.reg_src.sum()), np.nan),
        ('  excluded (uniform criteria)', int(ex_mask.sum()), IA[ex_mask].uid.nunique()),
        ('baseline phenotype set', len(T), T.uid.nunique()),
        ('  PAF / PeAF / unclassified', f"{int((T.af_type == 'PAF').sum())} / {int((T.af_type == 'PeAF').sum())} / {int(T.af_type.isna().sum())}", np.nan),
        ('  MR and/or TR >=1', len(R1), R1.uid.nunique()),
        ('    no 1-month TTE', int((~notp.has1m).sum()), notp[~notp.has1m].uid.nunique()),
        ('    1-month TTE but not in paired set (check)', int(notp.has1m.sum()), np.nan),
        ('    paired response set', len(paired), paired.uid.nunique()),
        ('paired procedures overall (registry groups 1-3) / matched to archive list', int(pd.to_numeric(g['_R'].iloc[:, 5], errors='coerce').isin([1, 2, 3]).sum()), int(IA.in844.sum()))]
F = pd.DataFrame(flow, columns=['step', 'cases', 'patients']); F.to_csv(TAB + 'R59_A_case_flow.csv', index=False); print(F.to_string(index=False))
miss = IA[IA.in844 & ~IA.index.isin(T.index)]
print('paired procedures matched but not in baseline set:', len(miss), miss[['inset', 'excl_valve_surgery', 'excl_team_review', 'reinstated']].sum().to_dict() if len(miss) else '')

# ---------------- B. exclusion categories (cases)
rv = REV.assign(u=REV['患者唯一号'].astype(str).str.strip()).drop_duplicates('u').set_index('u')['初筛理由']
cat = np.where(IA.loc[ex_mask, 'excl_valve_surgery'], 'Prior or concurrent valve surgery', IA.loc[ex_mask, 'uid'].map(rv).fillna('unknown'))
pd.Series(cat).value_counts().to_csv(TAB + 'R59_B_exclusions.csv'); print(pd.Series(cat).value_counts())

# ---------------- C. S1B: paired set vs MR/TR >=1 without 1-month TTE
E = pd.concat([paired.assign(included=1), notp[~notp.has1m].assign(included=0)])
E['female'] = E.female.astype(float); E['persistent'] = np.select([E.af_type == 'PeAF', E.af_type == 'PAF'], [1.0, 0.0], np.nan)
E['MR_ge_mod'] = (E.MRg >= 2).astype(float); E['TR_ge_mod'] = (E.TRg >= 2).astype(float)
E['pheno'] = np.select([(E.MRp == 1) & (E.TRp == 1), E.MRp == 1, E.TRp == 1], ['C', 'A', 'B'], 'NA')
def smd(x1, x0, b):
    x1, x0 = x1.dropna(), x0.dropna(); p1, p0 = x1.mean(), x0.mean()
    s = np.sqrt((p1 * (1 - p1) + p0 * (1 - p0)) / 2) if b else np.sqrt((x1.var() + x0.var()) / 2); return (p1 - p0) / s
inc, exc = E[E.included == 1], E[E.included == 0]; t = []
for v, lab, b in [('age', 'Age, y', 0), ('female', 'Female sex', 1), ('persistent', 'Persistent AF', 1), ('LA', 'LA diameter, mm', 0), ('RA', 'RA diameter, mm', 0),
                  ('RV', 'RV diameter, mm', 0), ('LVEF', 'LVEF, %', 0), ('MR_ge_mod', 'MR grade >=2', 1), ('TR_ge_mod', 'TR grade >=2', 1)]:
    f = (lambda x: f'{100 * x.mean():.1f}%') if b else (lambda x: f'{x.mean():.1f} ± {x.std():.1f}')
    t.append(dict(variable=lab, included=f(inc[v].dropna()), not_included=f(exc[v].dropna()), SMD=round(smd(inc[v], exc[v], b), 2), n_not_included=int(exc[v].notna().sum())))
for ph, lab in (('A', 'isolated MR'), ('B', 'isolated TR'), ('C', 'combined MR/TR')):
    x1, x0 = (inc.pheno == ph).astype(float), (exc.pheno == ph).astype(float)
    t.append(dict(variable=lab, included=f'{100 * x1.mean():.1f}%', not_included=f'{100 * x0.mean():.1f}%', SMD=round(smd(x1, x0, 1), 2), n_not_included=len(exc)))
T7 = pd.DataFrame(t); T7.to_csv(TAB + 'R59_C_S1B_included_vs_not.csv', index=False); print(T7.to_string(index=False))
print('S1B: included', len(inc), 'cases /', inc.uid.nunique(), 'patients; not included', len(exc), 'cases /', exc.uid.nunique(), 'patients; AF type missing among not included',
      int(exc.persistent.isna().sum()), '| delta coefficient', round(len(exc) / len(E), 3))

# ---------------- D. IPW for 1-month follow-up; outcomes of included procedures from the registry (same row via 病案号)
RB = g['_RB']; mk = g['_mk']; c = lambda p: pd.to_numeric(RB.iloc[:, p - 1], errors='coerce')
O = pd.DataFrame({'mk': mk(RB.iloc[:, 2]), 'grp': c(6), 'mr_b': c(93), 'tr_b': c(95), 'mr_1': c(157), 'tr_1': c(158)}).drop_duplicates('mk').set_index('mk')
Z = E.copy(); covs = ['age', 'female', 'persistent', 'LA', 'LVEF', 'RA']
for v in covs: Z[v + '_miss'] = Z[v].isna().astype(float); Z[v] = Z[v].fillna(Z[v].median())
Z['phB'] = (Z.pheno == 'B').astype(float); Z['phC'] = (Z.pheno == 'C').astype(float); Z['MRg2'] = Z.MR_ge_mod; Z['TRg2'] = Z.TR_ge_mod
Xc = covs + [v + '_miss' for v in covs if Z[v + '_miss'].sum() > 0] + ['phB', 'phC', 'MRg2', 'TRg2']
X = np.column_stack([np.ones(len(Z))] + [((Z[k] - Z[k].mean()) / (Z[k].std() or 1)).values for k in Xc]); y = Z.included.values.astype(float)
b = np.zeros(X.shape[1])
for _ in range(50):
    p = 1 / (1 + np.exp(-X @ b)); W = p * (1 - p); st = np.linalg.solve(X.T @ (X * W[:, None]), X.T @ (y - p)); b += st
    if np.abs(st).max() < 1e-9: break
ps = 1 / (1 + np.exp(-X @ b)); o = ps.argsort(); rk = np.empty(len(ps)); rk[o] = np.arange(1, len(ps) + 1)
cstat = (rk[y == 1].sum() - y.sum() * (y.sum() + 1) / 2) / (y.sum() * (len(y) - y.sum()))
w = np.where(y == 1, 1 / ps, 0); lo_, hi_ = np.percentile(w[y == 1], [1, 99]); Z['w'] = np.clip(w, lo_, hi_) * y.mean()
for k in ('grp', 'mr_b', 'tr_b', 'mr_1', 'tr_1'): Z[k] = Z.mk.map(O[k])
rng = np.random.default_rng(20260924); t8 = []
for k, (gg, b_, a_) in {'Isolated MR': (1, 'mr_b', 'mr_1'), 'Isolated TR': (2, 'tr_b', 'tr_1'), 'Combined disease: MR': (3, 'mr_b', 'mr_1'), 'Combined disease: TR': (3, 'tr_b', 'tr_1')}.items():
    x = Z[(Z.included == 1) & (Z.grp == gg)]; ev = (x[a_] < x[b_]).values.astype(float); ww = x.w.values
    bs = [(ev[s] * ww[s]).sum() / ww[s].sum() for s in (rng.integers(0, len(x), len(x)) for _ in range(2000))]
    t8.append(dict(series=k, n=len(x), crude=round(100 * ev.mean(), 1), weighted=round(100 * (ev * ww).sum() / ww.sum(), 1), lo=round(np.percentile(bs, 2.5) * 100, 1), hi=round(np.percentile(bs, 97.5) * 100, 1)))
T8 = pd.DataFrame(t8); T8['c_stat'] = round(cstat, 2); T8['weight_range'] = f"{Z.loc[Z.included == 1, 'w'].min():.2f}–{Z.loc[Z.included == 1, 'w'].max():.2f}"
T8['delta_coefficient'] = round(len(exc) / len(E), 4)
T8.to_csv(TAB + 'R59_D_S1B_ipw_rates.csv', index=False); print(T8.to_string(index=False))
assert list(T8.n) == [132, 325, 387, 387], list(T8.n)
print('missing among not included:', {v: int(exc[v].isna().sum()) for v in ('persistent', 'LA', 'RA', 'LVEF')})
