# -*- coding: utf-8 -*-
"""63. 【E30b，2026-10-09 作者决定】分析单位改为病例（消融住院），不区分首次或重复消融。基线分级均取自同次住院的 TTE 报告：电子档案能导出的用文本解析分级，未能导出、由课题组补录到研究登记表的用登记分级。登记表二尖瓣、三尖瓣分级均为空白 = 无基线心超；无电子报告且无 1 个月心超的登记住院不纳入；有 1 个月心超的登记住院以登记分级为准（配对集 844 次因此全部在基线集内）。不再做报告可得性加权。
63. 【E30，2026-10-08 作者决定】分析集改为同次住院有电子档案 TTE 报告者：有手术日期时取手术日或之前的报告，无手术日期时取该次住院最早的一份报告（62 号脚本规则；判断依据只记录于内部核实文件）；配对队列患者由团队评估纳入，不再被结构性病变复核排除（9 例）。以下为原说明。
63. 基线表型分析（补导出前，可用数据版；2026-10-03 作者决定：只纳入能确认为术前的报告——总表有手术日期且报告日期不晚于手术日（同日报告视为术前），无手术日期者不纳入；与配对人群重叠者以配对分析所用基线分级为准）。排除规则与配对人群一致：主分析排除既往或同期瓣膜置换/修复，以及团队复核判定的器质性/结构性病变（2026-10-03，68 例）；64a 严格规则作敏感性分析。在项目根目录运行。输入：62 号脚本输出 + 数据.zip 中的心动超声（心腔径线）。
分析集：每例患者首次消融住院（索引住院）中，同一次住院有术前 TTE 文本者（报告解析分级；与是否有反流无关）。
  总表已提取、无文本的 476 次住院因按有反流选择提取，只用于敏感性分析。缺失不视为无反流；以有无术前 TTE 文本的逆概率加权作敏感性分析。
表型：反流存在 = 8 级 ≥1（极轻度及以上，与现分析队列一致）；微量记 0。
输出（汇总，不含标识符）：11_顶刊修订分析/tables/R48_*.csv；11_顶刊修订分析/figures/R48_baseline_phenotype.png"""
import io, zipfile, numpy as np, pandas as pd
RAW = '09_原始数据_含直接标识符/'; TAB = '11_顶刊修订分析/tables/'; FIG = '11_顶刊修订分析/figures/'
rng = np.random.default_rng(20261002)
X = pd.read_excel(RAW + '基线TTE两瓣分级_全部消融住院_20261002.xlsx', dtype={'就诊ID': str, '登记号': str, '病案号': str, '患者唯一号': str})
z = zipfile.ZipFile(RAW + '数据.zip')
E = [pd.read_excel(io.BytesIO(z.read(i))) for i in z.infolist() if i.filename.encode('cp437').decode('gbk').endswith('心动超声.csv')][0]
norm = lambda s: s.astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
E['vid'] = norm(E['就诊ID']); E['dt'] = pd.to_datetime(E['报告日期'], errors='coerce')
for c, n in (('LA', 'LA'), ('RA', 'RA'), ('LV', 'LVEDD'), ('RV', 'RV'), ('EF＇', 'LVEF')):
    E[n] = pd.to_numeric(E[c], errors='coerce')
X['vid'] = norm(X['就诊ID']); X['tte_dt'] = pd.to_datetime(X['术前TTE报告时间'], errors='coerce')
# 按就诊ID + 报告日期（精确到分钟）匹配同一份报告（Excel 往返会丢失秒以下精度）
E['k'] = E.dt.dt.floor('min'); X['k'] = X.tte_dt.dt.floor('min')
X = X.merge(E[['vid', 'k', 'LA', 'RA', 'LVEDD', 'RV', 'LVEF']].drop_duplicates(['vid', 'k']), on=['vid', 'k'], how='left')
print('chamber data matched:', X.LA.notna().sum(), 'of', X.tte_dt.notna().sum())
for c, lo, hi in (('LA', 15, 90), ('RA', 15, 100), ('LVEDD', 25, 90), ('RV', 10, 60), ('LVEF', 10, 90)):
    X.loc[(X[c] < lo) | (X[c] > hi), c] = np.nan
X['age'] = pd.to_numeric(X['年龄'], errors='coerce'); X['female'] = (X['性别'].astype(str).str.contains('女')).astype(float)
X['year'] = pd.to_datetime(X['出院时间']).dt.year; X['laao'] = X['手术名称'].astype(str).str.contains('左心耳').astype(float)
X['PeAF'] = (X.af_type == 'PeAF').astype(float); X['AFunk'] = X.af_type.isna().astype(float)
# E28：71 号脚本对分析集中 86 例未分型者补充了 AF 类型；报告可得性模型只用总表/出院诊断记录的 AF 类型（两组均可获得，避免完全分离）
_reg = X.af_type.where(~X.af_type_source.astype(str).str.contains('A峰'))
X['PeAF_reg'] = (_reg == 'PeAF').astype(float); X['AFunk_reg'] = _reg.isna().astype(float)
X['in844'] = X['现分析分组(1A/2B/3C)'].isin([1, 2, 3])
I = X.copy()                                   # E30b：全部消融住院（病例为单位），不再只取每例首次消融住院
# ---------- 排除规则（与配对人群一致）：主分析排除既往或同期瓣膜置换/修复；严格规则（64a）作敏感性分析
import importlib.util as _ilu
_sp = _ilu.spec_from_file_location('rule', '05_统计代码/可移植版/顶刊修订/64a_structural_exclusion_rule.py'); rule = _ilu.module_from_spec(_sp); _sp.loader.exec_module(rule)
HB = [pd.read_excel(io.BytesIO(z.read(i))) for i in z.infolist() if i.filename.encode('cp437').decode('gbk').endswith('病案首页表.csv')][0]
HB['mrn'] = HB['病案号'].astype(str).str.strip(); HB['dx'] = HB[[c for c in HB.columns if '诊断' in c]].astype(str).agg(' | '.join, axis=1)
I['mrn'] = I['病案号'].astype(str).str.strip(); I = I.merge(HB[['mrn', 'dx']].drop_duplicates('mrn'), on='mrn', how='left').set_index(I.index)
_K = pd.DataFrame([rule.classify(t) for t in (I['诊断意见'].fillna('').astype(str) + ' | ' + I.dx.fillna('') + ' | ' + I['手术名称'].fillna('').astype(str))], index=I.index)
I['excl_valve_surgery'] = _K['prosthesis_or_valve_repair']; I['excl_strict'] = _K.any(axis=1)
EXCL = I[['grade_source', 'excl_valve_surgery', 'excl_strict']].copy(); EXCL_CATS = _K[I.grade_source == 'TTE文本解析'].sum()
# 团队复核（2026-10-03，方案 A，与 844 次筛选标准一致）：报告结论或出院诊断提示结构性病变的 135 例，按 09/基线人群_结构性病变待复核清单_20261003.xlsx 的“团队判断”排除 68 例
REV = pd.read_excel(RAW + '基线人群_结构性病变待复核清单_20261003.xlsx', dtype=str)
REV_EX = set(REV.loc[REV['团队判断(排除/保留)'] == '排除', '患者唯一号'].str.strip())
I['excl_team_review'] = I['患者唯一号'].astype(str).str.strip().isin(REV_EX) & ~I.excl_valve_surgery
REV_CATS = REV[REV['团队判断(排除/保留)'] == '排除']['初筛理由'].value_counts()
_nrev = I.loc[I.excl_team_review, '患者唯一号'].astype(str).str.strip().nunique()   # E30b：病例为单位，一例患者可有多次住院
assert _nrev == len(REV_EX), (_nrev, len(REV_EX))
I['op_dt'] = pd.to_datetime(I['手术时间(总表)'], errors='coerce')
I['confirmed'] = (I.grade_source == 'TTE文本解析') & I.op_dt.notna()          # 62 号脚本在有手术日期时只取手术日或之前的报告
# 与配对人群重叠者：以配对分析所用的基线分级（总表 = 06 分析数据 mrr_b/trr_b，已核对一致）为准
I['archive'] = I.grade_source == 'TTE文本解析'                                   # E30
_ov = I.archive & I.in844 & I.tab_MR8.notna() & I.tab_TR8.notna()
N_OVERRIDE_CHANGED = int((_ov & ((I.MR8 != I.tab_MR8) | (I.TR8 != I.tab_TR8))).sum())
I.loc[_ov, 'MR8'] = I.loc[_ov, 'tab_MR8']; I.loc[_ov, 'TR8'] = I.loc[_ov, 'tab_TR8']
# ---------- E30b：登记补录分级（有 1 个月心超、二尖瓣或三尖瓣至少一项有分级）；登记中另一瓣空白记 0
_RB = pd.read_excel(RAW + '出院时间2020-1-2024-9总表改2.xlsx')
_mk = lambda x: x.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
_c = lambda p: pd.to_numeric(_RB.iloc[:, p - 1], errors='coerce')
_lvcol = next((j for j, n in enumerate(_RB.columns) if any(k in str(n) for k in ('左室舒张末', 'LVEDD', 'LVDd', 'LVIDd'))), None)
_B = pd.DataFrame({'mk': _mk(_RB.iloc[:, 2]), 'has1m': _c(157).notna() | _c(158).notna(), 'r_LA': _c(105), 'r_RV': _c(106), 'r_RA': _c(107), 'r_LVEF': _c(127),
                   'r_LVEDD': pd.to_numeric(_RB.iloc[:, _lvcol], errors='coerce') if _lvcol is not None else np.nan})
_B = _B[_B.mk.ne('nan')].drop_duplicates('mk').set_index('mk')
print('E30b: registry LVEDD column', _RB.columns[_lvcol] if _lvcol is not None else 'not found (LVEDD only from archive reports)')
I['mk'] = _mk(I['病案号'])
I['has1m'] = I.mk.map(_B.has1m).fillna(False).astype(bool)
I['reg_src'] = (I.tab_MR8.notna() | I.tab_TR8.notna()) & I.has1m
I.loc[I.reg_src, 'MR8'] = I.loc[I.reg_src, 'tab_MR8'].fillna(0); I.loc[I.reg_src, 'TR8'] = I.loc[I.reg_src, 'tab_TR8'].fillna(0)
for c, lo, hi in (('LA', 15, 90), ('RA', 15, 100), ('LVEDD', 25, 90), ('RV', 10, 60), ('LVEF', 10, 90)):   # 无电子报告者用登记中的心腔径线
    v = I.mk.map(_B['r_' + c]); v = v.where((v >= lo) & (v <= hi)); I[c] = I[c].fillna(v)
I['inset'] = I.reg_src | (I.archive & I.MR8.notna() & I.TR8.notna())
I['grade_from'] = np.select([I.reg_src, I.inset], ['登记补录', '电子档案文本'], '无')
_m, _t = I.MR8 >= 1, I.TR8 >= 1
I['phenotype'] = np.select([I.MR8.isna() | I.TR8.isna(), _m & _t, _m, _t], ['缺失', '双瓣共存', '仅MR', '仅TR'], '无MR/TR')
# E30：配对队列（844 次）中的患者经团队补录与评估后纳入，在基线集中同样保留
_R = pd.read_excel(RAW + '出院时间2020-1-2024-9总表改2.xlsx'); _g = pd.to_numeric(_R.iloc[:, 5], errors='coerce')
_nu = lambda x: x.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
P844_UID = set(_nu(_R[_g.isin([1, 2, 3])].iloc[:, 6]))
I['uid'] = _nu(I['患者唯一号']); I['reinstated'] = I.in844 & (I.excl_valve_surgery | I.excl_team_review)   # E30b：按病例恢复配对手术
I_ALL = I.copy(); I = I[(~I.excl_valve_surgery & ~I.excl_team_review) | I.reinstated].copy()
I['text'] = I.inset                            # E30b：基线表型集标志（下游 66、72、74 号脚本沿用 I.text）
print('E30b baseline set (cases):', int(I.text.sum()), '| by source', I[I.text].grade_from.value_counts().to_dict(), '| paired procedures in set', int((I.text & I.in844).sum()), 'of', int(I_ALL.in844.sum()), 'matched in archive list',
      '| AF type unclassified in set', int(I[I.text].af_type.isna().sum()))
MRmap = lambda g: np.select([g == 0, g <= 2, g <= 4], [0, 1, 2], 3); TRmap = lambda g: np.select([g == 0, g <= 2, g <= 4, g <= 6], [0, 1, 2, 3], 4)
I['MRg'] = np.where(I.MR8.notna(), MRmap(I.MR8.fillna(0)), np.nan); I['TRg'] = np.where(I.TR8.notna(), TRmap(I.TR8.fillna(0)), np.nan)
I['MRp'] = (I.MR8 >= 1).astype(float); I['TRp'] = (I.TR8 >= 1).astype(float)
I['both'] = I.MRp * I.TRp

# ---------- small GLM helpers (numpy IRLS; sandwich SE)
def design(d, cols):
    M = np.column_stack([np.ones(len(d))] + [d[c].values.astype(float) for c in cols]); return M
def logit(y, M, w=None, it=50):
    w = np.ones(len(y)) if w is None else w; b = np.zeros(M.shape[1])
    for _ in range(it):
        eta = M @ b; p = 1 / (1 + np.exp(-eta)); W = w * p * (1 - p)
        H = M.T @ (M * W[:, None]); g = M.T @ (w * (y - p)); step = np.linalg.solve(H, g); b += step
        if np.abs(step).max() < 1e-9: break
    p = 1 / (1 + np.exp(-(M @ b))); Hinv = np.linalg.inv(M.T @ (M * (w * p * (1 - p))[:, None]))
    S = M * (w * (y - p))[:, None]; V = Hinv @ (S.T @ S) @ Hinv
    return b, np.sqrt(np.diag(V)), p
def poisson(y, M, it=100):
    b = np.zeros(M.shape[1]); b[0] = np.log(max(y.mean(), 1e-6))
    for _ in range(it):
        mu = np.exp(M @ b); H = M.T @ (M * mu[:, None]); step = np.linalg.solve(H, M.T @ (y - mu)); b += step
        if np.abs(step).max() < 1e-10: break
    mu = np.exp(M @ b); Hinv = np.linalg.inv(M.T @ (M * mu[:, None])); S = M * (y - mu)[:, None]; V = Hinv @ (S.T @ S) @ Hinv
    return b, np.sqrt(np.diag(V))
def ols(y, M):
    b = np.linalg.lstsq(M, y, rcond=None)[0]; r = y - M @ b; Hinv = np.linalg.inv(M.T @ M); S = M * r[:, None]
    V = Hinv @ (S.T @ S) @ Hinv * len(y) / (len(y) - M.shape[1]); return b, np.sqrt(np.diag(V))
def wilson(x, n, z=1.959964):
    if n == 0: return (np.nan, np.nan)
    p = x / n; d = 1 + z * z / n; c = (p + z * z / (2 * n)) / d; h = z * np.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d; return (100 * (c - h), 100 * (c + h))
f1 = lambda x: f'{x:.1f}'; f2 = lambda x: f'{x:.2f}'

# ---------- 1. availability of pre-op TTE text (selection) and IPW
I['ycat'] = I.year.astype(int)
yrs = sorted(I.ycat.unique())[1:]
for y_ in yrs: I[f'y{y_}'] = (I.ycat == y_).astype(float)
cov_sel = ['age', 'female', 'PeAF_reg', 'AFunk_reg', 'laao'] + [f'y{y_}' for y_ in yrs]
S = I.dropna(subset=['age']).copy()
bs, ses, ps = logit(S.text.astype(float).values, design(S, cov_sel))
pt = S.text.mean(); w = np.where(S.text, pt / ps, np.nan); lo_, hi_ = np.nanpercentile(w, [1, 99]); S['ipw'] = np.clip(w, lo_, hi_)
I = I.join(S[['ipw']])
c_stat = (lambda y, p: (np.argsort(np.argsort(p))[y == 1].mean() - (y.sum() - 1) / 2) / (len(y) - y.sum()))(S.text.astype(int).values, ps)
sel = []
for lab, d in (('with confirmed pre-procedural report', I[I.text]), ('without confirmed pre-procedural report', I[~I.text]), ('  of which: report but no procedure date', I[(I.grade_source == 'TTE文本解析') & ~I.text]), ('  of which: no report in the same admission', I[I.grade_source != 'TTE文本解析'])):
    sel.append(dict(group=lab, n=len(d), age_mean=d.age.mean(), female_pct=100 * d.female.mean(), PeAF_pct=100 * (d.af_type == 'PeAF').mean(), PAF_pct=100 * (d.af_type == 'PAF').mean(),
                    AF_type_unknown_pct=100 * d.af_type.isna().mean(), LAAO_pct=100 * d.laao.mean(), in_analysis_cohort_pct=100 * d.in844.mean(),
                    **{f'year_{y_}_pct': 100 * (d.ycat == y_).mean() for y_ in sorted(I.ycat.unique())}))
sel = pd.DataFrame(sel); sel['note'] = f'C statistic of availability model {c_stat:.2f}; IPW truncated 1st-99th percentile'
sel.to_csv(TAB + 'R48_A_text_availability.csv', index=False)

# ---------- 2. phenotype distribution (text set; IPW; sensitivity incl. table-graded)
T = I[I.text].copy()
PH = ['无MR/TR', '仅MR', '仅TR', '双瓣共存']; PHE = {'无MR/TR': 'Neither', '仅MR': 'MR only', '仅TR': 'TR only', '双瓣共存': 'Both'}
rows = []
def dist(d, lab, wcol=None):
    for ph in PH:
        x = (d.phenotype == ph); n = len(d)
        if wcol is None:
            lo, hi = wilson(x.sum(), n); rows.append(dict(analysis=lab, phenotype=PHE[ph], n=int(x.sum()), N=n, pct=100 * x.mean(), lo=lo, hi=hi))
        else:
            ww = d[wcol].values; rows.append(dict(analysis=lab, phenotype=PHE[ph], n=int(x.sum()), N=n, pct=100 * (ww * x).sum() / ww.sum(), lo=np.nan, hi=np.nan))
for lab, d in (('text set, all', T), ('text set, PAF', T[T.af_type == 'PAF']), ('text set, PeAF', T[T.af_type == 'PeAF'])):
    dist(d, lab)
Tw = T.dropna(subset=['ipw'])
for lab, d in (('text set IPW, all', Tw), ('text set IPW, PAF', Tw[Tw.af_type == 'PAF']), ('text set IPW, PeAF', Tw[Tw.af_type == 'PeAF'])):
    dist(d, lab, 'ipw')
G = I[I.grade_source != '缺失']
dist(G, 'sensitivity: text + table-graded, all')
D = pd.DataFrame(rows); D.to_csv(TAB + 'R48_B_phenotype_distribution.csv', index=False)
# severity: analysis grade >=2 among present
sev = []
for v, gcol in (('MR', 'MRg'), ('TR', 'TRg')):
    for lab, d in (('all', T), ('PAF', T[T.af_type == 'PAF']), ('PeAF', T[T.af_type == 'PeAF'])):
        n = len(d); x = (d[gcol] >= 2).sum(); lo, hi = wilson(x, n)
        sev.append(dict(valve=v, group=lab, N=n, present_n=int((d[gcol] >= 1).sum()), grade_ge2_n=int(x), grade_ge2_pct=100 * x / n, lo=lo, hi=hi))
    for ph in PH:
        d = T[T.phenotype == ph]
        if len(d): sev.append(dict(valve=v, group='phenotype ' + PHE[ph], N=len(d), present_n=int((d[gcol] >= 1).sum()), grade_ge2_n=int((d[gcol] >= 2).sum()), grade_ge2_pct=100 * (d[gcol] >= 2).mean()))
pd.DataFrame(sev).to_csv(TAB + 'R48_C_severity.csv', index=False)

# ---------- 3. AF type and coexistence (prevalence ratios, modified Poisson)
A = T[T.af_type.isin(['PAF', 'PeAF'])].dropna(subset=['age']).copy()
pr = []
for out, lab in (('both', 'Both valves'), ('MRp', 'Any MR'), ('TRp', 'Any TR')):
    for adj, cols in (('crude', ['PeAF']), ('age, sex', ['PeAF', 'age', 'female'])):
        b, se = poisson(A[out].values.astype(float), design(A, cols))
        pr.append(dict(outcome=lab, adjustment=adj, PR_PeAF_vs_PAF=np.exp(b[1]), lo=np.exp(b[1] - 1.96 * se[1]), hi=np.exp(b[1] + 1.96 * se[1]), n=len(A),
                       PAF_pct=100 * A[A.PeAF == 0][out].mean(), PeAF_pct=100 * A[A.PeAF == 1][out].mean()))
# coexistence among those with any regurgitation
Aa = A[(A.MRp + A.TRp) > 0]
for adj, cols in (('crude', ['PeAF']), ('age, sex', ['PeAF', 'age', 'female'])):
    b, se = poisson(Aa.both.values.astype(float), design(Aa, cols))
    pr.append(dict(outcome='Both valves, among any regurgitation', adjustment=adj, PR_PeAF_vs_PAF=np.exp(b[1]), lo=np.exp(b[1] - 1.96 * se[1]), hi=np.exp(b[1] + 1.96 * se[1]), n=len(Aa),
                   PAF_pct=100 * Aa[Aa.PeAF == 0].both.mean(), PeAF_pct=100 * Aa[Aa.PeAF == 1].both.mean()))
pd.DataFrame(pr).to_csv(TAB + 'R48_D_AFtype_prevalence_ratios.csv', index=False)

# ---------- 4. MR-TR co-occurrence beyond chance at baseline (marginal and conditional O/E; adjusted OR)
def oe_marg(d):
    return d.both.sum() / (len(d) * d.MRp.mean() * d.TRp.mean())
covO = ['age', 'female', 'PeAF'] + (['AFunk'] if T.AFunk.sum() > 0 else [])   # E28：分析集已无未分型者
def oe_cond(d):
    M = design(d, covO); _, _, pm = logit(d.MRp.values, M); _, _, ptr = logit(d.TRp.values, M); return d.both.sum() / (pm * ptr).sum()
co = []
Tc = T.dropna(subset=['age']).copy()
for lab, d in (('all', Tc), ('PAF', Tc[Tc.af_type == 'PAF']), ('PeAF', Tc[Tc.af_type == 'PeAF'])):
    idx = np.arange(len(d)); bm = []; bc = []
    for _ in range(1000):
        s = d.iloc[rng.choice(idx, len(idx))]; bm.append(oe_marg(s))
        if lab == 'all' and len(bc) < 500: bc.append(oe_cond(s))
    r = dict(group=lab, n=len(d), observed_both=int(d.both.sum()), OE_marginal=oe_marg(d), OE_marg_lo=np.percentile(bm, 2.5), OE_marg_hi=np.percentile(bm, 97.5))
    if lab == 'all': r.update(OE_conditional_age_sex_AFtype=oe_cond(d), OE_cond_lo=np.percentile(bc, 2.5), OE_cond_hi=np.percentile(bc, 97.5))
    co.append(r)
for adj, cols in (('crude', ['MRp']), ('age, sex, AF type', ['MRp'] + covO), ('+ LA, RA, LVEDD, LVEF', ['MRp'] + covO + ['LA', 'RA', 'LVEDD', 'LVEF'])):
    d = Tc.dropna(subset=[c for c in cols]); b, se, _ = logit(d.TRp.values, design(d, cols))
    co.append(dict(group=f'OR TR~MR, {adj}', n=len(d), OR=np.exp(b[1]), OR_lo=np.exp(b[1] - 1.96 * se[1]), OR_hi=np.exp(b[1] + 1.96 * se[1])))
pd.DataFrame(co).to_csv(TAB + 'R48_E_baseline_cooccurrence.csv', index=False)

# ---------- 5. structural background: chamber sizes by phenotype (adjusted differences vs neither)
ch = []
for c in ('LA', 'RA', 'LVEDD', 'RV', 'LVEF'):
    for ph in PH:
        v = T[T.phenotype == ph][c].dropna(); ch.append(dict(measure=c, phenotype=PHE[ph], n=len(v), median=v.median(), q1=v.quantile(.25), q3=v.quantile(.75)))
    d = Tc.dropna(subset=[c]).copy()
    for ph in PH[1:]: d['ph_' + PHE[ph]] = (d.phenotype == ph).astype(float)
    cols = ['ph_MR only', 'ph_TR only', 'ph_Both', 'age', 'female', 'PeAF'] + (['AFunk'] if d.AFunk.sum() > 0 else [])
    b, se = ols(d[c].values, design(d, cols))
    for k, ph in enumerate(['MR only', 'TR only', 'Both'], start=1):
        ch.append(dict(measure=c, phenotype=f'adjusted difference vs Neither: {ph}', n=len(d), median=b[k], q1=b[k] - 1.96 * se[k], q3=b[k] + 1.96 * se[k]))
pd.DataFrame(ch).to_csv(TAB + 'R48_F_chambers_by_phenotype.csv', index=False)

# ---------- 6. representativeness of the paired cohort among baseline regurgitation (text set)
Rg = T[(T.MRp + T.TRp) > 0].copy()
rp = []
for lab, d in (('baseline regurgitation, in paired analysis cohort', Rg[Rg.in844]), ('baseline regurgitation, not in cohort', Rg[~Rg.in844])):
    rp.append(dict(group=lab, n=len(d), age_mean=d.age.mean(), female_pct=100 * d.female.mean(), PeAF_pct=100 * (d.af_type == 'PeAF').mean(),
                   both_pct=100 * (d.phenotype == '双瓣共存').mean(), MR_only_pct=100 * (d.phenotype == '仅MR').mean(), TR_only_pct=100 * (d.phenotype == '仅TR').mean(),
                   MR_ge2_pct=100 * (d.MRg >= 2).mean(), TR_ge2_pct=100 * (d.TRg >= 2).mean(), LA_median=d.LA.median(), RA_median=d.RA.median(), LAAO_pct=100 * d.laao.mean()))
pd.DataFrame(rp).to_csv(TAB + 'R48_G_representativeness.csv', index=False)

# ---------- 7. exclusions and strict-rule sensitivity
_C = I_ALL.confirmed
REV_CATS = REV[REV['团队判断(排除/保留)'] == '排除'].assign(u=lambda r: r['患者唯一号'].str.strip()).merge(I_ALL.loc[_C & I_ALL.excl_team_review, ['患者唯一号']].astype(str).apply(lambda c: c.str.strip()).rename(columns={'患者唯一号': 'u'}), on='u')['初筛理由'].value_counts()
ex = dict(index_admissions=len(I_ALL), report_same_admission=int((I_ALL.grade_source == 'TTE文本解析').sum()), report_without_procedure_date=int(((I_ALL.grade_source == 'TTE文本解析') & ~_C).sum()),
          text_set_before_exclusion=int(_C.sum()),
          excluded_valve_surgery_text_set=int((_C & I_ALL.excl_valve_surgery).sum()),
          excluded_team_review_text_set=int((_C & I_ALL.excl_team_review).sum()), analysis_set=int(I.text.sum()), overlap_with_paired=int(I[I.text].in844.sum()), grades_replaced_by_paired=int((I.text & I.in844 & I.tab_MR8.notna()).sum()), grades_changed_by_paired=N_OVERRIDE_CHANGED,
          excluded_strict_text_set=int((_C & I_ALL.excl_strict).sum()), **{'strict_' + k: int(v) for k, v in EXCL_CATS.items()}, **{'review_' + k: int(v) for k, v in REV_CATS.items()})
pd.DataFrame([ex]).to_csv(TAB + 'R48_H0_exclusions.csv', index=False)
Ts = T[~I.loc[T.index, 'excl_strict']].copy(); Tsc = Ts.dropna(subset=['age'])
sv = []
for lab, d in (('all', Ts), ('PAF', Ts[Ts.af_type == 'PAF']), ('PeAF', Ts[Ts.af_type == 'PeAF'])):
    r = dict(analysis='strict rule: phenotype', group=lab, N=len(d)); r.update({PHE[ph]: 100 * (d.phenotype == ph).mean() for ph in PH}); sv.append(r)
As = Tsc[Tsc.af_type.isin(['PAF', 'PeAF'])]
b, se = poisson(As.both.values.astype(float), design(As, ['PeAF', 'age', 'female']))
sv.append(dict(analysis='strict rule: PR both valves PeAF vs PAF (age, sex)', group='', N=len(As), est=np.exp(b[1]), lo=np.exp(b[1] - 1.96 * se[1]), hi=np.exp(b[1] + 1.96 * se[1])))
bc = [oe_cond(Tsc.iloc[rng.choice(len(Tsc), len(Tsc))]) for _ in range(300)]
sv.append(dict(analysis='strict rule: O/E conditional on age, sex, AF type', group='', N=len(Tsc), est=oe_cond(Tsc), lo=np.percentile(bc, 2.5), hi=np.percentile(bc, 97.5)))
sv.append(dict(analysis='strict rule: O/E marginal', group='', N=len(Tsc), est=oe_marg(Tsc)))
pd.DataFrame(sv).to_csv(TAB + 'R48_H_strict_rule_sensitivity.csv', index=False)

# ---------- figure (aggregate)
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 8, 'axes.spines.top': False, 'axes.spines.right': False})
COL = {'Neither': '#d9d9d9', 'MR only': '#b5533c', 'TR only': '#6fa3c7', 'Both': '#1f5a8a'}
fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.6), dpi=300, gridspec_kw={'width_ratios': [1.4, 1]})
grp = [('All', 'text set, all'), ('Paroxysmal AF', 'text set, PAF'), ('Persistent AF', 'text set, PeAF')]
for yi, (lab, key) in enumerate(grp[::-1]):
    left = 0; d = D[D.analysis == key].set_index('phenotype')
    for ph in ['Neither', 'MR only', 'TR only', 'Both']:
        w_ = d.loc[ph, 'pct']; ax[0].barh(yi, w_, left=left, color=COL[ph], edgecolor='white', height=0.6)
        ax[0].text(left + w_ / 2, yi, f'{w_:.0f}', ha='center', va='center', fontsize=7, color='white' if ph in ('Both', 'MR only') else '#222'); left += w_
    ax[0].text(101, yi, f'n={int(d.N.iloc[0])}', va='center', fontsize=7)
ax[0].set_yticks(range(3)); ax[0].set_yticklabels([g[0] for g in grp[::-1]]); ax[0].set_xlim(0, 100); ax[0].set_xlabel('Patients, %')
ax[0].legend([plt.Rectangle((0, 0), 1, 1, color=COL[k]) for k in COL], list(COL), ncol=4, frameon=False, fontsize=7, loc='lower left', bbox_to_anchor=(0, 1.0))
ax[0].text(-0.28, 1.12, 'A', transform=ax[0].transAxes, fontsize=10, fontweight='bold')
F = pd.read_csv(TAB + 'R48_F_chambers_by_phenotype.csv'); F = F[~F.phenotype.str.startswith('adjusted')]
x = np.arange(4); wdt = 0.38
for k, (m, colr) in enumerate((('LA', '#b5533c'), ('RA', '#1f5a8a'))):
    d = F[F.measure == m].set_index('phenotype').loc[['Neither', 'MR only', 'TR only', 'Both']]
    ax[1].bar(x + (k - 0.5) * wdt, d['median'], wdt, color=colr, label=m + ' diameter')
    ax[1].errorbar(x + (k - 0.5) * wdt, d['median'], yerr=[d['median'] - d.q1, d.q3 - d['median']], fmt='none', ecolor='#333', lw=0.7, capsize=1.5)
ax[1].set_xticks(x); ax[1].set_xticklabels(['Neither', 'MR only', 'TR only', 'Both'], fontsize=7); ax[1].set_ylabel('Median (IQR), mm'); ax[1].set_ylim(20, 60)
ax[1].legend(frameon=False, fontsize=7); ax[1].text(-0.25, 1.12, 'B', transform=ax[1].transAxes, fontsize=10, fontweight='bold')
fig.tight_layout(); fig.savefig(FIG + 'R48_baseline_phenotype.png', dpi=300)
print('done')
