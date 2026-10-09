# -*- coding: utf-8 -*-
"""76. (E30) 配对反应分析集（分组 1/2/3，844 次手术）的患者计数键核查。在项目根目录运行；只输出汇总与 Y/N 比对，不输出任何标识符。
背景：
  - 23/27a 号脚本：患者键 = 患者唯一号（第 7 列），缺失时用登记号（第 4 列）→ 837 例，“7 例做了 2 次消融”。
  - 74 号脚本：只用患者唯一号，缺失值被转成字符串 'nan' 后合并 → 836 例（R56_A），这是一个程序错误。
  - 17 号 ANCOVA 质控：病案号（第 3 列）非缺失 843、互不重复；患者唯一号非缺失 842、去重后 835。
    若病案号为患者级编号，则 844 次手术对应 844 例患者，“7 例两次消融”不成立。
本脚本逐组比对患者唯一号重复的 7 组，判断它们是否为同一人。
输出：屏幕 + 09_原始数据_含直接标识符/R58_paired_patient_key_audit.csv（只含 Y/N 与差值，仍按内部文件管理）"""
import numpy as np, pandas as pd
raw = '09_原始数据_含直接标识符/出院时间2020-1-2024-9总表改2.xlsx'
R = pd.read_excel(raw, header=0, dtype=str)
c = lambda p: R.iloc[:, p - 1]
assert '病案号' in str(R.columns[2]) and '登记号' in str(R.columns[3]) and '患者唯一号' in str(R.columns[6]), list(R.columns[:8])
find = lambda key: next((j for j, n in enumerate(R.columns) if key in str(n)), None)
J = {k: find(k) for k in ('姓名', '出生', '出院时间', '入院时间', '手术日期')}
print('辅助列位置（0 起）:', J)
clean = lambda s: s.astype(str).str.strip().str.replace(r'\.0$', '', regex=True).replace({'nan': np.nan, 'None': np.nan, '': np.nan})
grp = pd.to_numeric(c(6), errors='coerce'); F = grp.isin([1, 2, 3])
D = pd.DataFrame({'mrn': clean(c(3)), 'reg': clean(c(4)), 'uid': clean(c(7)), 'sex': clean(c(12)), 'age': pd.to_numeric(c(13), errors='coerce')})
for k in ('姓名', '出生'):
    D[k] = clean(R.iloc[:, J[k]]) if J[k] is not None else np.nan
for k in ('出院时间', '入院时间', '手术日期'):
    D[k] = pd.to_datetime(R.iloc[:, J[k]], errors='coerce') if J[k] is not None else pd.NaT
P = D[F].copy(); assert len(P) == 844

# 1. 各种计数键下的患者数
uid_fb = P.uid.fillna('R' + P.reg.astype(str))
print('\n1. 844 次手术在不同计数键下的患者数')
print('  患者唯一号，缺失值误合并为一例（74 号脚本）:', P.uid.astype(str).nunique())
print('  患者唯一号，缺失时用登记号（23/27a 号脚本）:', uid_fb.nunique(), '| 患者唯一号缺失', int(P.uid.isna().sum()))
print('  登记号:', P.reg.nunique(), '| 缺失', int(P.reg.isna().sum()))
print('  病案号（缺失者各算一例）:', P.mrn.nunique() + int(P.mrn.isna().sum()), '| 缺失', int(P.mrn.isna().sum()))
if J['姓名'] is not None:
    print('  姓名+性别（仅参考，重名会低估）:', (P['姓名'].astype(str) + '|' + P.sex.astype(str)).nunique())

# 2. 患者唯一号重复组逐组比对（Y/N）
rows = []
for k, (u, g) in enumerate(P[P.uid.notna()].groupby('uid'), 1):
    if len(g) < 2: continue
    a, b = g.iloc[0], g.iloc[1]
    dt = (b['出院时间'] - a['出院时间']).days / 365.25 if pd.notna(a['出院时间']) and pd.notna(b['出院时间']) else np.nan
    yn = lambda x, y: ('—' if pd.isna(x) or pd.isna(y) else ('Y' if x == y else 'N'))
    rows.append(dict(group=len(rows) + 1, n_procedures=len(g), same_reg=yn(a.reg, b.reg), same_mrn=yn(a.mrn, b.mrn), same_sex=yn(a.sex, b.sex),
                     same_name=yn(a['姓名'], b['姓名']), same_birth=yn(a['出生'], b['出生']),
                     years_between_discharges=round(dt, 2) if pd.notna(dt) else np.nan, age_difference=b.age - a.age if pd.notna(a.age) and pd.notna(b.age) else np.nan,
                     same_admission_date=yn(a['入院时间'], b['入院时间'])))
A = pd.DataFrame(rows)
print('\n2. 患者唯一号重复的组（同一人应为：登记号、性别、姓名、出生日期相同；年龄差约等于两次出院相隔年数）')
print(A.to_string(index=False) if len(A) else '  无重复')
A.to_csv('09_原始数据_含直接标识符/R58_paired_patient_key_audit.csv', index=False)

# 3. 病案号是否为患者级编号：全表中同一登记号是否对应多个病案号
M = D.dropna(subset=['reg', 'mrn'])
k = M.groupby('reg').mrn.nunique()
print('\n3. 全表：有多次住院的登记号', int((M.groupby('reg').size() > 1).sum()), '个，其中对应多个病案号的', int((k > 1).sum()), '个')
print('   若后者接近前者，病案号是按住院分配的，不能用来数患者；若为 0，病案号是患者级编号。')
