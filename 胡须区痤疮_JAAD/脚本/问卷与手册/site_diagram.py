# 七部位边界图（患者版，中性配色，不突出任何区域）
import sys
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Ellipse
from matplotlib import font_manager
fp = [f for f in font_manager.findSystemFonts() if 'wqy-zenhei' in f.lower() or 'zenhei' in f.lower()][0]
font_manager.fontManager.addfont(fp)
plt.rcParams.update({'font.family': font_manager.FontProperties(fname=fp).get_name(), 'font.size': 9})
INK, FILL, LINE = '#1a1a1a', '#e9e7e2', '#ffffff'
fig = plt.figure(figsize=(5.2, 4.6), dpi=300)
ax = fig.add_axes([0.0, 0.08, 1.0, 0.92]); ax.set_xlim(-3.0, 13.0); ax.set_ylim(-0.6, 10.6); ax.set_aspect('equal'); ax.axis('off')
clip = Ellipse((5, 5.3), 6.6, 9.0, transform=ax.transData)
zones = {
 '前额': [(1, 7.1), (9, 7.1), (9, 10.5), (1, 10.5)],
 '右面颊': [(1, 7.0), (4.2, 7.0), (4.15, 4.6), (4.0, 3.5), (3.4, 3.6), (2.0, 3.9), (1, 4.6)],
 '左面颊': [(9, 7.0), (5.8, 7.0), (5.85, 4.6), (6.0, 3.5), (6.6, 3.6), (8.0, 3.9), (9, 4.6)],
 '鼻部': [(4.55, 7.0), (5.45, 7.0), (5.85, 4.65), (4.15, 4.65)],
 '人中': [(4.15, 4.62), (5.85, 4.62), (5.95, 3.95), (4.05, 3.95)],
 '下巴': [(4.0, 3.45), (6.0, 3.45), (6.3, 1.9), (6.6, 0.4), (3.4, 0.4), (3.7, 1.9)],
 'JawR': [(1, 4.6), (2.0, 3.9), (3.4, 3.6), (4.0, 3.45), (3.7, 1.9), (3.4, 0.4), (0.5, 2.5)],
 'JawL': [(9, 4.6), (8.0, 3.9), (6.6, 3.6), (6.0, 3.45), (6.3, 1.9), (6.6, 0.4), (9.5, 2.5)],
}
for pts in zones.values():
    p = Polygon(pts, closed=True, fc=FILL, ec=LINE, lw=2.2, zorder=1); ax.add_patch(p); p.set_clip_path(clip)
ax.add_patch(Ellipse((5, 5.3), 6.6, 9.0, fill=False, lw=1.1, ec=INK, zorder=5))
ax.add_patch(Ellipse((5, 3.7), 1.5, 0.38, fc='white', ec=INK, lw=0.7, zorder=4))
for x in (3.3, 6.7): ax.add_patch(Ellipse((x, 7.45), 1.0, 0.36, fc='white', ec=INK, lw=0.7, zorder=4))
lab = dict(ha='center', va='center', fontsize=14, color=INK, zorder=6)
ax.text(5, 8.6, '① 前额', **lab)
ax.text(2.95, 5.6, '③\n右面颊', **dict(lab, fontsize=13)); ax.text(7.05, 5.6, '②\n左面颊', **dict(lab, fontsize=13))
ax.text(5, 5.75, '④\n鼻部', **dict(lab, fontsize=12.5))
ax.text(5, 1.9, '⑥ 下巴', **lab)
for xy, xt, ha in (((2.75, 3.0), (-0.2, 1.6), 'right'), ((7.25, 3.0), (10.2, 1.6), 'left')):
    ax.annotate('⑦ 下颌缘', xy=xy, xytext=xt, fontsize=14, color=INK, ha=ha, va='center',
                arrowprops=dict(arrowstyle='-', lw=0.6, color=INK), zorder=7)
ax.annotate('⑤ 人中\n（鼻下至上唇）', xy=(5.75, 4.3), xytext=(10.2, 4.3), fontsize=13, color=INK, ha='left', va='center',
            arrowprops=dict(arrowstyle='-', lw=0.6, color=INK), zorder=7)
ax.text(5, -0.35, '颈部不计入', ha='center', va='center', fontsize=12, color='#5a5955')
fig.savefig(sys.argv[1], dpi=300, facecolor='white', bbox_inches='tight', pad_inches=0.05)
