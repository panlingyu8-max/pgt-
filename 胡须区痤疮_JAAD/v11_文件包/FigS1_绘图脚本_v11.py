import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Ellipse
from matplotlib.colors import LinearSegmentedColormap, Normalize
import numpy as np
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':8,'axes.linewidth':0.6})
OR, BL, INK, MUTED = '#eb6834', '#2a78d6', '#1a1a1a', '#6b6a66'
cmap = LinearSegmentedColormap.from_list('g', ['#f4f3f0', '#7d7b75'])
norm = Normalize(0, 50)

fig = plt.figure(figsize=(3.4, 3.35), dpi=600)
axA = fig.add_axes([0.0, 0.17, 0.84, 0.82]); 
axA.set_xlim(0, 10); axA.set_ylim(0, 10.6); axA.set_aspect('equal'); axA.axis('off')
head = Ellipse((5, 5.3), 6.6, 9.0, fill=False, lw=1.0, ec=INK, zorder=5); axA.add_patch(head)
clip = Ellipse((5, 5.3), 6.6, 9.0, transform=axA.transData)
zones = {
 'Forehead': ([(1, 7.1), (9, 7.1), (9, 10.5), (1, 10.5)], 37.9, 25.9, (5, 8.25), False),
 'Right cheek': ([(1, 7.0), (4.2, 7.0), (4.15, 4.6), (4.0, 3.5), (3.4, 3.6), (2.0, 3.9), (1, 4.6)], 39.7, 27.8, (2.75, 5.6), False),
 'Left cheek': ([(9, 7.0), (5.8, 7.0), (5.85, 4.6), (6.0, 3.5), (6.6, 3.6), (8.0, 3.9), (9, 4.6)], 41.4, 30.0, (7.25, 5.6), False),
 'Nose': ([(4.55, 7.0), (5.45, 7.0), (5.85, 4.65), (4.15, 4.65)], 22.4, 20.0, (5, 5.75), False),
 'Upper lip': ([(4.15, 4.62), (5.85, 4.62), (5.95, 3.95), (4.05, 3.95)], 25.9, 56.0, None, True),
 'Chin': ([(4.0, 3.45), (6.0, 3.45), (6.3, 1.9), (6.6, 0.4), (3.4, 0.4), (3.7, 1.9)], 39.7, 60.0, (5, 2.35), True),
 'JawR': ([(1, 4.6), (2.0, 3.9), (3.4, 3.6), (4.0, 3.45), (3.7, 1.9), (3.4, 0.4), (0.5, 2.5)], 32.8, 56.8, (2.55, 2.75), True),
 'JawL': ([(9, 4.6), (8.0, 3.9), (6.6, 3.6), (6.0, 3.45), (6.3, 1.9), (6.6, 0.4), (9.5, 2.5)], 32.8, 56.8, None, True),
}
for name, (pts, act, deep, lab, beard) in zones.items():
    p = Polygon(pts, closed=True, fc=cmap(norm(act)), ec='white', lw=0.8, zorder=1); axA.add_patch(p); p.set_clip_path(clip)
    if beard:
        q = Polygon(pts, closed=True, fill=False, ec=OR, lw=1.8, zorder=3); axA.add_patch(q); q.set_clip_path(clip)
    if lab:
        axA.text(lab[0], lab[1] + 0.22, f'{act:.0f}%', ha='center', va='center', fontsize=8.5, fontweight='bold', color=INK, zorder=6)
        axA.text(lab[0], lab[1] - 0.3, f'III–IV {deep:.0f}%', ha='center', va='center', fontsize=6, color=INK, zorder=6)
# mouth
axA.add_patch(Ellipse((5, 3.7), 1.5, 0.38, fc='white', ec=INK, lw=0.6, zorder=4))
# eyes
for x in (3.3, 6.7): axA.add_patch(Ellipse((x, 7.45), 1.0, 0.36, fc='white', ec=INK, lw=0.6, zorder=4))
# upper lip label outside with leader
axA.annotate('Upper lip  26%\nIII–IV 56%', xy=(5.6, 4.3), xytext=(9.0, 1.2), fontsize=6.3, color=INK, ha='left', va='center',
             arrowprops=dict(arrowstyle='-', lw=0.5, color=MUTED), zorder=7)
axA.text(2.55, 1.35, 'Jawline', ha='center', fontsize=6, color=MUTED, zorder=6)
axA.text(5, 1.3, 'Chin', ha='center', fontsize=6, color=MUTED, zorder=6)
axA.text(5, 9.35, 'Forehead', ha='center', fontsize=6, color=MUTED, zorder=6)

# legend A
fig.add_artist(plt.Line2D([0.52, 0.6], [0.126, 0.126], color=OR, lw=1.8)); fig.text(0.615, 0.126, 'Beard-associated site', fontsize=6.3, va='center')
cax = fig.add_axes([0.06, 0.115, 0.36, 0.022]); cb = fig.colorbar(plt.cm.ScalarMappable(norm=norm, cmap=cmap), cax=cax, orientation='horizontal')
cb.set_ticks([0, 25, 50]); cb.set_ticklabels(['0%', '25%', '50%']); cb.ax.tick_params(labelsize=6, width=0.5, length=2); cb.outline.set_linewidth(0.5)
cb.set_label('Participants with qualifying activity', fontsize=6.3, labelpad=1)

fig.savefig('figS1_v11.png', dpi=600, facecolor='white'); fig.savefig('figS1_preview.png', dpi=150, facecolor='white')
