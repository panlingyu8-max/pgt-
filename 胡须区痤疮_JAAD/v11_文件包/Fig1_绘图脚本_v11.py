import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':8,'axes.linewidth':0.6})
OR, BL, INK, MUTED = '#eb6834', '#2a78d6', '#1a1a1a', '#6b6a66'
rows = [('Age >25 y', 2.10, 0.98, '2.14 (1.51–3.05)'), ('Premenstrual\naggravation*', 2.375, 1.175, '2.02 (1.29–3.17)'),
        ('BMI ≥24 kg/m²', 1.85, 1.075, '1.72 (1.19–2.50)'), ('Listed high-GI\nfoods ≥3/wk', 1.72, 1.195, '1.44 (1.01–2.05)')]
fig = plt.figure(figsize=(5.2, 2.7), dpi=600)
ax = fig.add_axes([0.25, 0.20, 0.47, 0.66])
y = np.array([3.0, 2.0, 1.0, 0.0])
ax.axvline(1, color=MUTED, lw=0.6, ls=(0, (3, 2)), zorder=0)
for yi, (lab, b, o, ror) in zip(y, rows):
    ax.plot([o, b], [yi, yi], color='#c9c8c3', lw=2, zorder=1, solid_capstyle='round')
    ax.scatter([b], [yi], s=34, color=OR, zorder=3, edgecolor='white', linewidth=0.8)
    ax.scatter([o], [yi], s=30, marker='s', color=BL, zorder=3, edgecolor='white', linewidth=0.8)
    ax.text(3.45, yi, ror, va='center', ha='left', fontsize=7, color=INK, clip_on=False)
ax.set_xscale('log'); ax.set_xlim(0.7, 3.2); ax.set_xticks([0.75, 1, 1.5, 2, 3]); ax.set_xticklabels(['0.75', '1', '1.5', '2', '3'])
ax.minorticks_off(); ax.set_yticks(y); ax.set_yticklabels([r[0] for r in rows]); ax.set_ylim(-0.6, 3.6)
ax.tick_params(labelsize=7, width=0.5, length=2); ax.tick_params(axis='y', length=0)
for s_ in ('top', 'right', 'left'): ax.spines[s_].set_visible(False)
ax.set_xlabel('Adjusted OR by region (log scale)', fontsize=7)
ax.text(3.45, 3.75, 'ROR (95% CI)', fontsize=7, fontweight='bold', ha='left', clip_on=False)
ax.scatter([], [], s=34, color=OR, label='Beard-associated sites'); ax.scatter([], [], s=30, marker='s', color=BL, label='Other sites')
ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.07), ncol=2, frameon=False, fontsize=6.8, handletextpad=0.3, columnspacing=1.0)
fig.text(0.02, 0.03, '*Women only. Adjusted for age, sex, BMI, and current global grade. FICTIONAL DATA.', fontsize=5.8, color=MUTED)
fig.savefig('fig1_v11.png', dpi=600, facecolor='white'); fig.savefig('fig1_v11_preview.png', dpi=150, facecolor='white')
