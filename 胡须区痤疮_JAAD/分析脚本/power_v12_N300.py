# v12 person-level logistic: power & CI width for a binary correlate vs binary outcome.
# Var(log OR) from expected 2x2 cells (Woolf), inflated 10% for covariate adjustment.
import math
from statistics import NormalDist
Z=NormalDist()
def cells(n, px, p0, OR):
    o1 = OR*p0/(1-p0); p1 = o1/(1+o1)
    n1, n0 = n*px, n*(1-px)
    return n1*p1, n1*(1-p1), n0*p0, n0*(1-p0)
def se_logor(n, px, py, OR, vif=1.10):
    # solve p0 so overall outcome prevalence = py
    lo, hi = 1e-6, 1-1e-6
    for _ in range(100):
        p0=(lo+hi)/2; a,b,c,d=cells(n,px,p0,OR)
        if (a+c)/n < py: lo=p0
        else: hi=p0
    a,b,c,d=cells(n,px,p0,OR)
    return math.sqrt(vif*(1/a+1/b+1/c+1/d))
def power(n,px,py,OR):
    se=se_logor(n,px,py,OR); z=math.log(OR)/se
    return Z.cdf(z-1.96)+Z.cdf(-z-1.96), se
# assumptions (design assumptions, not results)
N=300; women_def=0.74*0.85  # women with a definite premenstrual answer
factors=[('Age >25 y',1.0,0.45),('BMI >=24',1.0,0.25),('Premenstrual (women)',women_def,0.45),('Listed foods >=3/wk',1.0,0.35)]
for py,label in [(0.50,'any beard-associated site (~50%)'),(0.26,'upper lip (~26%)')]:
    print(f'\n== outcome: {label}')
    for n in (300,400,500):
        print(f' N={n}')
        for name,frac,px in factors:
            m=n*frac; row=[]
            for OR in (1.5,2.0,2.5):
                pw,se=power(m,px,py,OR); row.append(f'OR{OR}: power {pw:.0%}')
            _,se2=power(m,px,py,2.0)
            ci=(2*math.exp(-1.96*se2), 2*math.exp(1.96*se2))
            print(f'  {name:22s} n={m:4.0f} | '+' | '.join(row)+f' | if OR=2.0, CI ~{ci[0]:.2f}-{ci[1]:.2f}')
