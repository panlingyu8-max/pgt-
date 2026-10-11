import math, importlib.util, sys
spec=importlib.util.spec_from_file_location('p', sys.argv[1]); 
src=open(sys.argv[1]).read().split('# assumptions')[0]; ns={}; exec(src, ns)
power=ns['power']
women_def=0.74*0.85
factors=[('Age>25',1.0,0.45),('BMI>=24',1.0,0.25),('Premenstrual',women_def,0.45),('Foods',1.0,0.35)]
for py in (0.50,0.40,0.30,0.25):
    for n in (300,350,400,500):
        row=[f'{name} {power(n*f,px,py,2.0)[0]:.0%}' for name,f,px in factors]
        print(f'prev {py:.0%} N={n}: '+' | '.join(row))
    print()
