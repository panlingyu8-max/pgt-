"""[@key] -> 顺序编号引用，并在“## 附表1”前插入参考文献表。
用法: python3 number_refs.py keyed.md refs.json out.md"""
import json, re, sys
src, lib_path, dst = sys.argv[1:4]
s = open(src).read(); lib = json.load(open(lib_path))
order = []
for m in re.finditer(r'\[@([^\]]+)\]', s):
    for k in m.group(1).split(';@'):
        if k not in order: order.append(k)
missing = [k for k in order if k not in lib]
assert not missing, missing
print('unused', [k for k in lib if k not in order])
num = {k: i + 1 for i, k in enumerate(order)}
def fmt(m):
    ns = sorted(num[k] for k in m.group(1).split(';@')); out = []; i = 0
    while i < len(ns):
        j = i
        while j + 1 < len(ns) and ns[j + 1] == ns[j] + 1: j += 1
        out.append(f"{ns[i]}–{ns[j]}" if j - i >= 2 else ",".join(map(str, ns[i:j + 1]))); i = j + 1
    return '[' + ','.join(out) + ']'
body = re.sub(r'\[@([^\]]+)\]', fmt, s)
refs = '## 参考文献\n\n' + '\n'.join(f"{i + 1}. {lib[k]}" for i, k in enumerate(order)) + '\n\n'
body = body.replace('## 附表1', refs + '## 附表1', 1) if '## 附表1' in body else body + '\n' + refs
open(dst, 'w').write(body); print(len(order), 'refs')
