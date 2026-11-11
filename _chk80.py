# -*- coding: utf-8 -*-
"""对比 8.3 ratio 图 2026 折线 vs 8.0 表格 rat 列。i4 用 off=4, f8 用 off=0。"""
import re, json, base64, struct
from json import JSONDecoder

HTML = r'd:\AAA东方嘉信\备份数据\qinrex\_analysis\seasonal\dashboard_total.html'
html = open(HTML, encoding='utf-8').read()

def decode_i4(obj):
    raw = base64.b64decode(obj['bdata'])
    n = (len(raw) - 4) // 4
    return list(struct.unpack('<%di' % n, raw[4:4+n*4]))

def decode_f8(obj):
    raw = base64.b64decode(obj['bdata'])
    n = len(raw) // 8
    return list(struct.unpack('<%dd' % n, raw[:n*8]))

# 8.0 表 rat
m = re.search(r'var T80_DATA = (\[.*?\]);', html, re.S)
t80 = json.loads(m.group(1))
print('8.0 表 rat（2026, 泰/印/越/马/科）:')
for r in t80:
    if r['y'] == 2026:
        print(f"  {r['d']}: {[None if v is None else round(v,1) for v in r['rat']]}")

for c in ['泰国', '印尼', '越南', '马来西亚', '科特迪瓦']:
    mm = re.search(r'id=["\']ratio_%s["\']' % c, html)
    if not mm:
        print(f'{c}: 未找到'); continue
    seg = html[mm.end(): mm.end()+30000]
    np = re.search(r"Plotly\.newPlot\(\s*['\"]ratio_%s['\"]\s*,\s*\[" % c, seg, re.S)
    pos = np.start() + np.end() - np.start() - 1
    obj, _ = JSONDecoder().raw_decode(seg[pos:])
    for tr in obj:
        if tr.get('name') == '2026' and tr.get('type') == 'scatter':
            xs = decode_i4(tr['x'])
            ys = decode_f8(tr['y'])
            pairs = [(a, round(b, 2)) for a, b in zip(xs, ys)]
            print(f'{c} 8.3 2026: {pairs}')
