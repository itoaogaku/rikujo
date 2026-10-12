"""2026年 合宿走行距離・消化率PDF(pdftotext -layout の出力)から2026年分を camp_distance.csv / camp_completion.csv に反映
usage: python3 scripts/extract_camp_pdf2026.py <走行距離1次.txt> <走行距離2次.txt> <消化率.txt>"""
import csv, re, sys, unicodedata

def clean(text):
    return (unicodedata.normalize('NFKC', text)
            .replace('黑', '黒').replace('⻄', '西').replace('⻯', '竜'))

def norm(n):
    return re.sub(r'\s', '', n).replace('熙', '煕').replace('昴', '昂')

NUM = re.compile(r'^(-|\d+(\.\d+)?%?)$')

def rows(lines):
    """名前 + 数値列 の行を (名前, [数値トークン]) で返す。平均行で止まる"""
    for line in lines:
        toks = line.split()
        if not toks:
            continue
        i = next((k for k, t in enumerate(toks) if NUM.match(t)), None)
        if not i:
            continue
        name = norm(''.join(toks[:i]))
        if '平均' in name:
            break
        yield name, toks[i:]

def section(text, start, end=None):
    s = text.index(start)
    e = text.index(end, s) if end else len(text)
    return text[s:e].splitlines()[2:]

d1 = clean(open(sys.argv[1]).read())
d2 = clean(open(sys.argv[2]).read())
cp = clean(open(sys.argv[3]).read())

camp1 = {n: float(v[-1]) for n, v in rows(section(d1, '2026年1次合宿走行距離'))}
camp2 = {}
for title in ['2026年2次合宿走行距離(選抜)', '2026年2次合宿走行距離(チャレンジ)']:
    for n, v in rows(section(d2, title)):
        camp2[n] = float(v[-1])
comp = {n: [float(x.rstrip('%')) / 100 for x in v[:4]]
        for n, v in rows(cp.splitlines()[5:]) if len(v) >= 4 and all(x.endswith('%') for x in v[:4])}

fmt = lambda v: '' if v is None else (int(v) if v == int(v) else v)
old = [r for r in csv.DictReader(open('data/camp_distance.csv')) if r['year'] != '2026']
with open('data/camp_distance.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['year', 'name', 'camp1_km', 'camp2_km'])
    for r in old:
        w.writerow([r['year'], r['name'], r['camp1_km'], r['camp2_km']])
    for n in camp1:
        w.writerow(['2026', n, fmt(camp1[n]), fmt(camp2.get(n))])

old = [r for r in csv.DictReader(open('data/camp_completion.csv')) if r['year'] != '2026']
with open('data/camp_completion.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['year', 'name', 'camp1', 'camp2', 'camp3', 'summer_avg'])
    for r in old:
        w.writerow([r['year'], r['name'], r['camp1'], r['camp2'], r['camp3'], r['summer_avg']])
    for n, v in comp.items():
        w.writerow(['2026', n] + [round(x, 3) for x in v])
print(len(camp1), len(camp2), len(comp), file=sys.stderr)
