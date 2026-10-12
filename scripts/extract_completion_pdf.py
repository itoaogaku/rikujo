"""合宿消化率PDF(pdftotext -layout の出力)の先頭まとめ表から、指定年の消化率を camp_completion.csv に反映
usage: python3 scripts/extract_completion_pdf.py <year> <消化率.txt>"""
import csv, re, sys, unicodedata

year, path = sys.argv[1], sys.argv[2]
text = (unicodedata.normalize('NFKC', open(path).read())
        .replace('黑', '黒').replace('⻄', '西').replace('⻯', '竜'))
row = re.compile(r'^\s*(\D+?)\s+([\d.]+)%\s+([\d.]+)%\s+([\d.]+)%\s+([\d.]+)%')
comp = {}
for line in text.splitlines()[5:]:
    m = row.match(line)
    if not m:
        continue
    name = re.sub(r'\s', '', m.group(1)).replace('熙', '煕').replace('昴', '昂')
    if '平均' in name:
        break
    comp[name] = [round(float(x) / 100, 3) for x in m.groups()[1:]]

old = [r for r in csv.DictReader(open('data/camp_completion.csv')) if r['year'] != year]
with open('data/camp_completion.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['year', 'name', 'camp1', 'camp2', 'camp3', 'summer_avg'])
    for r in old:
        w.writerow([r['year'], r['name'], r['camp1'], r['camp2'], r['camp3'], r['summer_avg']])
    for n, v in comp.items():
        w.writerow([year, n] + v)
print(len(comp), file=sys.stderr)
