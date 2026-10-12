"""2023年 1次・2次合宿走行距離ランキングPDF(pdftotext -layout の出力)から data/camp_distance.csv に2023年分を追記
usage: pdftotext -layout <pdf> - | python3 scripts/extract_camp_pdf2023.py"""
import csv, re, sys, unicodedata

def norm(n):
    return unicodedata.normalize('NFKC', n).replace(' ', '').replace('熙', '煕').replace('昴', '昂').replace('黑', '黒').replace('⻄', '西')

text = unicodedata.normalize('NFKC', sys.stdin.read()).replace('黑', '黒').replace('⻄', '西')
part1, part2 = text.split('2次合宿走行距離ランキング')
camp1, camp2 = {}, {}
row = re.compile(r'(\d+|-)\s+(\S+)\s+([\d.]+)')
for line in part1.splitlines():
    m = row.findall(line)
    if len(m) == 4:                      # 4列目が 7/26-8/16 の合計
        camp1[norm(m[3][1])] = float(m[3][2])
for line in part2.splitlines():
    m = row.findall(line)
    if m and m[0][0] != '-':             # 1列目が全体ランキング
        camp2[norm(m[0][1])] = float(m[0][2])

rows = [r for r in csv.DictReader(open('data/camp_distance.csv')) if r['year'] != '2023']
fmt = lambda v: '' if v is None else (int(v) if v == int(v) else v)
with open('data/camp_distance.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['year', 'name', 'camp1_km', 'camp2_km'])
    for n in camp1:
        w.writerow(['2023', n, fmt(camp1[n]), fmt(camp2.get(n))])
    for r in rows:
        w.writerow([r['year'], r['name'], r['camp1_km'], r['camp2_km']])
print(len(camp1), len(camp2), file=sys.stderr)
