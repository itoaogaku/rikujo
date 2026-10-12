"""MARCH対抗戦（10000m）リザルトPDFのテキスト(pdftotext -layout)から青学選手の結果を data/march_10000m.csv に反映
usage: pdftotext -layout <pdf> - | python3 scripts/extract_march.py <year>"""
import csv, os, re, sys, unicodedata

year = sys.argv[1]
text = unicodedata.normalize('NFKC', sys.stdin.read()).replace('黑', '黒')
KANA = r'[゠-ヿ]'
out = {}
total = text[min(text.index(k) for k in ('総合成績', '全体順位') if k in text):]
for line in total.splitlines():
    m = re.match(rf'\s*(\d+)\s+(\d+)\s+(A\d+)\s+(.+?)\s+(?:{KANA}[^\d]*?\s)?(\d)\s+青山学院大学\s+(?:[^\d\s]+\s+)?(\d+:\d+\.\d+)\s+(\d)組\s+(\d+)着(.*)', line)
    if m and re.sub(r'\s', '', m.group(4)) not in out:
        name = re.sub(r'\s', '', m.group(4)).replace('熙', '煕')
        name = {'西澤': '西澤マハロ'}.get(name, name)
        out[name] = dict(year=year, name=name, grade=m.group(5), time=m.group(6), heat=m.group(7),
                         heat_place=m.group(8), overall=m.group(1), team_place=m.group(2),
                         pb='自己新' if '自己新' in m.group(9) else '', status='完走')
    else:
        # 2022年形式：ゼッケンなし（全体順位 チーム順位 氏名 学年 大学名 陸協 記録 組 組順位）
        m2 = re.match(r'\s*(\d+)\s+(\d+)\s+(\D+?)\s+(\d)\s+青山学院大学\s+\S+\s+(\d+:\d+\.\d+)\s+(\d)組\s+(\d+)着(.*)', line)
        if m2 and re.sub(r'\s', '', m2.group(3)) not in out:
            name = re.sub(r'\s', '', m2.group(3)).replace('熙', '煕')
            out[name] = dict(year=year, name=name, grade=m2.group(4), time=m2.group(5), heat=m2.group(6),
                             heat_place=m2.group(7), overall=m2.group(1), team_place=m2.group(2),
                             pb='自己新' if '自己新' in m2.group(8) else '', status='完走')
heads = [text.index(k) for k in ('総合成績', 'チーム賞', '全体順位') if k in text]
heats = text[:min(heads)]
heat = None
for line in heats.splitlines():
    h = re.match(r'\s*(\d)組競技結果', line)
    if h:
        heat = h.group(1)
    m = re.match(rf'\s*(A\d+)\s+(.+?)\s+(?:{KANA}[^\d]*?\s)?(\d)\s+青山学院大学\s+(?:[^\d\s]+\s+)?(欠場|途中棄権|途棄)', line)
    if m and re.sub(r'\s', '', m.group(2)) not in out:
        name = re.sub(r'\s', '', m.group(2)).replace('熙', '煕')
        out[name] = dict(year=year, name=name, grade=m.group(3), time='', heat=heat, heat_place='',
                         overall='', team_place='', pb='', status={'途棄': '途中棄権'}.get(m.group(4), m.group(4)))
path = 'data/march_10000m.csv'
old = [r for r in csv.DictReader(open(path))] if os.path.exists(path) else []
cols = ['year', 'name', 'grade', 'time', 'heat', 'heat_place', 'overall', 'team_place', 'pb', 'status']
with open(path, 'w', newline='') as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    for r in old:
        if r['year'] != year:
            w.writerow(r)
    for r in sorted(out.values(), key=lambda r: (r['status'] != '完走', int(r['overall'] or 999))):
        w.writerow(r)
print(len(out), sum(r['status'] == '完走' for r in out.values()), file=sys.stderr)
