"""日本海駅伝の総合成績PDF（pdftotext -layout の出力）から区間成績のCSVを作る
usage: python3 scripts/extract_nihonkai.py nihonkai.pdf data/nihonkai_2026.csv"""
import csv, re, subprocess, sys

lines = subprocess.run(['pdftotext', '-layout', sys.argv[1], '-'], capture_output=True, text=True).stdout.splitlines()
head = next(l for l in lines if '１区' in l)
DIST = [float(d) for d in re.findall(r'\(([\d.]+)㎞\)', head)]
PAIR = re.compile(r'(\d:\d\d:\d\d)\s+(\d+)')
out = []
for i, l in enumerate(lines):
    m = re.fullmatch(r'\s*(\d+)\s+(\d+)\s*', l)
    if not m:
        continue
    toks = [[t.group(), t.start(), t.end()] for t in re.finditer(r'\S+(?: \S+)*', lines[i - 1])]
    while len(toks) > 8:  # 「山副   昇太郎」のように姓名の間が広い選手は、間隔が一番狭い2語をつなぐ
        j = min(range(len(toks) - 1), key=lambda j: toks[j + 1][1] - toks[j][2])
        toks[j:j + 2] = [[toks[j][0] + ' ' + toks[j + 1][0], toks[j][1], toks[j + 1][2]]]
    names = [t[0] for t in toks]
    assert len(names) == 8, (i, names)
    team, runners = names[0], names[1:]
    pref = re.fullmatch(r'\s*\((.+)\)\s*', lines[i + 1]).group(1)
    total = lines[i + 2].split()[0]
    sec = PAIR.findall(lines[i + 3])
    assert len(sec) == 7, (i, lines[i + 3])
    for k, (name, (t, rk)) in enumerate(zip(runners, sec), 1):
        out.append({'team_rank': m.group(1), 'team': team, 'pref': pref, 'team_time': total, 'section': k,
                    'distance_km': DIST[k - 1], 'name': name, 'time': t, 'rank': rk})
w = csv.DictWriter(open(sys.argv[2], 'w', encoding='utf-8', newline=''), fieldnames=list(out[0]))
w.writeheader()
w.writerows(out)
print(len(out) // 7, 'teams', len(out), 'runners')
