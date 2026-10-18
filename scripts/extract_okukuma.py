"""奥球磨駅伝の成績PDF（pdftotext -layout の出力）から高校の部の区間成績のCSVを作る
（列は extract_nihonkai.py と同じ。都道府県はPDFに無いので空欄）
usage: python3 scripts/extract_okukuma.py okkm_results.pdf data/okukuma_2026.csv"""
import csv, re, subprocess, sys

lines = subprocess.run(['pdftotext', '-layout', sys.argv[1], '-'], capture_output=True, text=True).stdout.splitlines()
starts = [i for i, l in enumerate(lines) if '奥球磨男子駅伝競走大会' in l]
lines = lines[:starts[1]]  # 2つ目の見出し以降は大学・一般の部（4区間）
head = next(l for l in lines if '第1区' in l)
DIST = [float(d) for d in re.findall(r'区\(([\d.]+)km\)', head)]
T = r'(?:\d+時間)?\d+分\d+秒'
PAIR = re.compile(rf'({T})\s+(\d+)')


def hms(s):
    m = re.fullmatch(r'(?:(\d+)時間)?(\d+)分(\d+)秒', s)
    return '{}:{:02d}:{:02d}'.format(int(m.group(1) or 0), int(m.group(2)), int(m.group(3)))


out = []
for i, l in enumerate(lines):
    m = re.match(rf'\s*(\d+)\s+(.*?)\s*({T})\s', l)
    if not m or len(PAIR.findall(l)) != 7:
        continue
    j = next(j for j in range(i + 1, i + 4) if len(PAIR.findall(lines[j])) == 7)
    school = m.group(2) + ''.join(x.strip() for x in lines[i + 1:j]) + lines[j].split()[0] * (not re.match(rf'\s*{T}', lines[j]))
    toks = [[t.group(), t.start(), t.end()] for t in re.finditer(r'\S+(?: \S+)*', lines[i - 1])][1:]  # 先頭はNo.
    while len(toks) > 7:  # 姓名の間が広い選手は、間隔が一番狭い2語をつなぐ
        k = min(range(len(toks) - 1), key=lambda k: toks[k + 1][1] - toks[k][2])
        toks[k:k + 2] = [[toks[k][0] + ' ' + toks[k + 1][0], toks[k][1], toks[k + 1][2]]]
    assert len(toks) == 7, (i, toks)
    sec = PAIR.findall(lines[j])
    for k, (name, (t, rk)) in enumerate(zip(toks, sec), 1):
        out.append({'team_rank': m.group(1), 'team': school, 'pref': '', 'team_time': hms(m.group(3)), 'section': k,
                    'distance_km': DIST[k - 1], 'name': name[0], 'time': hms(t), 'rank': rk})
w = csv.DictWriter(open(sys.argv[2], 'w', encoding='utf-8', newline=''), fieldnames=list(out[0]))
w.writeheader()
w.writerows(out)
print(len(out) // 7, 'teams', len(out), 'runners')
