"""消化率Excelの1〜3次合宿シートから、出雲メンバーの練習ごとの組・消化率を縦持ちCSVで出力
usage: python3 scripts/extract_workouts_xlsx.py <消化率.xlsx> > out.csv"""
import csv, datetime, re, sys
import openpyxl

MEMBERS = {
    '2023': ['野村昭夢', '黒田朝日', '佐藤一世', '山内健登', '鳥井健太', '鶴川正也'],
    '2025': ['小河原陽琉', '折田壮太', '飯田翔大', '神邑亮佑', '塩出翔太', '黒田朝日'],
}

def date(v):
    if isinstance(v, (int, float)) or (isinstance(v, str) and v.isdigit()):
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    return None

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
w = csv.writer(sys.stdout)
w.writerow(['year', 'camp', 'col', 'date', 'menu', 'name', 'group', 'completion'])
for year, members in MEMBERS.items():
    for camp in ['1次', '2次', '3次']:
        rows = list(ws_rows := wb[f'{year} {camp}合宿消化率'].iter_rows(values_only=True))
        cur = None
        for j in range(1, len(rows[2]), 2):
            if rows[1][j] is not None:
                cur = date(rows[1][j])
            menu = rows[2][j]
            if not menu or '消化率' in str(menu) or '名前' in str(menu):
                continue
            if 'ランキング' in str(menu):
                break
            menu = re.sub(r'\s+', ' ', str(menu)).strip()
            for r in rows[5:]:
                if r[0] and re.sub(r'\s', '', str(r[0])) in members:
                    g, c = r[j], r[j + 1] if j + 1 < len(r) else None
                    w.writerow([year, camp, j, cur.isoformat() if cur else '', menu,
                                re.sub(r'\s', '', str(r[0])), '' if g is None else str(g).replace('\n', ' '),
                                '' if c is None else c])
