"""月間走行距離Excelから year,name,month,km の縦持ちCSVを作る
usage: python3 scripts/extract_monthly.py <月間走行距離.xlsx>"""
import csv, re, sys
import openpyxl

def norm(n):
    return re.sub(r'\s', '', str(n)).replace('熙', '煕').replace('昴', '昂')

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
with open('data/monthly_distance.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['year', 'name', 'month', 'km'])
    for year in ['2022', '2023', '2024']:
        ws = next(s for s in wb if s.title.strip() == f'{year}月間走行距離')
        rows = list(ws.iter_rows(values_only=True))
        header = rows[1]
        cols = {int(h[:-1]): j for j, h in enumerate(header) if isinstance(h, str) and re.fullmatch(r'\d+月', h)}
        for r in rows[2:]:
            if not r[0] or '平均' in str(r[0]):
                continue
            for m, j in sorted(cols.items()):
                if isinstance(r[j], (int, float)):
                    w.writerow([year, norm(r[0]), m, round(r[j], 1)])
