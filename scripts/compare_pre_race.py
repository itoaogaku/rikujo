"""出雲の5km換算ペースと2週間前5000mを比較し data/izumo_vs_pre_race.csv を出力"""
import csv
def s(t):
    m, x = t.split(':'); return int(m) * 60 + float(x)
iz = list(csv.DictReader(open('data/izumo_sections.csv')))
pb = {(r['year'], r['name']): r for r in csv.DictReader(open('data/pre_race_5000m.csv'))}
out = []
for r in iz:
    p = pb[(r['year'], r['aogaku_name'])]
    pace5 = s(r['aogaku_time']) * 5 / float(r['distance_km'])
    t = s(p['time_5000m'])
    out.append(dict(year=r['year'], section=r['section'], name=r['aogaku_name'],
                    pre_5000m=p['time_5000m'], pre_race=p['race'], izumo_time=r['aogaku_time'],
                    izumo_rank=r['aogaku_rank'], izumo_pace5k_sec=round(pace5, 1),
                    diff_sec=round(pace5 - t, 1), drop_pct=round((pace5 / t - 1) * 100, 2)))
w = csv.DictWriter(open('data/izumo_vs_pre_race.csv', 'w', newline=''), fieldnames=list(out[0]))
w.writeheader(); w.writerows(out)
