import csv,sys
def s(t): m,x=t.split(':'); return int(m)*60+int(x)
def f(v): v=round(v); return f"{v//60}:{v%60:02d}"
rows=list(csv.DictReader(open(sys.argv[1])))
out=[]
for r in rows:
    d=float(r['distance_km']); base=10 if r['section']=='6' else 5
    w=s(r['winner_time'])*base/d; a=s(r['aogaku_time'])*base/d
    r.update(base_km=base,winner_conv=f(w),aogaku_conv=f(a),gap_conv_sec=f"{a-w:+.1f}")
    out.append(r)
w=csv.DictWriter(open(sys.argv[2],'w',newline=''),fieldnames=list(out[0].keys())); w.writeheader(); w.writerows(out)
for y in sorted({r['year'] for r in out}):
    print(f"\n### {y}\n|区間|距離|区間賞|換算|青学|順位|換算|差(秒)|\n|---|---|---|---|---|---|---|---|")
    for r in out:
        if r['year']==y: print(f"|{r['section']}区|{r['distance_km']}|{r['winner_name']}（{r['winner_team']}）{r['winner_time']}|{r['winner_conv']}|{r['aogaku_name']} {r['aogaku_time']}|{r['aogaku_rank']}|{r['aogaku_conv']}|{r['gap_conv_sec']}|")
