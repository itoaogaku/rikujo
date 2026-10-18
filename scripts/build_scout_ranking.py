"""陸マガの高校ランキング（data/ranking_2028.xlsx：5000m・3000mSC・1500mが横並びの1シート）から
選手単位にまとめたスカウト用Excelを作る
usage: python3 scripts/build_scout_ranking.py data/ranking_2028.xlsx output/2028入学世代_スカウト一覧.xlsx"""
import re, sys, unicodedata
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = 'HGPｺﾞｼｯｸM'
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FONT = Font(name=FONT, bold=True, color='000000', size=10)
HEAD_FILL = PatternFill('solid', fgColor='A9D08E')
INPUT_FILL = PatternFill('solid', fgColor='FFFF99')
BASE = Font(name=FONT, size=10, color='000000')
BOLD = Font(name=FONT, size=10, color='000000', bold=True)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=False)
TFMT = '[m]:ss.00'
TOP10 = PatternFill('solid', fgColor='63BE7B')
TOP50 = PatternFill('solid', fgColor='A9D08E')
TOP100 = PatternFill('solid', fgColor='E2EFDA')

EVENTS = ['5000m', '3000mSC', '1500m']
# 同一人物で表記が違うもの（同じ高校・同じ種目に二重掲載されているもの含む）
NAME_ALIAS = {'ポールンジョロゲ': 'ポールジョロゲ', 'PAREIYOANTONYITA': 'パレイヨアントニーイタ'}
# 同名で高校が違っても同一人物とみなす下限（校名の文字の重なり率）
SAME_SCHOOL = 0.6


def key(name):
    k = re.sub(r'\s', '', unicodedata.normalize('NFKC', name))
    return NAME_ALIAS.get(k, k)


def school_core(s):
    s = unicodedata.normalize('NFKC', s)
    return re.sub(r'(高校|高|陸協)$', '', s.split('・')[0])


def similar(a, b):
    a, b = school_core(a), school_core(b)
    if not a or not b:
        return False
    return len(set(a) & set(b)) / min(len(set(a)), len(set(b))) >= SAME_SCHOOL


def tval(s):
    """'mm.ss.xx'（陸マガ表記）を日数（Excelの時刻値）に"""
    m, sec, frac = s.split('.')
    return (int(m) * 60 + int(sec) + int(frac) / 10 ** len(frac)) / 86400


def put(ws, r, c, v, fmt=None, font=BASE, align=CENTER):
    cell = ws.cell(r, c, v)
    cell.font, cell.border, cell.alignment = font, BORDER, align
    if fmt:
        cell.number_format = fmt
    return cell


def header(ws, r, labels, widths):
    for j, lab in enumerate(labels, 1):
        c = ws.cell(r, j, lab)
        c.font, c.fill, c.alignment, c.border = HEAD_FONT, HEAD_FILL, CENTER, BORDER
        ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    ws.row_dimensions[r].height = 14 * max(l.count('\n') + 1 for l in labels) + 6


def title(ws, text, sub):
    ws['A1'] = text
    ws['A1'].font = Font(name=FONT, bold=True, size=14, color='000000')
    ws['A2'] = sub
    ws['A2'].font = Font(name=FONT, size=9, color='000000')


def rank_colors(ws, ref):
    for op, val, fill in [('<=', 10, TOP10), ('<=', 50, TOP50), ('<=', 100, TOP100)]:
        ws.conditional_formatting.add(ref, CellIsRule(operator='lessThanOrEqual', formula=[str(val)], fill=fill,
                                                      stopIfTrue=True))


# ---------------- 読み込み ----------------
src = load_workbook(sys.argv[1]).active
head = [c.value for c in src[1]]
entries = []  # (ev, rank, name, school, time_str)
for row in src.iter_rows(min_row=2, values_only=True):
    for ev in EVENTS:
        i = head.index(f'{ev}_順位')
        rk, name, school, t = row[i:i + 4]
        if name:
            entries.append((ev, int(rk), str(name), str(school), str(t)))

# ---------------- 選手単位に統合 ----------------
athletes = []  # dict(name, schools, ev->(rank,time), notes)
by_key = defaultdict(list)
for ev, rk, name, school, t in entries:
    k = key(name)
    a = next((a for a in by_key[k] if any(similar(school, s) for s in a['schools'])), None)
    if a is None:
        a = {'name': name, 'schools': [], 'ev': {}, 'notes': []}
        by_key[k].append(a)
        athletes.append(a)
    if school not in a['schools']:
        a['schools'].append(school)
    if re.fullmatch(r'[゠-ヿ　\s]+', name) and not re.search(r'[゠-ヿ]', a['name']):
        a['name'] = name  # ローマ字表記よりカタカナ表記を優先
    if ev in a['ev']:
        a['notes'].append(f'{ev}は二重掲載（{a["ev"][ev][0]}位と{rk}位）→上位を採用')
        if rk > a['ev'][ev][0]:
            continue
    a['ev'][ev] = (rk, t)
for k, group in by_key.items():
    if len(group) > 1:
        for a in group:
            others = '・'.join(s for b in group if b is not a for s in b['schools'])
            a['notes'].append(f'同名の別選手あり（{others}）')
for a in athletes:
    # 表示用の校名：「高」で終わり、中黒を含まない最長の表記
    cands = sorted(a['schools'], key=lambda s: (s.endswith('高'), '・' not in s, len(s)), reverse=True)
    a['school'] = cands[0]
    if len(a['schools']) > 1:
        a['notes'].append('校名表記：' + '／'.join(a['schools']))
    hand = [ev for ev, (_, t) in a['ev'].items() if len(t.split('.')[-1]) == 1]
    if hand:
        a['notes'].append('手動計時：' + '・'.join(hand))
    a['best'] = min(r for r, _ in a['ev'].values())
athletes.sort(key=lambda a: (a['best'], -len(a['ev']), sum(r for r, _ in a['ev'].values())))

wb = Workbook()

# ---------------- 選手一覧 ----------------
ws = wb.active
ws.title = '選手一覧'
title(ws, '2028年入学世代（現高校2年）中長距離 選手一覧',
      '陸マガの5000m・3000mSC・1500mランキングを選手ごとに1行にまとめたもの。並び順＝3種目の中での最高順位。'
      '順位の色：濃い緑＝10位以内、緑＝50位以内、薄緑＝100位以内。見出しのフィルターで高校名・種目数などを絞り込めます。')
ws['A3'] = '基準記録（黄色セルを編集）→'
ws['A3'].font = BOLD
TH = {}
for j, (ev, mn, sc) in enumerate([('5000m', 14, 30), ('3000mSC', 9, 20), ('1500m', 3, 55)]):
    c0 = 4 + j * 4
    put(ws, 3, c0, ev, font=BOLD)
    put(ws, 3, c0 + 1, mn, '0"分"').fill = INPUT_FILL
    put(ws, 3, c0 + 2, sc, '0.00"秒"').fill = INPUT_FILL
    TH[ev] = f'(${get_column_letter(c0 + 1)}$3*60+${get_column_letter(c0 + 2)}$3)/86400'
ws['A4'] = '基準記録の初期値は仮の値です。「基準クリア」列はその記録以内の種目数を数えます。メモ列（黄色）は自由に記入してください。'
ws['A4'].font = Font(name=FONT, size=9, color='000000')

H = ['No', '氏名', '高校', '5000m\n記録', '5000m\n順位', '3000mSC\n記録', '3000mSC\n順位', '1500m\n記録', '1500m\n順位',
     'ランク入り\n種目数', '最高\n順位', '最高順位\nの種目', '基準\nクリア', '備考', 'メモ']
HR = 6
header(ws, HR, H, [6, 18, 20, 10, 7, 10, 7, 10, 7, 9, 7, 9, 7, 40, 30])
r = HR + 1
for no, a in enumerate(athletes, 1):
    put(ws, r, 1, no)
    put(ws, r, 2, a['name'], align=LEFT)
    put(ws, r, 3, a['school'], align=LEFT)
    for j, ev in enumerate(EVENTS):
        rk, t = a['ev'].get(ev, (None, None))
        put(ws, r, 4 + j * 2, tval(t) if t else None, TFMT)
        put(ws, r, 5 + j * 2, rk, '0')
    put(ws, r, 10, f'=COUNT(E{r},G{r},I{r})', '0')
    put(ws, r, 11, f'=MIN(E{r},G{r},I{r})', '0', BOLD)
    put(ws, r, 12, f'=IF(E{r}=K{r},"5000m",IF(G{r}=K{r},"3000mSC","1500m"))')
    put(ws, r, 13, f'=(D{r}<>"")*(D{r}<={TH["5000m"]})+(F{r}<>"")*(F{r}<={TH["3000mSC"]})'
                   f'+(H{r}<>"")*(H{r}<={TH["1500m"]})', '0;;""')
    put(ws, r, 14, '／'.join(a['notes']) or None, align=LEFT)
    put(ws, r, 15, None, align=LEFT).fill = INPUT_FILL
    r += 1
LAST = r - 1
rank_colors(ws, f'E{HR + 1}:E{LAST} G{HR + 1}:G{LAST} I{HR + 1}:I{LAST} K{HR + 1}:K{LAST}')
ws.conditional_formatting.add(f'M{HR + 1}:M{LAST}', CellIsRule(operator='greaterThanOrEqual', formula=['1'],
                                                                fill=PatternFill('solid', fgColor='FFEB9C')))
ws.freeze_panes = ws.cell(HR + 1, 4)
ws.auto_filter.ref = f'A{HR}:O{LAST}'
L = f"'選手一覧'!"

# ---------------- 高校別 ----------------
ws = wb.create_sheet('高校別')
title(ws, '高校別：ランク入り選手の人数（スカウト訪問先の検討用）',
      '人数は「選手一覧」から数式で集計。並び順＝100位以内の人数→ランク入り人数。選手名はいずれかの種目で100位以内の選手（最高順位）。')
by_school = defaultdict(list)
for a in athletes:
    by_school[a['school']].append(a)
schools = sorted(by_school, key=lambda s: (-sum(a['best'] <= 100 for a in by_school[s]), -len(by_school[s]),
                                           min(a['best'] for a in by_school[s])))
H = ['高校', 'ランク入り\n人数', 'いずれか\n100位以内', '5000m\n100位以内', '3000mSC\n100位以内', '1500m\n100位以内',
     '最上位\n順位', '100位以内の選手（最高順位）']
header(ws, 4, H, [22, 9, 9, 9, 9, 9, 7, 80])
SR = 5
rng = lambda col: f'{L}${col}${HR + 1}:${col}${LAST}'
for r, s in enumerate(schools, SR):
    put(ws, r, 1, s, align=LEFT)
    put(ws, r, 2, f'=COUNTIF({rng("C")},A{r})', '0')
    put(ws, r, 3, f'=COUNTIFS({rng("C")},A{r},{rng("K")},"<=100")', '0;;""')
    for j, col in enumerate('EGI'):
        put(ws, r, 4 + j, f'=COUNTIFS({rng("C")},A{r},{rng(col)},"<=100")', '0;;""')
    put(ws, r, 7, f'=_xlfn.MINIFS({rng("K")},{rng("C")},A{r})', '0', BOLD)
    top = [f'{a["name"].replace("　", "")}({a["best"]})' for a in by_school[s] if a['best'] <= 100]
    put(ws, r, 8, '、'.join(top) or None, align=LEFT)
SLAST = SR + len(schools) - 1
rank_colors(ws, f'G{SR}:G{SLAST}')
ws.freeze_panes = 'B5'
ws.auto_filter.ref = f'A4:H{SLAST}'

# ---------------- 種目別ランキング ----------------
for ev in EVENTS:
    ws = wb.create_sheet(ev)
    rows = sorted((e for e in entries if e[0] == ev), key=lambda e: e[1])
    others = [o for o in EVENTS if o != ev]
    title(ws, f'{ev} ランキング（陸マガ掲載順）',
          f'記録は元データのまま。トップ差・{others[0]}・{others[1]}は「選手一覧」から数式で参照（同一人物と判定できた場合のみ）。')
    H = ['順位', '氏名', '高校', '記録', 'トップ差\n(秒)', 'No', f'{others[0]}\n記録', f'{others[0]}\n順位',
         f'{others[1]}\n記録', f'{others[1]}\n順位']
    header(ws, 4, H, [7, 18, 20, 10, 9, 6, 10, 8, 10, 8])
    col = {'5000m': ('D', 'E'), '3000mSC': ('F', 'G'), '1500m': ('H', 'I')}
    r = 5
    for _, rk, name, school, t in rows:
        a = next(a for a in by_key[key(name)] if school in a['schools'])
        no = athletes.index(a) + 1
        put(ws, r, 1, rk, '0')
        put(ws, r, 2, name, align=LEFT)
        put(ws, r, 3, school, align=LEFT)
        put(ws, r, 4, tval(t), TFMT)
        put(ws, r, 5, f'=ROUND((D{r}-$D$5)*86400,2)', '+0.00;-0.00;0')
        put(ws, r, 6, no, '0')
        for j, o in enumerate(others):
            tc, rc = col[o]
            put(ws, r, 7 + j * 2, f'=IF(INDEX({rng(tc)},$F{r})="","",INDEX({rng(tc)},$F{r}))', TFMT)
            put(ws, r, 8 + j * 2, f'=IF(INDEX({rng(rc)},$F{r})="","",INDEX({rng(rc)},$F{r}))', '0')
        r += 1
    rank_colors(ws, f'A5:A{r - 1} H5:H{r - 1} J5:J{r - 1}')
    ws.freeze_panes = 'D5'
    ws.auto_filter.ref = f'A4:J{r - 1}'

for ws in wb:
    ws.sheet_view.zoomScale = 90
wb.save(sys.argv[2])
print(len(entries), 'entries ->', len(athletes), 'athletes,', len(schools), 'schools')
