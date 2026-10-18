"""陸マガの高校ランキング（data/ranking_2028.xlsx：5000m・3000mSC・1500mが横並びの1シート）から
選手単位にまとめたスカウト用Excelを作る。5000mを主に、並び順・集計は5000mの順位を基準にする。
駅伝の区間成績（RACESのCSV）も、ランキングに載っている選手（＝2028年入学世代）だけ照合して追記する
usage: python3 scripts/build_scout_ranking.py data/ranking_2028.xlsx output/2028入学世代_スカウト一覧.xlsx"""
import csv, re, sys, unicodedata
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string as col_idx, get_column_letter

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
# 駅伝：(シート名, 区間成績CSV, 大会名)。CSVは scripts/extract_nihonkai.py・extract_okukuma.py で作る
RACES = [('日本海駅伝', 'data/nihonkai_2026.csv', '第46回 日本海駅伝（2026/10/4）'),
         ('奥球磨駅伝', 'data/okukuma_2026.csv', '第5回 奥球磨駅伝 高校の部（2026/10/4）')]
# 旧字体など（陸マガのランキングは使っていない字）を揃える
KANJI = str.maketrans('﨑髙濵澤邊邉', '崎高浜沢辺辺')
MAIN_FILL = PatternFill('solid', fgColor='548235')
MAIN_HEAD = Font(name=FONT, bold=True, color='FFFFFF', size=10)


def key(name):
    k = re.sub(r'\s', '', unicodedata.normalize('NFKC', name)).translate(KANJI)
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


def team_schools(team):
    """駅伝のチーム名（例：佐久長聖A、鹿児島実業高校B、小林・延岡工業高校合同）を校名の候補に"""
    team = re.sub(r'(合同|[A-ZＡ-Ｚ])$', '', team)
    return [re.sub(r'(高校|高)$', '', t) + '高' for t in team.split('・')]


def time_val(s):
    """'h:mm:ss' を日数に"""
    h, m, sec = map(int, s.split(':'))
    return (h * 3600 + m * 60 + sec) / 86400


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
    a['r5000'] = a['ev'].get('5000m', (None,))[0]
# 並び順：5000mの順位。5000mの記録がない選手はその後ろに、他種目の最高順位順
athletes.sort(key=lambda a: (a['r5000'] is None, a['r5000'] or a['best'], -len(a['ev'])))
NO = {id(a): i for i, a in enumerate(athletes, 1)}

# ---------------- 駅伝：名前が同じで、チーム名と校名が似ている選手だけ照合 ----------------
races = []
for sheet, path, label in RACES:
    runs = []
    for x in csv.DictReader(open(path, encoding='utf-8')):
        a = next((a for a in by_key.get(key(x['name']), [])
                  if any(similar(t, s) for t in team_schools(x['team']) for s in a['schools'])), None)
        if a:
            x['no'] = NO[id(a)]
            a.setdefault('races', {})[sheet] = f"{x['section']}区 {x['rank']}位（{x['time'].lstrip('0:')}）"
            runs.append(x)
    runs.sort(key=lambda x: (int(x['section']), int(x['rank'])))
    races.append((sheet, label, runs))


wb = Workbook()

# ---------------- 選手一覧 ----------------
ws = wb.active
ws.title = '選手一覧'
title(ws, '2028年入学世代（現高校2年）中長距離 選手一覧（5000m順）',
      '陸マガの5000m・3000mSC・1500mランキングを選手ごとに1行にまとめたもの。並び順＝5000mの順位（5000mの記録がない選手は'
      'その後ろに他種目の最高順位順）。順位の色：濃い緑＝10位以内、緑＝50位以内、薄緑＝100位以内。'
      '見出しのフィルターで高校名・駅伝出場などを絞り込めます。')
ws['A3'] = '基準記録（黄色セルを編集）→'
ws['A3'].font = BOLD
TH = {}
for j, (ev, mn, sc) in enumerate([('5000m', 14, 30), ('3000mSC', 9, 20), ('1500m', 3, 55)]):
    c0 = 4 + j * 4
    put(ws, 3, c0, ev, font=BOLD)
    put(ws, 3, c0 + 1, mn, '0"分"').fill = INPUT_FILL
    put(ws, 3, c0 + 2, sc, '0.00"秒"').fill = INPUT_FILL
    TH[ev] = f'(${get_column_letter(c0 + 1)}$3*60+${get_column_letter(c0 + 2)}$3)/86400'
ws['A4'] = ('基準記録の初期値は仮の値です。「5000m基準」列は5000mが基準以内なら○、「基準クリア」列は3種目のうち基準以内の種目数。'
            'メモ列（黄色）は自由に記入してください。')
ws['A4'].font = Font(name=FONT, size=9, color='000000')

H = ['No', '氏名', '高校', '5000m\n記録', '5000m\n順位', '5000m\nトップ差(秒)', '5000m\n基準',
     '3000mSC\n記録', '3000mSC\n順位', '1500m\n記録', '1500m\n順位', 'ランク入り\n種目数', '基準\nクリア']
H += [f'{sheet}\n（区・区間順位）' for sheet, _, _ in races] + ['備考', 'メモ']
HR = 6
header(ws, HR, H, [6, 18, 20, 11, 8, 10, 7, 10, 8, 10, 8, 9, 7] + [18] * len(races) + [40, 30])
for c in range(4, 8):
    ws.cell(HR, c).fill, ws.cell(HR, c).font = MAIN_FILL, MAIN_HEAD
C = {'5000m': ('D', 'E'), '3000mSC': ('H', 'I'), '1500m': ('J', 'K')}
NOTE = 14 + len(races)
r = HR + 1
for a in athletes:
    put(ws, r, 1, NO[id(a)])
    put(ws, r, 2, a['name'], align=LEFT, font=BOLD if a['r5000'] else BASE)
    put(ws, r, 3, a['school'], align=LEFT)
    for ev, (tc, rc) in C.items():
        rk, t = a['ev'].get(ev, (None, None))
        f = BOLD if ev == '5000m' else BASE
        put(ws, r, col_idx(tc), tval(t) if t else None, TFMT, f)
        put(ws, r, col_idx(rc), rk, '0', f)
    put(ws, r, 6, f'=IF(D{r}="","",ROUND((D{r}-MIN($D${HR + 1}:$D$9999))*86400,2))', '+0.00;-0.00;0')
    put(ws, r, 7, f'=IF(D{r}="","",IF(D{r}<={TH["5000m"]},"○",""))')
    put(ws, r, 12, f'=COUNT(E{r},I{r},K{r})', '0')
    put(ws, r, 13, f'=(D{r}<>"")*(D{r}<={TH["5000m"]})+(H{r}<>"")*(H{r}<={TH["3000mSC"]})'
                   f'+(J{r}<>"")*(J{r}<={TH["1500m"]})', '0;;""')
    for j, (sheet, _, _) in enumerate(races):
        put(ws, r, 14 + j, a.get('races', {}).get(sheet), align=LEFT)
    put(ws, r, NOTE, '／'.join(a['notes']) or None, align=LEFT)
    put(ws, r, NOTE + 1, None, align=LEFT).fill = INPUT_FILL
    r += 1
LAST = r - 1
rank_colors(ws, f'E{HR + 1}:E{LAST} I{HR + 1}:I{LAST} K{HR + 1}:K{LAST}')
ws.conditional_formatting.add(f'G{HR + 1}:G{LAST}', CellIsRule(operator='equal', formula=['"○"'],
                                                                fill=PatternFill('solid', fgColor='FFEB9C')))
ws.freeze_panes = ws.cell(HR + 1, 4)
ws.auto_filter.ref = f'A{HR}:{get_column_letter(NOTE + 1)}{LAST}'
L = f"'選手一覧'!"
rng = lambda col: f'{L}${col}${HR + 1}:${col}${LAST}'


def other_events(ws, r, c, no_col):
    """選手一覧のNoから5000m・3000mSC・1500mの記録と順位を参照する列を c 列目から書く"""
    for j, ev in enumerate(EVENTS):
        tc, rc = C[ev]
        f = BOLD if ev == '5000m' else BASE
        put(ws, r, c + j * 2, f'=IF(INDEX({rng(tc)},${no_col}{r})="","",INDEX({rng(tc)},${no_col}{r}))', TFMT, f)
        put(ws, r, c + j * 2 + 1, f'=IF(INDEX({rng(rc)},${no_col}{r})="","",INDEX({rng(rc)},${no_col}{r}))', '0', f)


# ---------------- 種目別ランキング ----------------
for ev in EVENTS:
    ws = wb.create_sheet(ev)
    rows = sorted((e for e in entries if e[0] == ev), key=lambda e: e[1])
    others = [o for o in EVENTS if o != ev]
    title(ws, f'{ev} ランキング（陸マガ掲載順）',
          f'記録は元データのまま。{others[0]}・{others[1]}は「選手一覧」から数式で参照（同一人物と判定できた場合のみ）。')
    H = ['順位', '氏名', '高校', '記録', 'トップ差\n(秒)', 'No', f'{others[0]}\n記録', f'{others[0]}\n順位',
         f'{others[1]}\n記録', f'{others[1]}\n順位']
    header(ws, 4, H, [7, 18, 20, 10, 9, 6, 10, 8, 10, 8])
    r = 5
    for _, rk, name, school, t in rows:
        a = next(a for a in by_key[key(name)] if school in a['schools'])
        put(ws, r, 1, rk, '0')
        put(ws, r, 2, name, align=LEFT)
        put(ws, r, 3, school, align=LEFT)
        put(ws, r, 4, tval(t), TFMT)
        put(ws, r, 5, f'=ROUND((D{r}-$D$5)*86400,2)', '+0.00;-0.00;0')
        put(ws, r, 6, NO[id(a)], '0')
        for j, o in enumerate(others):
            tc, rc = C[o]
            put(ws, r, 7 + j * 2, f'=IF(INDEX({rng(tc)},$F{r})="","",INDEX({rng(tc)},$F{r}))', TFMT)
            put(ws, r, 8 + j * 2, f'=IF(INDEX({rng(rc)},$F{r})="","",INDEX({rng(rc)},$F{r}))', '0')
        r += 1
    rank_colors(ws, f'A5:A{r - 1} H5:H{r - 1} J5:J{r - 1}')
    ws.freeze_panes = 'D5'
    ws.auto_filter.ref = f'A4:J{r - 1}'

# ---------------- 高校別 ----------------
ws = wb.create_sheet('高校別', 2)
title(ws, '高校別：5000mランク入り選手の人数（スカウト訪問先の検討用）',
      '人数は「選手一覧」から数式で集計。並び順＝5000m100位以内の人数→5000mランク入り人数→5000m最上位順位。'
      '選手名は5000m100位以内の選手（5000m順位）。')
by_school = defaultdict(list)
for a in athletes:
    by_school[a['school']].append(a)
r5 = lambda s: [a['r5000'] for a in by_school[s] if a['r5000']]
schools = sorted(by_school, key=lambda s: (-sum(x <= 100 for x in r5(s)), -len(r5(s)), min(r5(s) or [9999]),
                                           -len(by_school[s])))
H = ['高校', '5000m\n100位以内', '5000m\nランク入り', '5000m\n最上位', '5000m100位以内の選手（5000m順位）',
     '3000mSC\n100位以内', '1500m\n100位以内', '全種目\nランク入り']
header(ws, 4, H, [22, 9, 9, 8, 70, 9, 9, 9])
for c in range(2, 6):
    ws.cell(4, c).fill, ws.cell(4, c).font = MAIN_FILL, MAIN_HEAD
SR = 5
for r, s in enumerate(schools, SR):
    put(ws, r, 1, s, align=LEFT)
    put(ws, r, 2, f'=COUNTIFS({rng("C")},A{r},{rng("E")},"<=100")', '0;;""', BOLD)
    put(ws, r, 3, f'=COUNTIFS({rng("C")},A{r},{rng("E")},">0")', '0;;""')
    put(ws, r, 4, f'=IF(C{r}=0,"",_xlfn.MINIFS({rng("E")},{rng("C")},A{r}))', '0', BOLD)
    top = [f'{a["name"].replace("　", "")}({a["r5000"]})' for a in by_school[s] if a['r5000'] and a['r5000'] <= 100]
    put(ws, r, 5, '、'.join(top) or None, align=LEFT)
    put(ws, r, 6, f'=COUNTIFS({rng("C")},A{r},{rng("I")},"<=100")', '0;;""')
    put(ws, r, 7, f'=COUNTIFS({rng("C")},A{r},{rng("K")},"<=100")', '0;;""')
    put(ws, r, 8, f'=COUNTIF({rng("C")},A{r})', '0')
SLAST = SR + len(schools) - 1
rank_colors(ws, f'D{SR}:D{SLAST}')
ws.freeze_panes = 'B5'
ws.auto_filter.ref = f'A4:H{SLAST}'

# ---------------- 駅伝 ----------------
for pos, (sheet, label, runs) in enumerate(races, 3):
    ws = wb.create_sheet(sheet, pos)
    title(ws, f'{label}：2028年入学世代（現高2）の区間成績',
          '全出走者のうち、陸マガランキングに名前と高校が一致する選手のみ掲載（成績表に学年の記載がないため）。'
          'ランキング圏外の2年生は含まれません。区間順位は全学年での順位、2年生内順位はこの表の中での順位。'
          '区間順位の色：濃い緑＝3位以内、緑＝10位以内、薄緑＝30位以内。')
    H = ['区', '距離\n(km)', '区間\n順位', '2年生内\n順位', '区間記録', '1km\nペース', '氏名', 'チーム', '都道府県',
         'チーム\n総合順位', 'No', '5000m\n記録', '5000m\n順位', '3000mSC\n記録', '3000mSC\n順位', '1500m\n記録', '1500m\n順位']
    header(ws, 4, H, [5, 8, 7, 8, 10, 8, 18, 20, 9, 8, 6, 11, 8, 10, 8, 10, 8])
    for c in (12, 13):
        ws.cell(4, c).fill, ws.cell(4, c).font = MAIN_FILL, MAIN_HEAD
    E = 4 + len(runs)
    r = 5
    for x in runs:
        put(ws, r, 1, int(x['section']), '0"区"')
        put(ws, r, 2, float(x['distance_km']), '0.0###')
        put(ws, r, 3, int(x['rank']), '0', BOLD)
        put(ws, r, 4, f'=COUNTIFS($A$5:$A${E},A{r},$E$5:$E${E},"<"&E{r})+1', '0')
        put(ws, r, 5, time_val(x['time']), '[m]:ss')
        put(ws, r, 6, f'=E{r}/B{r}', '[m]:ss')
        put(ws, r, 7, x['name'], align=LEFT)
        put(ws, r, 8, x['team'], align=LEFT)
        put(ws, r, 9, x['pref'] or None)
        put(ws, r, 10, int(x['team_rank']), '0')
        put(ws, r, 11, x['no'], '0')
        other_events(ws, r, 12, 'K')
        r += 1
    for val, fill in [(3, TOP10), (10, TOP50), (30, TOP100)]:
        ws.conditional_formatting.add(f'C5:C{E}', CellIsRule(operator='lessThanOrEqual', formula=[str(val)],
                                                             fill=fill, stopIfTrue=True))
    rank_colors(ws, f'M5:M{E} O5:O{E} Q5:Q{E}')
    prev, on = None, False
    for rr in range(5, E + 1):  # 区が変わるごとに薄い帯色
        if ws.cell(rr, 1).value != prev:
            on, prev = not on, ws.cell(rr, 1).value
        if on:
            for c in (1, 2, 4, 5, 6, 7, 8, 9, 10, 11):
                ws.cell(rr, c).fill = PatternFill('solid', fgColor='F2F2F2')
    if not any(x['pref'] for x in runs):
        ws.column_dimensions['I'].hidden = True  # 都道府県の記載がない大会
    ws.freeze_panes = 'H5'
    ws.auto_filter.ref = f'A4:Q{E}'

for ws in wb:
    ws.sheet_view.zoomScale = 90
wb.save(sys.argv[2])
print(len(entries), 'entries ->', len(athletes), 'athletes,', len(schools), 'schools,',
      ', '.join(f'{s} {len(x)}' for s, _, x in races))
