"""data/ 以下のCSVと docs/camp_workouts_izumo_members.md から、テーマ別シートの分析用Excelを作る
usage: python3 scripts/build_workbook.py output/出雲駅伝_振り返り分析.xlsx"""
import csv, re, sys
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = 'HGPｺﾞｼｯｸM'
GREEN = '1E5945'
LIGHT = 'E8F1EC'
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FONT = Font(name=FONT, bold=True, color='000000', size=10)
HEAD_FILL = PatternFill('solid', fgColor='A9D08E')
SUB_FILL = PatternFill('solid', fgColor=LIGHT)
BASE = Font(name=FONT, size=10, color='000000')
BLUE = Font(name=FONT, size=10, color='000000')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

YEARS = ['2022', '2023', '2024', '2025', '2026']
SEC_KM = {'1': 8.0, '2': 5.8, '3': 8.5, '4': 6.2, '5': 6.4, '6': 10.2}
RACE = {'adizero5K': 'ロード5km（adizero 5K、9/27）'}


def nospace(s):
    return re.sub(r'\s', '', s)


def tval(s):
    """'mm:ss' / 'mm:ss.xx' / 'h:mm:ss' を日数（Excelの時刻値）に"""
    parts = [float(p) for p in s.split(':')]
    sec = parts[-1] + parts[-2] * 60 + (parts[-3] * 3600 if len(parts) == 3 else 0)
    return sec / 86400


def race_label(r):
    if r in RACE:
        return RACE[r]
    m = re.match(r'track5000m_heat(\d+)', r)
    return f'トラック5000m（{m.group(1)}組）' if m else r


def rows(path):
    return list(csv.DictReader(open(path, encoding='utf-8')))


def title(ws, text, sub=None, width=10):
    ws['A1'] = text
    ws['A1'].font = Font(name=FONT, bold=True, size=14, color='000000')
    if sub:
        ws['A2'] = sub
        ws['A2'].font = Font(name=FONT, size=9, color='000000')
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
        ws['A2'].alignment = LEFT
        ws.row_dimensions[2].height = 30


def header(ws, r, labels, widths=None):
    for j, lab in enumerate(labels, 1):
        c = ws.cell(r, j, lab)
        c.font, c.fill, c.alignment, c.border = HEAD_FONT, HEAD_FILL, CENTER, BORDER
    ws.row_dimensions[r].height = 14 * max(str(l).count('\n') + 1 for l in labels) + 6
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w


def put(ws, r, c, v, fmt=None, font=BASE, align=CENTER):
    cell = ws.cell(r, c, v)
    cell.font, cell.border, cell.alignment = font, BORDER, align
    if fmt:
        cell.number_format = fmt
    return cell


def band(ws, r1, r2, ncol, key_col=1):
    """年が変わるごとに薄い帯色を付けて読みやすくする"""
    prev, on = None, False
    for r in range(r1, r2 + 1):
        v = ws.cell(r, key_col).value
        if v != prev:
            on, prev = not on, v
        if on:
            for c in range(1, ncol + 1):
                ws.cell(r, c).fill = SUB_FILL


SECTION = {(x['year'], nospace(x['name'])): int(x['section']) for x in rows('data/izumo_vs_pre_race.csv')}
MEMBERS = set(SECTION)
PLACE = {(x['year'], nospace(x['name'])): x['place'] for x in rows('data/pre_race_5000m.csv')}

rank = {'2022': 4, '2023': 4, '2024': 3, '2025': 7, '2026': 11}

wb = Workbook()

# ---------------- 出雲区間記録 ----------------
ws = wb.active
ws.title = '出雲区間記録'
title(ws, '出雲駅伝 区間記録：日本人トップと青学の比較（2022〜2026）',
      '換算タイム＝実タイム×換算距離÷区間距離（1〜5区は5km、6区は10km）。差＝換算後の青学−日本人トップ（秒）。'
      '日本人トップ＝留学生を除いた区間最速。区間順位は留学生を含む公式記録。出典：各年の公式リザルトPDF', 13)
H = ['年', '区間', '距離\n(km)', '換算距離\n(km)', '日本人トップ', '所属', '実タイム', '換算', '青学選手', '実タイム',
     '区間順位', '換算', '差（秒）']
header(ws, 4, H, [7, 6, 7, 8, 14, 12, 9, 9, 13, 9, 7, 9, 9])
iz = rows('data/izumo_sections.csv')
r = 5
for x in iz:
    put(ws, r, 1, x['year'], font=BLUE)
    put(ws, r, 2, int(x['section']), font=BLUE)
    put(ws, r, 3, float(x['distance_km']), '0.0', BLUE)
    put(ws, r, 4, f'=IF(B{r}=6,10,5)', '0')
    put(ws, r, 5, x['winner_name'], font=BLUE)
    put(ws, r, 6, x['winner_team'], font=BLUE)
    put(ws, r, 7, tval(x['winner_time']), '[m]:ss', BLUE)
    put(ws, r, 8, f'=G{r}*D{r}/C{r}', '[m]:ss')
    put(ws, r, 9, x['aogaku_name'], font=BLUE)
    put(ws, r, 10, tval(x['aogaku_time']), '[m]:ss', BLUE)
    put(ws, r, 11, int(x['aogaku_rank']), '0', BLUE)
    put(ws, r, 12, f'=J{r}*D{r}/C{r}', '[m]:ss')
    put(ws, r, 13, f'=ROUND((L{r}-H{r})*86400,1)', '+0.0;-0.0;0')
    r += 1
IZ_LAST = r - 1
band(ws, 5, IZ_LAST, 13)
ws.conditional_formatting.add(f'M5:M{IZ_LAST}', ColorScaleRule(
    start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=80, end_color='F8696B'))
ws.conditional_formatting.add(f'K5:K{IZ_LAST}', FormulaRule(
    formula=[f'$K5=1'], font=Font(name=FONT, bold=True, color='000000'), fill=PatternFill('solid', fgColor='FFEB9C')))
ws.freeze_panes = 'A5'
ws.auto_filter.ref = f'A4:M{IZ_LAST}'

# ---------------- 5000m比較 ----------------
ws = wb.create_sheet('5000m比較')
title(ws, '2週間前の5000m（5km）と出雲の比較：出雲出走メンバー',
      '出雲5km換算＝出雲の実タイム×5÷区間距離（6区も5kmあたりに換算）。悪化率＝出雲5km換算÷2週間前5000m−1。'
      '2026年はロード5km、他の年はトラック5000m。組内順位は留学生・他大学を含むその組の順位。出典：各年の記録会リザルト、出雲公式リザルト', 12)
H = ['年', '区間', '選手', '2週間前\n5000m', 'レース種別', '出雲\n実タイム', '区間順位', '区間距離\n(km)', '出雲\n5km換算',
     '差（秒）', '悪化率', '5000m\n組内順位']
header(ws, 4, H, [7, 6, 13, 10, 26, 9, 7, 8, 10, 9, 9, 9])
vp = rows('data/izumo_vs_pre_race.csv')
r = 5
for x in vp:
    put(ws, r, 1, x['year'], font=BLUE)
    put(ws, r, 2, int(x['section']), font=BLUE)
    put(ws, r, 3, nospace(x['name']), font=BLUE)
    fmt = '[m]:ss.00' if '.' in x['pre_5000m'] else '[m]:ss'
    put(ws, r, 4, tval(x['pre_5000m']), fmt, BLUE)
    put(ws, r, 5, race_label(x['pre_race']), font=BLUE)
    put(ws, r, 6, tval(x['izumo_time']), '[m]:ss', BLUE)
    put(ws, r, 7, int(x['izumo_rank']), '0', BLUE)
    put(ws, r, 8, SEC_KM[x['section']], '0.0', BLUE)
    put(ws, r, 9, f'=F{r}*5/H{r}', '[m]:ss.0')
    put(ws, r, 10, f'=ROUND((I{r}-D{r})*86400,1)', '+0.0;-0.0;0')
    put(ws, r, 11, f'=I{r}/D{r}-1', '+0.0%;-0.0%;0.0%')
    put(ws, r, 12, int(PLACE[(x['year'], nospace(x['name']))]), '0', BLUE)
    r += 1
VP_LAST = r - 1
band(ws, 5, VP_LAST, 11)
ws.conditional_formatting.add(f'K5:K{VP_LAST}', ColorScaleRule(
    start_type='num', start_value=0, start_color='63BE7B', mid_type='num', mid_value=0.05, mid_color='FFEB84',
    end_type='num', end_value=0.095, end_color='F8696B'))
ws.freeze_panes = 'A5'
ws.auto_filter.ref = f'A4:L{VP_LAST}'
MEMBER = f"SUMIFS('5000m比較'!$B$5:$B${VP_LAST},'5000m比較'!$A$5:$A${VP_LAST},{{y}},'5000m比較'!$C$5:$C${VP_LAST},{{n}})"

# ---------------- 合宿走行距離 ----------------
ws = wb.create_sheet('合宿走行距離')
title(ws, '夏合宿 走行距離（1次・2次）：出雲メンバー',
      '出雲出走メンバーのみ。出雲区間＝その年に走った区間。1次合宿の期間：2023年 7/26〜8/16（22日）、2024・2025年 8/1〜8/19（19日）、2026年 8/2〜8/19（18日）。'
      '2次合宿は各年10〜12日。出典：各年の走行距離ランキング（Excel・PDF）', 7)
header(ws, 4, ['年', '選手', '1次合宿\n(km)', '2次合宿\n(km)', '1次\n日数', '1次\n1日あたり', '出雲\n区間'],
       [7, 13, 10, 10, 7, 9, 8])
days = {'2023': 22, '2024': 19, '2025': 19, '2026': 18}
cd = sorted([x for x in rows('data/camp_distance.csv') if (x['year'], x['name']) in MEMBERS], key=lambda x: (x['year'], SECTION[(x['year'], x['name'])]))
r = 5
for x in cd:
    put(ws, r, 1, x['year'], font=BLUE)
    put(ws, r, 2, x['name'], font=BLUE)
    put(ws, r, 3, float(x['camp1_km']) if x['camp1_km'] else None, '0.0', BLUE)
    put(ws, r, 4, float(x['camp2_km']) if x['camp2_km'] else None, '0.0', BLUE)
    put(ws, r, 5, days[x['year']], '0', BLUE)
    put(ws, r, 6, f'=IF(C{r}="","",C{r}/E{r})', '0.0')
    put(ws, r, 7, f'={MEMBER.format(y=f"A{r}", n=f"B{r}")}')
    r += 1
CD_LAST = r - 1
band(ws, 5, CD_LAST, 7)
ws.freeze_panes = 'A5'
ws.auto_filter.ref = f'A4:G{CD_LAST}'

# ---------------- 合宿距離×区間順位 ----------------
ws = wb.create_sheet('合宿距離×区間順位')
title(ws, '夏合宿の走行距離と出雲の区間順位：出雲メンバー（2023〜2026）',
      '値はすべて「合宿走行距離」「5000m比較」「出雲区間記録」シートからの参照式。区間順位は留学生を含む公式記録。'
      '日本人トップとの差は1〜5区が5kmあたり、6区が10kmあたりの秒差。悪化率＝出雲5km換算÷2週間前5000m−1', 10)
header(ws, 4, ['年', '区間', '選手', '区間順位', '1次合宿\n(km)', '1次\n1日あたり\n(km)', '2次合宿\n(km)',
               '1次+2次\n合計(km)', '日本人トップ\nとの差(秒)', '悪化率'], [7, 6, 13, 8, 9, 9, 9, 9, 11, 9])
CDn, VPn, IZn = "'合宿走行距離'", "'5000m比較'", "'出雲区間記録'"
crit = lambda sh, last, r: f"{sh}!$A$5:$A${last},$A{r},{sh}!${'B' if sh == CDn else 'C'}$5:${'B' if sh == CDn else 'C'}${last},$C{r}"
r = 5
for y in ['2023', '2024', '2025', '2026']:
    for (yy, name), sec in sorted(SECTION.items(), key=lambda kv: kv[1]):
        if yy != y:
            continue
        put(ws, r, 1, y, font=BLUE)
        put(ws, r, 2, sec, '0', BLUE)
        put(ws, r, 3, name, font=BLUE)
        put(ws, r, 4, f"=SUMIFS({VPn}!$G$5:$G${VP_LAST},{crit(VPn, VP_LAST, r)})", '0')
        put(ws, r, 5, f"=SUMIFS({CDn}!$C$5:$C${CD_LAST},{crit(CDn, CD_LAST, r)})", '0.0')
        put(ws, r, 6, f"=SUMIFS({CDn}!$F$5:$F${CD_LAST},{crit(CDn, CD_LAST, r)})", '0.0')
        put(ws, r, 7, f"=SUMIFS({CDn}!$D$5:$D${CD_LAST},{crit(CDn, CD_LAST, r)})", '0.0')
        put(ws, r, 8, f'=E{r}+G{r}', '0.0')
        put(ws, r, 9, f"=SUMIFS({IZn}!$M$5:$M${IZ_LAST},{IZn}!$A$5:$A${IZ_LAST},$A{r},{IZn}!$B$5:$B${IZ_LAST},$B{r})", '+0.0;-0.0;0')
        put(ws, r, 10, f"=SUMIFS({VPn}!$K$5:$K${VP_LAST},{crit(VPn, VP_LAST, r)})", '+0.0%;-0.0%;0.0%')
        r += 1
DR_LAST = r - 1
band(ws, 5, DR_LAST, 10)
ws.conditional_formatting.add(f'D5:D{DR_LAST}', FormulaRule(
    formula=['$D5=1'], font=Font(name=FONT, bold=True, color='000000'), fill=PatternFill('solid', fgColor='FFEB9C')))
ws.conditional_formatting.add(f'F5:F{DR_LAST}', ColorScaleRule(
    start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84',
    end_type='max', end_color='F8696B'))
ws.freeze_panes = 'D5'
ws.auto_filter.ref = f'A4:J{DR_LAST}'

r = DR_LAST + 2
ws.cell(r, 1, '年ごとの6人平均').font = Font(name=FONT, bold=True, size=11, color='000000')
r += 1
header(ws, r, ['年', '出雲\n総合順位', '', '区間順位\n平均', '1次合宿\n(km)', '1次\n1日あたり\n(km)', '2次合宿\n(km)',
               '1次+2次\n合計(km)', '日本人トップ\nとの差(秒)', '悪化率'])
for y in ['2023', '2024', '2025', '2026']:
    r += 1
    put(ws, r, 1, y, font=BLUE)
    put(ws, r, 2, rank[y], '0', BLUE)
    put(ws, r, 3, None)
    for col, fmt in zip('DEFGHIJ', ['0.0', '0.0', '0.0', '0.0', '0.0', '+0.0;-0.0;0', '+0.0%;-0.0%']):
        put(ws, r, ' ABCDEFGHIJ'.index(col),
            f'=AVERAGEIFS(${col}$5:${col}${DR_LAST},$A$5:$A${DR_LAST},$A{r})', fmt)

r += 2
ws.cell(r, 1, '区間順位との相関係数（延べ24人）').font = Font(name=FONT, bold=True, size=11, color='000000')
r += 1
header(ws, r, ['指標', '', '', '', '相関係数'])
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
for lab, col in [('1次合宿の距離', 'E'), ('1次合宿の1日あたり距離', 'F'), ('2次合宿の距離', 'G'),
                 ('1次+2次の合計', 'H'), ('悪化率', 'J')]:
    r += 1
    for j in range(1, 5):
        put(ws, r, j, lab if j == 1 else None, align=LEFT)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    put(ws, r, 5, f'=CORREL({col}5:{col}{DR_LAST},D5:D{DR_LAST})', '0.00')
ws.cell(r + 1, 1, '※正の値＝その指標が大きい選手ほど区間順位が悪い。2026年は順位・距離ともに全体が高く、年の影響が混ざっている点に注意').font = Font(name=FONT, size=9, color='000000')

r += 3
ws.cell(r, 1, '区間3位以内とそれ以外の比較').font = Font(name=FONT, bold=True, size=11, color='000000')
r += 1
header(ws, r, ['区分', '', '', '人数', '1次合宿\n(km)', '1次\n1日あたり\n(km)', '2次合宿\n(km)', '1次+2次\n合計(km)'])
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
for lab, cond in [('区間1〜3位', '"<=3"'), ('区間4位以下', '">3"')]:
    r += 1
    for j in range(1, 4):
        put(ws, r, j, lab if j == 1 else None)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    put(ws, r, 4, f'=COUNTIFS($D$5:$D${DR_LAST},{cond})', '0')
    for col in 'EFGH':
        put(ws, r, ' ABCDEFGH'.index(col), f'=AVERAGEIFS({col}$5:{col}${DR_LAST},$D$5:$D${DR_LAST},{cond})', '0.0')

# ---------------- 合宿消化率 ----------------
ws = wb.create_sheet('合宿消化率')
title(ws, '夏合宿 消化率（1〜3次）：出雲メンバー',
      '出雲出走メンバーのみ。消化率はポイント練習の実施度（100%＝全メニュー完遂）。出典：各年の合宿消化率（2023・2025年はExcel、2024・2026年はPDF）', 7)
header(ws, 4, ['年', '選手', '1次合宿', '2次合宿', '3次合宿', '夏合宿\n平均', '出雲\n区間'], [7, 13, 9, 9, 9, 9, 8])
cc = sorted([x for x in rows('data/camp_completion.csv') if (x['year'], x['name']) in MEMBERS], key=lambda x: (x['year'], SECTION[(x['year'], x['name'])]))
r = 5
for x in cc:
    put(ws, r, 1, x['year'], font=BLUE)
    put(ws, r, 2, x['name'], font=BLUE)
    for j, k in enumerate(['camp1', 'camp2', 'camp3'], 3):
        put(ws, r, j, float(x[k]) if x[k] not in ('', None) else None, '0%', BLUE)
    put(ws, r, 6, f'=IFERROR(AVERAGE(C{r}:E{r}),"")', '0.0%')
    put(ws, r, 7, f'={MEMBER.format(y=f"A{r}", n=f"B{r}")}')
    r += 1
CC_LAST = r - 1
band(ws, 5, CC_LAST, 7)
ws.conditional_formatting.add(f'C5:F{CC_LAST}', ColorScaleRule(
    start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=0.7, mid_color='FFEB84',
    end_type='num', end_value=1, end_color='63BE7B'))
ws.freeze_panes = 'A5'
ws.auto_filter.ref = f'A4:G{CC_LAST}'

# ---------------- 月間走行距離 ----------------
ws = wb.create_sheet('月間走行距離')
title(ws, '月間走行距離：出雲メンバー（2022・2023年は通年、2024年は1〜6月のみ）',
      '出雲出走メンバーのみ。出典：月間走行距離Excel。2024年7月以降と2025・2026年はデータ未提供', 16)
header(ws, 4, ['年', '選手'] + [f'{m}月' for m in range(1, 13)] + ['9月÷8月', '出雲\n区間'],
       [7, 13] + [7] * 12 + [9, 8])
md = defaultdict(dict)
for x in rows('data/monthly_distance.csv'):
    if (x['year'], x['name']) not in MEMBERS:
        continue
    md[(x['year'], x['name'])][int(x['month'])] = float(x['km'])
r = 5
for (y, n), mv in sorted(md.items(), key=lambda kv: (kv[0][0], SECTION[kv[0]])):
    put(ws, r, 1, y, font=BLUE)
    put(ws, r, 2, n, font=BLUE)
    for m in range(1, 13):
        put(ws, r, 2 + m, mv.get(m), '0', BLUE)
    put(ws, r, 15, f'=IF(OR(J{r}="",K{r}="",J{r}=0),"",K{r}/J{r})', '0.00')
    put(ws, r, 16, f'={MEMBER.format(y=f"A{r}", n=f"B{r}")}')
    r += 1
MD_LAST = r - 1
band(ws, 5, MD_LAST, 16)
ws.freeze_panes = 'C5'
ws.auto_filter.ref = f'A4:P{MD_LAST}'

# ---------------- 合宿メニュー（年別） ----------------
md_text = open('docs/camp_workouts_izumo_members.md', encoding='utf-8').read()
notes = [l[2:] for l in md_text.split('## 2023年')[0].splitlines() if l.startswith('- ')]
for block in re.split(r'\n## ', md_text)[1:]:
    lines = block.splitlines()
    year = lines[0][:4]
    ws = wb.create_sheet(f'合宿メニュー_{year}')
    title(ws, f'{year}年 夏合宿ポイント練習：出雲メンバー6人の組・消化率')
    for k, n in enumerate(notes, 2):
        ws.cell(k, 1, '※ ' + n).font = Font(name=FONT, size=9, color='000000')
    r = len(notes) + 3
    ws.cell(r, 1, lines[0] + '　' + next(l for l in lines if l.startswith('メンバー'))).font = Font(name=FONT, bold=True, size=10)
    r += 2
    widths = [9, 46] + [15] * 6
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    i = 0
    while i < len(lines):
        l = lines[i]
        if l.startswith('### '):
            c = ws.cell(r, 1, l[4:])
            c.font = Font(name=FONT, bold=True, size=11, color='000000')
            r += 1
        elif l.startswith('|日付'):
            labs = l.strip('|').split('|')
            for j, lab in enumerate(labs, 1):
                c = ws.cell(r, j, lab)
                c.font, c.fill, c.alignment, c.border = HEAD_FONT, HEAD_FILL, CENTER, BORDER
            r += 1
            i += 2
            while i < len(lines) and lines[i].startswith('|'):
                vals = lines[i].strip('|').split('|')
                for j, v in enumerate(vals, 1):
                    cell = put(ws, r, j, v, align=LEFT if j == 2 else CENTER)
                    if re.search(r'(^|\s)(0|30|50|60|70)%$', v):
                        cell.fill = PatternFill('solid', fgColor='FCE4D6')
                ws.row_dimensions[r].height = 15 * max(2, max(-(-len(v) // 14) for v in vals[2:]) if vals[2:] else 2)
                r += 1
                i += 1
            r += 1
            continue
        i += 1
    ws.freeze_panes = 'C1'

# ---------------- 年別まとめ ----------------
ws = wb.create_sheet('年別まとめ', 0)
title(ws, '年別まとめ：出雲の成績と、2週間前5000m・合宿データの比較',
      'メンバー＝その年の出雲出走6人。順位・チーム平均・日付は元資料からの入力値（出典は各セルのコメント）、それ以外は他シートからの集計式。'
      '空欄はデータ未提供', 17)
H = ['年', '出雲\n総合順位', '1〜5区\n平均差\n(秒/5km)', '6区差\n(秒/10km)', 'メンバー\n5000m平均', '悪化率\n平均',
     'メンバー\n1次距離\n(km)', 'チーム\n1次平均\n(km)', '1次距離\nチーム比', 'メンバー\n2次距離\n(km)',
     'メンバー\n夏消化率', 'メンバー\n8月月間\n(km)', 'メンバー\n9月月間\n(km)', '9月÷8月', '3次合宿\n最終日', '出雲\n開催日',
     '3次終了→\n出雲(日)']
header(ws, 4, H, [7, 8, 9, 9, 10, 8, 9, 9, 9, 9, 9, 9, 9, 8, 10, 10, 9])
team1 = {'2023': 507.9, '2024': 510.6, '2025': 481, '2026': 460}
end3 = {'2023': '2023-09-17', '2024': '2024-09-21', '2025': '2025-09-20', '2026': '2026-09-23'}
izdate = {'2022': '2022-10-10', '2023': '2023-10-09', '2024': '2024-10-14', '2025': '2025-10-13', '2026': '2026-10-12'}
IZ, VP, CDs, CCs, MDs = "'出雲区間記録'", "'5000m比較'", "'合宿走行距離'", "'合宿消化率'", "'月間走行距離'"
import datetime as dt
for k, y in enumerate(YEARS):
    r = 5 + k
    put(ws, r, 1, y, font=BLUE)
    c = put(ws, r, 2, rank[y], '0', BLUE)
    c.comment = Comment('出典：各年の出雲駅伝公式リザルトPDF', 'analysis')
    put(ws, r, 3, f'=IFERROR(AVERAGEIFS({IZ}!$M$5:$M${IZ_LAST},{IZ}!$A$5:$A${IZ_LAST},A{r},{IZ}!$B$5:$B${IZ_LAST},"<6"),"")', '+0.0;-0.0;0')
    put(ws, r, 4, f'=IFERROR(AVERAGEIFS({IZ}!$M$5:$M${IZ_LAST},{IZ}!$A$5:$A${IZ_LAST},A{r},{IZ}!$B$5:$B${IZ_LAST},6),"")', '+0.0;-0.0;0')
    put(ws, r, 5, f'=IFERROR(AVERAGEIFS({VP}!$D$5:$D${VP_LAST},{VP}!$A$5:$A${VP_LAST},A{r}),"")', '[m]:ss.0')
    put(ws, r, 6, f'=IFERROR(AVERAGEIFS({VP}!$K$5:$K${VP_LAST},{VP}!$A$5:$A${VP_LAST},A{r}),"")', '+0.0%;-0.0%')
    put(ws, r, 7, f'=IFERROR(AVERAGEIFS({CDs}!$C$5:$C${CD_LAST},{CDs}!$A$5:$A${CD_LAST},A{r},{CDs}!$G$5:$G${CD_LAST},">0"),"")', '0.0')
    if y in team1:
        c = put(ws, r, 8, team1[y], '0.0', BLUE)
        c.comment = Comment('出典：各年の1次合宿走行距離ランキングに記載の「チーム平均（全体平均）」', 'analysis')
    else:
        put(ws, r, 8, None)
    put(ws, r, 9, f'=IFERROR(G{r}/H{r},"")', '0%')
    put(ws, r, 10, f'=IFERROR(AVERAGEIFS({CDs}!$D$5:$D${CD_LAST},{CDs}!$A$5:$A${CD_LAST},A{r},{CDs}!$G$5:$G${CD_LAST},">0"),"")', '0.0')
    put(ws, r, 11, f'=IFERROR(AVERAGEIFS({CCs}!$F$5:$F${CC_LAST},{CCs}!$A$5:$A${CC_LAST},A{r},{CCs}!$G$5:$G${CC_LAST},">0"),"")', '0.0%')
    put(ws, r, 12, f'=IFERROR(AVERAGEIFS({MDs}!$J$5:$J${MD_LAST},{MDs}!$A$5:$A${MD_LAST},A{r},{MDs}!$P$5:$P${MD_LAST},">0"),"")', '0')
    put(ws, r, 13, f'=IFERROR(AVERAGEIFS({MDs}!$K$5:$K${MD_LAST},{MDs}!$A$5:$A${MD_LAST},A{r},{MDs}!$P$5:$P${MD_LAST},">0"),"")', '0')
    put(ws, r, 14, f'=IFERROR(M{r}/L{r},"")', '0.00')
    if y in end3:
        c = put(ws, r, 15, dt.date.fromisoformat(end3[y]), 'm/d', BLUE)
        c.comment = Comment('出典：各年の3次合宿消化率（選抜組の最終ポイント練習日）。2022年は未提供', 'analysis')
    else:
        put(ws, r, 15, None)
    put(ws, r, 16, dt.date.fromisoformat(izdate[y]), 'm/d', BLUE)
    put(ws, r, 17, f'=IF(O{r}="","",P{r}-O{r})', '0')
YR_LAST = 5 + len(YEARS) - 1
ws.conditional_formatting.add(f'F5:F{YR_LAST}', ColorScaleRule(
    start_type='min', start_color='63BE7B', end_type='max', end_color='F8696B'))
ws.conditional_formatting.add(f'B5:B{YR_LAST}', ColorScaleRule(
    start_type='min', start_color='63BE7B', end_type='max', end_color='F8696B'))
ws.conditional_formatting.add(f'I5:I{YR_LAST}', ColorScaleRule(
    start_type='min', start_color='63BE7B', end_type='max', end_color='F8696B'))

# 相関
r = YR_LAST + 2
ws.cell(r, 1, '年単位の相関係数（データのある年のみ。年数が少ないため参考値）').font = Font(name=FONT, bold=True, size=11, color='000000')
r += 1
header(ws, r, ['組み合わせ', '', '', '', '対象年', '相関係数'])
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
pairs = [
    ('メンバー5000m平均 × 悪化率平均', f'E5:E{YR_LAST}', f'F5:F{YR_LAST}', '2022〜26年'),
    ('メンバー5000m平均 × 総合順位', f'E5:E{YR_LAST}', f'B5:B{YR_LAST}', '2022〜26年'),
    ('悪化率平均 × 総合順位', f'F5:F{YR_LAST}', f'B5:B{YR_LAST}', '2022〜26年'),
    ('1次距離チーム比 × 総合順位', f'I6:I{YR_LAST}', f'B6:B{YR_LAST}', '2023〜26年'),
    ('1次距離チーム比 × 悪化率平均', f'I6:I{YR_LAST}', f'F6:F{YR_LAST}', '2023〜26年'),
]
for lab, a, b, span in pairs:
    r += 1
    for j in range(1, 5):
        put(ws, r, j, lab if j == 1 else None, align=LEFT)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    put(ws, r, 5, span)
    put(ws, r, 6, f'=CORREL({a},{b})', '0.00')

# 区間別
r += 2
ws.cell(r, 1, '区間別の平均（5年間）').font = Font(name=FONT, bold=True, size=11, color='000000')
r += 1
header(ws, r, ['区間', '距離\n(km)', '日本人\nトップとの\n平均差(秒)', '悪化率\n平均', '区間賞\n回数'])
for s in range(1, 7):
    r += 1
    put(ws, r, 1, s, '0', BLUE)
    put(ws, r, 2, SEC_KM[str(s)], '0.0', BLUE)
    put(ws, r, 3, f'=AVERAGEIFS({IZ}!$M$5:$M${IZ_LAST},{IZ}!$B$5:$B${IZ_LAST},A{r})', '+0.0;-0.0;0')
    put(ws, r, 4, f'=AVERAGEIFS({VP}!$K$5:$K${VP_LAST},{VP}!$B$5:$B${VP_LAST},A{r})', '+0.0%;-0.0%')
    put(ws, r, 5, f'=COUNTIFS({IZ}!$B$5:$B${IZ_LAST},A{r},{IZ}!$K$5:$K${IZ_LAST},1)', '0')
ws.cell(r + 1, 1, '※6区の平均差は10kmあたり。悪化率は全区間とも5kmあたりで比較').font = Font(name=FONT, size=9, color='000000')
ws.freeze_panes = 'B5'

# ---------------- 概要 ----------------
ws = wb.create_sheet('概要', 0)
ws.column_dimensions['A'].width = 24
ws.column_dimensions['B'].width = 100
ws['A1'] = '出雲駅伝 振り返り分析（2022〜2026年）'
ws['A1'].font = Font(name=FONT, bold=True, size=16, color='000000')
ws['A2'] = '目的：近年ピーキングが合わず順位が伸びない原因を、2週間前の5000mと夏合宿のデータから探る'
ws['A2'].font = Font(name=FONT, size=10)
r = 4
ws.cell(r, 1, 'シート一覧').font = Font(name=FONT, bold=True, size=12, color='000000')
sheets = [
    ('年別まとめ', '年ごとの総合順位・区間差・悪化率・合宿の負荷を1表に集約。相関係数と区間別平均つき'),
    ('出雲区間記録', '各区間の日本人トップと青学選手のタイム、5km（6区は10km）換算と差'),
    ('5000m比較', '出雲メンバーの2週間前5000m（5km）と組内順位、出雲5km換算ペース、悪化率'),
    ('合宿走行距離', '出雲メンバーの1次・2次合宿の距離（2023〜2026年）'),
    ('合宿距離×区間順位', '1次・2次合宿の距離と区間順位・日本人トップとの差・悪化率を1表に。年平均、相関係数、区間3位以内との比較つき'),
    ('合宿消化率', '出雲メンバーの1〜3次合宿の消化率（2023〜2026年）'),
    ('月間走行距離', '出雲メンバーの月間走行距離（2022・2023年通年、2024年1〜6月）'),
    ('合宿メニュー_2023〜2026', '1〜3次合宿のポイント練習と、出雲メンバー6人の組・消化率'),
]
for n, d in sheets:
    r += 1
    put(ws, r, 1, n, align=LEFT).font = Font(name=FONT, bold=True, size=10)
    put(ws, r, 2, d, align=LEFT)
r += 2
ws.cell(r, 1, '用語と計算方法').font = Font(name=FONT, bold=True, size=12, color='000000')
terms = [
    ('換算タイム', '実タイム×換算距離÷区間距離。1〜5区は5km、6区は10kmあたり'),
    ('日本人トップ', '留学生を除いた区間最速の選手（区間賞が留学生だったのは5年とも3区のみ）'),
    ('悪化率', '出雲の5km換算ペース÷2週間前5000m記録−1。大きいほど本番で落ちている'),
    ('チーム比', '出雲メンバー6人の1次合宿距離の平均÷チーム全体の平均'),
    ('対象', '出雲駅伝に出走した青学の選手（各年6人）のみ。比較用に日本人トップの区間記録と、1次合宿のチーム平均距離（年別まとめ）だけは参考値として掲載'),
    ('色の意味', '緑〜赤の背景色は良い〜悪い。区間賞は黄色の背景'),
]
for n, d in terms:
    r += 1
    put(ws, r, 1, n, align=LEFT).font = Font(name=FONT, bold=True, size=10)
    put(ws, r, 2, d, align=LEFT)
r += 2
ws.cell(r, 1, 'これまでの分析で見えたこと（仮説）').font = Font(name=FONT, bold=True, size=12, color='000000')
finds = [
    '2週間前の5000mが速い年ほど出雲での悪化率が大きい（年単位で強い負の相関）。個人単位では関係がほぼない',
    '総合3位の2024年だけ、出雲メンバーの1次合宿距離がチーム平均を下回った（93%）。他の年は113〜121%',
    '3次合宿の骨組み（距離走→変化走→スピード）と消化率は4年とも大差なく、2024年も最終日は強い練習だった',
    '2026年は3次合宿の最終日が9/23と遅く、出雲まで19日（他の年は22〜23日）。最後の練習の4日後の5kmで6人中5人が自己ベスト',
    '注意：年単位の比較は4〜5年分しかなく、統計的に確かとは言えない',
]
for d in finds:
    r += 1
    put(ws, r, 1, '・', align=CENTER)
    put(ws, r, 2, d, align=LEFT)
r += 2
ws.cell(r, 1, '未提供のデータ').font = Font(name=FONT, bold=True, size=12, color='000000')
for d in ['2023〜2025年の5000m記録会の日付', '3次合宿終了から出雲までの練習内容（2024〜2026年、2023年の9/18〜9/29）',
          '2024年7〜10月、2025・2026年の月間走行距離', '2022年の合宿データ（走行距離・消化率・3次合宿日程）']:
    r += 1
    put(ws, r, 1, '・', align=CENTER)
    put(ws, r, 2, d, align=LEFT)

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
    s.sheet_properties.tabColor = GREEN if s.title in ('概要', '年別まとめ') else '7FA99B'
    s.page_setup.orientation = 'landscape'
    s.page_setup.paperSize = s.PAPERSIZE_A4
    s.page_setup.fitToWidth, s.page_setup.fitToHeight = 1, 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(sys.argv[1])
print('saved', sys.argv[1])
