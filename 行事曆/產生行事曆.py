# -*- coding: utf-8 -*-
"""2027 行事曆（更新版）：華嚴海會 2027/12/11–12/22，共 12 天。
由原始 PDF 逐格擷取既有內容，只調整華嚴海會相關排程後重建 Excel／PDF。"""
import json, re, calendar, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT = "Microsoft JhengHei"
NAVY   = "1F3864"
HEADBG = "D6E4F0"
CREAM  = "FDF2D5"
HILITE = "FFE2E2"
GREY   = "F2F2F2"
thin = Side(style="thin", color="B7C4D6")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------- 1. 讀入由原 PDF 擷取的每日內容，修補缺字 ----------
cells = {int(m): {int(d): v for d, v in days.items()}
         for m, days in json.load(open("cells.json", encoding="utf-8")).items()}

FIX = {
    '企業專班：善': '企業專班：善首',
    '企業專班：□雄': '企業專班：高雄',
    '禪（新加坡）': '一日禪（新加坡）',
    '英□□□禪（新加坡）': '英文一日禪（新加坡）',
    '英□楞嚴經（新加坡）': '英文楞嚴經（新加坡）',
    '歲末□型華嚴法會': '歲末大型華嚴法會',
    '清明法會放蒙□（台北）': '清明法會放蒙山（台北）',
    '清明法會放蒙□（台中）': '清明法會放蒙山（台中）',
    '清明法會放蒙□（板橋）': '清明法會放蒙山（板橋）',
    '清明法會放蒙□（□雄）': '清明法會放蒙山（高雄）',
    '彰化供佛齋僧□會': '彰化供佛齋僧大會',
    '林□供佛齋僧□會': '林口供佛齋僧大會',
}
for n in range(1, 7):
    FIX['兒童夏令營第%d梯' % n] = '兒童夏令營 第%d梯' % n

def fix(label, m, d):
    if label in FIX:
        return FIX[label]
    if label == '僧眾禪五第□梯次':
        return '僧眾禪五第一梯次' if d <= 15 else '僧眾禪五第二梯次'
    if label == '夏安居禪七第□梯次':
        return '夏安居禪七第一梯次' if m == 6 else '夏安居禪七第二梯次'
    assert '□' not in label, (label, m, d)
    return label

for m in cells:
    for d, v in cells[m].items():
        v['ev'] = [fix(e, m, d) for e in v['ev']]

# ---------- 2. 本次更新：華嚴海會 2027/12/11–12/22（共 12 天） ----------
OLD_EVENT = '歲末大型華嚴法會'
NEW_EVENT = '華嚴海會'
HY_START, HY_END = datetime.date(2027, 12, 11), datetime.date(2027, 12, 22)

removed, added = [], []
for m in cells:
    for d, v in cells[m].items():
        if OLD_EVENT in v['ev']:
            v['ev'] = [e for e in v['ev'] if e != OLD_EVENT]
            removed.append('2027/%d/%d' % (m, d))
cancelled = []          # [(date, [被取消的活動])]
for d in range((HY_END - HY_START).days + 1):
    day = HY_START + datetime.timedelta(days=d)
    cell = cells[day.month][day.day]
    others = [e for e in cell['ev'] if e != NEW_EVENT]
    if others:
        cancelled.append((day, others))
    # 華嚴海會為重大活動：期間內其餘活動全數取消
    cell['ev'] = [NEW_EVENT]
    added.append(day.strftime('2027/%-m/%-d'))
assert len(added) == 12, added
print('移除 %s：%s' % (OLD_EVENT, '、'.join(removed)))
print('新增 %s：%s（共 %d 天）' % (NEW_EVENT, '、'.join(added), len(added)))
print('華嚴海會期間取消之衝突活動：')
for day, evs in cancelled:
    print('   %s  %s' % (day.strftime('2027/%-m/%-d'), '、'.join(evs)))

WD = '日一二三四五六'
def wd(dt):
    return '星期' + WD[(dt.weekday() + 1) % 7]

# ---------- 3. 活動總表 ----------
COURSES = [
    ('兒童哲學班（春季）', '共 16 堂', '2027/3/7、2027/3/14、2027/3/21、2027/3/28、2027/4/4、2027/4/11、2027/4/18、2027/4/25、2027/5/2、2027/5/9、2027/5/16、2027/5/23、2027/5/30、2027/6/6、2027/6/13、2027/6/20'),
    ('兒童哲學班（秋季）', '共 12 堂', '2027/9/26、2027/10/3、2027/10/10、2027/10/17、2027/10/24、2027/10/31、2027/11/7、2027/11/14、2027/11/21、2027/11/28、2027/12/5、2027/12/26（原 14 堂，12/12、12/19 因華嚴海會取消）'),
    ('春季班', '共 17 週', '2027/3/2–2027/6/18（週一至週五晚上，首週自週二起）'),
    ('暑期班', '共 5 週',  '2027/7/12–2027/8/13（週一至週五晚上）'),
    ('秋季班', '共 15 週', '2027/9/13–2027/12/24（週一至週五晚上）；12/11–12/22 華嚴海會期間停課，實際上課 9/13–12/10、12/23–12/24'),
    ('企業專班：台北', '共 8 堂', '2027/3/6、2027/4/3、2027/5/1、2027/6/5、2027/8/7、2027/10/2、2027/11/6、2027/12/4'),
    ('企業專班：高雄', '共 8 堂', '2027/3/7、2027/4/4、2027/5/2、2027/6/6、2027/8/1、2027/10/3、2027/11/7、2027/12/5'),
    ('企業專班：台中', '共 7 堂', '2027/3/14、2027/4/11、2027/5/9、2027/6/13、2027/8/8、2027/10/10、2027/11/14（原 8 堂，12/12 因華嚴海會取消）'),
    ('好好讀楞嚴：板橋', '共 6 堂', '2027/3/13、2027/4/10、2027/5/8、2027/6/12、2027/10/9、2027/11/13（原 7 堂，12/11 因華嚴海會取消）'),
    ('英文一日禪（新加坡）', '共 8 堂', '2027/3/13、2027/4/3、2027/5/1、2027/6/5、2027/8/7、2027/10/2、2027/11/6、2027/12/4'),
    ('英文楞嚴經（新加坡）', '共 8 堂', '2027/3/14、2027/4/4、2027/5/2、2027/6/6、2027/8/1、2027/10/3、2027/11/7、2027/12/5'),
    ('一日禪（新加坡）', '共 7 堂', '2027/3/20、2027/4/17、2027/5/15、2027/7/17、2027/9/18、2027/10/16、2027/11/20（原 8 堂，12/18 因華嚴海會取消）'),
    ('楞嚴經（新加坡）', '共 7 堂', '2027/3/21、2027/4/18、2027/5/16、2027/7/18、2027/9/19、2027/10/17、2027/11/21（原 8 堂，12/19 因華嚴海會取消）'),
]

EVENTS = [
    ('2027/1/1、2027/1/2、2027/1/3、2027/1/9、2027/1/10、2027/1/16、2027/1/17、2027/1/23、2027/1/24、2027/1/30、2027/1/31',
     '星期五～星期日', '華嚴法會', '法會／大型活動', '10:00–17:00', '各分院', '1/1，及其後每週六、日；共 11 天'),
    ('2027/1/11–2027/1/15', '星期一～星期五', '僧眾禪五第一梯次', '僧眾修行', '待確認', '待確認', ''),
    ('2027/1/18–2027/1/22', '星期一～星期五', '僧眾禪五第二梯次', '僧眾修行', '待確認', '待確認', ''),
    ('2027/2/6', '星期六', '新春上燈法會', '法會／大型活動', '待確認', '各分院', '農曆正月初一；依慣例辦理'),
    ('2027/2/7–2027/2/14', '星期日～星期日', '冬安居禪七', '禪七／禪修', '待確認', '待確認', '大眾初二至初九'),
    ('2027/3/2–2027/6/18', '星期一～星期五（晚上）', '春季班', '一般課程', '待確認', '各分院', '共 17 週'),
    ('2027/3/13–2027/3/21', '星期六～星期日', '新加坡禪七', '禪七／禪修', '待確認', '新加坡', ''),
    ('2027/4/3', '星期六', '清明法會放蒙山（台中）', '分院活動', '待確認', 'EPIC 禪藝實相人文空間（台中）', ''),
    ('2027/4/4', '星期日', '清明法會放蒙山（台北）', '分院活動', '待確認', '善首講堂', ''),
    ('2027/4/10', '星期六', '清明法會放蒙山（高雄）', '分院活動', '待確認', '圓道禪寺（高雄）', ''),
    ('2027/4/11', '星期日', '清明法會放蒙山（板橋）', '分院活動', '待確認', '善覺講堂', ''),
    ('2027/5/8–2027/5/16', '星期六～星期日', '浴佛節', '法會／大型活動', '待確認', '各分院', '啟建'),
    ('2027/6/23–2027/6/30', '依各列日期', '夏安居禪七第一梯次', '禪七／禪修', '待確認', '待確認', ''),
    ('2027/7/2–2027/7/9', '依各列日期', '夏安居禪七第二梯次', '禪七／禪修', '待確認', '待確認', ''),
    ('2027/7/12–2027/8/13', '星期一～星期五（晚上）', '暑期班', '一般課程', '待確認', '各分院', ''),
    ('2027/8/8', '星期日', '彰化供佛齋僧大會', '供佛齋僧', '待確認', '彰化', ''),
    ('2027/8/21–2027/8/28', '星期六～星期六', '盂蘭盆暨梁皇法會', '法會／大型活動', '待確認', '各分院', ''),
    ('2027/8/29', '星期日', '林口供佛齋僧大會', '供佛齋僧', '待確認', '林口', ''),
    ('2027/9/13–2027/12/24（12/11–12/22 停課）', '星期一～星期五（晚上）', '秋季班', '一般課程', '待確認', '各分院',
     '原共 15 週；華嚴海會期間（12/11–12/22）停課，實際上課 9/13–12/10、12/23–12/24'),
    ('2027/9/28', '星期二', '教師禪修營', '禪七／禪修', '待確認', '待確認', '新增活動'),
    ('2027/12/11–2027/12/22', '%s～%s' % (wd(HY_START), wd(HY_END)), '華嚴海會', '法會／大型活動', '待確認', '各分院',
     '共 12 天（本次更新）；原「歲末大型華嚴法會」2027/12/25–2028/1/2 調整為本案；本期間為重大活動，原有之其他活動均已取消'),
    ('2027/3/7 起每週日', '每週日', '兒童哲學班（春季）', '兒童教育', '待確認', '各分院', '共 16 堂；日期詳見上方課程場次表'),
    ('2027/9/26 起每週日', '每週日', '兒童哲學班（秋季）', '兒童教育', '待確認', '各分院', '共 12 堂（原 14 堂；12/12、12/19 因華嚴海會取消）；日期詳見上方課程場次表'),
    ('2027/7/11–2027/7/17', '星期日～星期六', '兒童夏令營 第1梯', '兒童教育', '待確認', '各分院', '第 1 梯次'),
    ('2027/7/18–2027/7/24', '星期日～星期六', '兒童夏令營 第2梯', '兒童教育', '待確認', '各分院', '第 2 梯次'),
    ('2027/7/25–2027/7/31', '星期日～星期六', '兒童夏令營 第3梯', '兒童教育', '待確認', '各分院', '第 3 梯次'),
    ('2027/8/1–2027/8/7',   '星期日～星期六', '兒童夏令營 第4梯', '兒童教育', '待確認', '各分院', '第 4 梯次'),
    ('2027/8/8–2027/8/14',  '星期日～星期六', '兒童夏令營 第5梯', '兒童教育', '待確認', '各分院', '第 5 梯次'),
    ('2027/8/15–2027/8/20', '星期日～星期五', '兒童夏令營 第6梯', '兒童教育', '待確認', '各分院', '第 6 梯次'),
    ('2027/1/2、2027/2/6、2027/3/6、2027/4/3、2027/5/1、2027/6/5、2027/7/3、2027/8/7、2027/9/4、2027/10/2、2027/11/6、2027/12/4',
     '依各列日期', '企業專班：台北', '企業專班', '待確認', '善首講堂', '每月第一個星期六'),
    ('2027/1/3、2027/2/7、2027/3/7、2027/4/4、2027/5/2、2027/6/6、2027/7/4、2027/8/1、2027/9/5、2027/10/3、2027/11/7、2027/12/5',
     '依各列日期', '企業專班：高雄', '企業專班', '待確認', '圓道禪寺（高雄）', '每月第一個星期日'),
    ('2027/1/10、2027/2/14、2027/3/14、2027/4/11、2027/5/9、2027/6/13、2027/7/11、2027/8/8、2027/9/12、2027/10/10、2027/11/14',
     '依各列日期', '企業專班：台中', '企業專班', '待確認', 'EPIC 禪藝實相人文空間（台中）', '每月第二個星期日；12/12 因華嚴海會取消'),
    ('2027/1/9、2027/2/13、2027/3/13、2027/4/10、2027/5/8、2027/6/12、2027/7/10、2027/8/14、2027/9/11、2027/10/9、2027/11/13',
     '依各列日期', '好好讀楞嚴（善覺）', '一般課程', '待確認', '善覺講堂', '每月第二週週六；12/11 因華嚴海會取消'),
]

# ---------- 4. 建立活頁簿 ----------
wb = Workbook()

def style(c, *, size=10, bold=False, color="000000", bg=None, h='left', v='top', wrap=True, border=True):
    c.font = Font(name=FONT, size=size, bold=bold, color=color)
    if bg: c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if border: c.border = BORDER
    return c

# --- 工作表 1：行事曆 ---
ws = wb.active
ws.title = '2027行事曆'
for i in range(1, 8):
    ws.column_dimensions[get_column_letter(i)].width = 16
r = 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
style(ws.cell(r, 1, '2027 年全年行事曆（更新版）'), size=18, bold=True, h='left', border=False)
ws.row_dimensions[r].height = 30; r += 1
style(ws.cell(r, 1, '每日顯示農曆日期；本版更新：華嚴海會 2027/12/11–2027/12/22，共 12 天；該期間原排定之其他活動均已取消。'),
      size=9, color="7F7F7F", border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
ws.row_dimensions[r].height = 18; r += 2

month_start_rows = []
for m in range(1, 13):
    month_start_rows.append(r)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    style(ws.cell(r, 1, '%d 月' % m), size=13, bold=True, color="FFFFFF", bg=NAVY, h='center', v='center')
    ws.row_dimensions[r].height = 24; r += 1
    for i, w in enumerate(WD):
        style(ws.cell(r, i + 1, w), bold=True, bg=HEADBG, h='center', v='center')
    ws.row_dimensions[r].height = 20; r += 1

    first_col = (calendar.weekday(2027, m, 1) + 1) % 7
    nd = calendar.monthrange(2027, m)[1]
    grid = [[None] * 7 for _ in range((first_col + nd + 6) // 7)]
    for d in range(1, nd + 1):
        idx = first_col + d - 1
        grid[idx // 7][idx % 7] = d
    for week in grid:
        maxlines = 1
        for col, d in enumerate(week):
            c = ws.cell(r, col + 1)
            if d is None:
                style(c, bg=GREY)
                continue
            v = cells[m][d]
            lines = ['%d（農%s）' % (d, v['lunar'])] + v['ev']
            maxlines = max(maxlines, len(lines))
            c.value = '\n'.join(lines)
            hl = (m == 12 and HY_START.day <= d <= HY_END.day)
            style(c, bg=HILITE if hl else None)
            if hl:
                c.font = Font(name=FONT, size=10, bold=True, color="9C0006")
        ws.row_dimensions[r].height = max(34, 13.5 * maxlines + 6)
        r += 1
    r += 1

ws.freeze_panes = 'A4'
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.35
ws.page_margins.header = ws.page_margins.footer = 0.15
ws.page_setup.orientation = 'portrait'
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.print_area = 'A1:G%d' % (r - 1)
for i in range(3, 12, 3):            # 每 3 個月換頁
    ws.row_breaks.append(__import__('openpyxl').worksheet.pagebreak.Break(id=month_start_rows[i] - 1))

# --- 工作表 2：活動總表 ---
ws2 = wb.create_sheet('活動總表')
for col, w in zip('ABCDEFG', (34, 12, 24, 16, 14, 30, 42)):
    ws2.column_dimensions[col].width = w
r = 1
style(ws2.cell(r, 1, '2027 年活動總表（更新版）'), size=16, bold=True, border=False)
ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7); ws2.row_dimensions[r].height = 26; r += 2

style(ws2.cell(r, 1, '一、課程場次'), size=12, bold=True, border=False); r += 1
for i, h in enumerate(['課程', '堂數／週數', '日期']):
    style(ws2.cell(r, i + 1, h), bold=True, bg=HEADBG, h='center')
ws2.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7); r += 1
for name, n, dates in COURSES:
    style(ws2.cell(r, 1, name), bg=CREAM)
    style(ws2.cell(r, 2, n), bg=CREAM, h='center')
    style(ws2.cell(r, 3, dates), bg=CREAM)
    ws2.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws2.row_dimensions[r].height = 30; r += 1
r += 1

style(ws2.cell(r, 1, '二、活動明細'), size=12, bold=True, border=False); r += 1
hdr_row = r
for i, h in enumerate(['日期／期間', '星期', '活動', '類別', '時間', '地點', '備註']):
    style(ws2.cell(r, i + 1, h), bold=True, color="FFFFFF", bg=NAVY, h='center', v='center')
ws2.row_dimensions[r].height = 22; r += 1
for row in EVENTS:
    hl = row[2] == NEW_EVENT
    for i, val in enumerate(row):
        c = style(ws2.cell(r, i + 1, val), bg=HILITE if hl else None)
        if hl: c.font = Font(name=FONT, size=10, bold=True, color="9C0006")
    ws2.row_dimensions[r].height = 34 if len(row[0]) < 60 else 46
    r += 1
ws2.freeze_panes = 'A%d' % (hdr_row + 1)
ws2.page_margins.left = ws2.page_margins.right = 0.3
ws2.page_margins.top = ws2.page_margins.bottom = 0.35
ws2.page_setup.orientation = 'landscape'
ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
ws2.sheet_properties.pageSetUpPr.fitToPage = True
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0
ws2.print_area = 'A1:G%d' % (r - 1)
ws2.print_title_rows = '%d:%d' % (hdr_row, hdr_row)

# --- 工作表 3：活動重疊檢查 ---
ws3 = wb.create_sheet('活動重疊檢查')
for col, w in zip('ABCD', (16, 10, 12, 76)):
    ws3.column_dimensions[col].width = w
r = 1
style(ws3.cell(r, 1, '2027 年活動重疊檢查'), size=16, bold=True, border=False)
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); ws3.row_dimensions[r].height = 26; r += 1
style(ws3.cell(r, 1, '列出同一日安排 2 項（含）以上活動之日期。華嚴海會（12/11–12/22）期間之衝突活動已全數取消，故該期間不再出現於本表。'),
      size=9, color="7F7F7F", border=False)
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); r += 2
for i, h in enumerate(['日期', '星期', '活動數', '活動內容']):
    style(ws3.cell(r, i + 1, h), bold=True, color="FFFFFF", bg=NAVY, h='center', v='center')
hdr3 = r; ws3.row_dimensions[r].height = 22; r += 1
n_overlap = 0
for m in range(1, 13):
    for d in sorted(cells[m]):
        ev = cells[m][d]['ev']
        if len(ev) < 2:
            continue
        n_overlap += 1
        dt = datetime.date(2027, m, d)
        hl = HY_START <= dt <= HY_END
        vals = ['2027/%d/%d' % (m, d), wd(dt), len(ev), '、'.join(ev)]
        for i, val in enumerate(vals):
            c = style(ws3.cell(r, i + 1, val), bg=HILITE if hl else None,
                      h='center' if i in (0, 1, 2) else 'left')
            if hl: c.font = Font(name=FONT, size=10, bold=True, color="9C0006")
        ws3.row_dimensions[r].height = 20; r += 1
ws3.freeze_panes = 'A%d' % (hdr3 + 1)
ws3.page_margins.left = ws3.page_margins.right = 0.3
ws3.page_margins.top = ws3.page_margins.bottom = 0.35
ws3.page_setup.orientation = 'landscape'
ws3.page_setup.paperSize = ws3.PAPERSIZE_A4
ws3.sheet_properties.pageSetUpPr.fitToPage = True
ws3.page_setup.fitToWidth = 1
ws3.page_setup.fitToHeight = 0
ws3.print_area = 'A1:D%d' % (r - 1)
ws3.print_title_rows = '%d:%d' % (hdr3, hdr3)

# --- 工作表 4：更新說明 ---
ws4 = wb.create_sheet('更新說明')
ws4.column_dimensions['A'].width = 18
ws4.column_dimensions['B'].width = 96
cancel_summary = {}
for day, evs in cancelled:
    for e in evs:
        cancel_summary.setdefault(e, []).append(day.strftime('%-m/%-d'))
rows = [
    ('更新項目', '華嚴海會'),
    ('期間', '2027/12/11（%s）– 2027/12/22（%s）' % (wd(HY_START), wd(HY_END))),
    ('天數', '共 12 天'),
    ('原排程', '歲末大型華嚴法會：2027/12/25 – 2028/1/2'),
    ('異動內容 1', '刪除 2027/12/25–12/31 之「歲末大型華嚴法會」，改為 2027/12/11–12/22「華嚴海會」，共 12 天。'),
    ('異動內容 2', '華嚴海會為重大活動，2027/12/11–12/22 期間原排定之所有活動全數取消，該 12 天僅保留華嚴海會。'),
    ('取消活動明細', '\n'.join('%s：%s（共 %d 場／日）' % (e, '、'.join(ds), len(ds))
                           for e, ds in cancel_summary.items())),
    ('連帶調整', '兒童哲學班（秋季）14 堂→12 堂；企業專班：台中 8 堂→7 堂；好好讀楞嚴：板橋 7 堂→6 堂；'
               '一日禪（新加坡）8 堂→7 堂；楞嚴經（新加坡）8 堂→7 堂；'
               '秋季班原 9/13–12/24 共 15 週，12/11–12/22 停課，實際上課 9/13–12/10、12/23–12/24。'),
    ('其他說明', '除上述異動外，其餘月份與活動內容均與原 2027 行事曆相同。'),
    ('製作日期', datetime.date.today().strftime('%Y/%m/%d')),
]
r = 1
style(ws4.cell(r, 1, '2027 行事曆更新說明'), size=16, bold=True, border=False)
ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); ws4.row_dimensions[r].height = 26; r += 2
for k, v in rows:
    style(ws4.cell(r, 1, k), bold=True, bg=HEADBG)
    c = style(ws4.cell(r, 2, v))
    if k.startswith('取消活動'):
        c.font = Font(name=FONT, size=10, bold=True, color="9C0006")
        c.fill = PatternFill('solid', fgColor=HILITE)
    ws4.row_dimensions[r].height = max(34, 14 * (str(v).count(chr(10)) + 1) + 8)
    r += 1
ws4.page_margins.left = ws4.page_margins.right = 0.3
ws4.page_margins.top = ws4.page_margins.bottom = 0.35
ws4.page_setup.orientation = 'landscape'
ws4.sheet_properties.pageSetUpPr.fitToPage = True
ws4.page_setup.fitToWidth = 1
ws4.page_setup.fitToHeight = 0
ws4.print_area = 'A1:B%d' % (r - 1)

out = '2027行事曆_華嚴海會更新版.xlsx'
wb.save(out)
print('已輸出', out, '；重疊日數：', n_overlap)
