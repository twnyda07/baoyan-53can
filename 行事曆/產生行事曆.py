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
PREPBG = "FFF2CC"
BAGUANBG = "E2EFDA"
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

# ---------- 2. 本次更新：前置作業 12/13–12/14、華嚴海會 12/15–12/26 ----------
OLD_EVENT  = '歲末大型華嚴法會'
NEW_EVENT  = '華嚴海會'
PREP_EVENT = '華嚴海會前置作業'
HY_START,   HY_END   = datetime.date(2027, 12, 15), datetime.date(2027, 12, 26)
PREP_START, PREP_END = datetime.date(2027, 12, 13), datetime.date(2027, 12, 14)
BLOCK_START, BLOCK_END = PREP_START, HY_END          # 整段封鎖區間

def daterange(a, b):
    for i in range((b - a).days + 1):
        yield a + datetime.timedelta(days=i)

removed = []
for m in cells:
    for d, v in cells[m].items():
        if OLD_EVENT in v['ev']:
            v['ev'] = [e for e in v['ev'] if e != OLD_EVENT]
            removed.append('2027/%d/%d' % (m, d))

cancelled = []          # [(date, [被取消的活動])]
prep_days, hy_days = [], []
for day in daterange(BLOCK_START, BLOCK_END):
    cell = cells[day.month][day.day]
    if cell['ev']:
        cancelled.append((day, list(cell['ev'])))
    # 華嚴海會為重大活動：前置作業與法會期間，其餘活動全數取消
    if day <= PREP_END:
        cell['ev'] = [PREP_EVENT]; prep_days.append(day)
    else:
        cell['ev'] = [NEW_EVENT];  hy_days.append(day)

assert len(prep_days) == 2, prep_days
assert len(hy_days) == 12, hy_days
fmt = lambda ds: '、'.join(d.strftime('2027/%-m/%-d') for d in ds)
print('移除 %s：%s' % (OLD_EVENT, '、'.join(removed)))
print('新增 %s：%s（共 %d 天）' % (PREP_EVENT, fmt(prep_days), len(prep_days)))
print('新增 %s：%s（共 %d 天）' % (NEW_EVENT, fmt(hy_days), len(hy_days)))
print('封鎖期間取消之衝突活動：')
for day, evs in cancelled:
    print('   %s  %s' % (day.strftime('2027/%-m/%-d'), '、'.join(evs)))

WD = '日一二三四五六'
def wd(dt):
    return '星期' + WD[(dt.weekday() + 1) % 7]

# ---------- 2a. 浴佛節拆分：5/8 浴佛節法會、5/9–5/16 浴佛週 ----------
for day in daterange(datetime.date(2027, 5, 8), datetime.date(2027, 5, 16)):
    ev = cells[5][day.day]['ev']
    for i, e in enumerate(ev):
        if e == '浴佛節':
            ev[i] = '浴佛節法會' if day.day == 8 else '浴佛週'
print('浴佛節拆分：5/8 浴佛節法會、5/9–5/16 浴佛週')

# ---------- 2b. 八關齋戒暨共修法會：每月第三個星期日，逢法會之月份停辦 ----------
BAGUAN = '八關齋戒暨共修法會'
FAHUI_KEY = ('法會', '浴佛節', '華嚴海會')      # 判定「該月有法會」的關鍵字

def third_sunday(month):
    days = [d for d in range(1, calendar.monthrange(2027, month)[1] + 1)
            if datetime.date(2027, month, d).weekday() == 6]
    return days[2]

BAGUAN_SKIP = {(7, 18): '兒童夏令營 第2梯'}      # 個別停辦（非整月停辦）
baguan_held, baguan_skipped = [], []     # [(date, …)] / [(date, [原因])]
for m in range(1, 13):
    fahui = sorted({e for v in cells[m].values() for e in v['ev']
                    if any(k in e for k in FAHUI_KEY)})
    day = datetime.date(2027, m, third_sunday(m))
    if (m, day.day) in BAGUAN_SKIP:
        baguan_skipped.append((day, [BAGUAN_SKIP[(m, day.day)]]))
    elif fahui:
        baguan_skipped.append((day, fahui))
    else:
        cells[m][day.day]['ev'].append(BAGUAN)
        baguan_held.append(day)
print('%s：舉辦 %d 場 → %s' % (BAGUAN, len(baguan_held), fmt(baguan_held)))
for day, fahui in baguan_skipped:
    print('   停辦 %s（%d 月有 %s）' % (day.strftime('2027/%-m/%-d'), day.month, '、'.join(fahui)))

# ---------- 2c. 課程定義（sync=True 者月曆與課程場次完全同步） ----------
TEACHER = '見輝法師'
#  (課程, 規則, 地點, 月曆標籤, [(月,日)…], 見輝法師, 月曆同步)
COURSE_DEFS = [
    ('企業專班：台北',    '每月第一個星期六', '善首講堂', '企業專班：善首',
     [(3, 6), (4, 3), (5, 1), (6, 5), (8, 7), (10, 2), (11, 6), (12, 4)], True, True),
    ('企業專班：高雄',    '每月第一個星期日', '圓道禪寺（高雄）', '企業專班：高雄',
     [(3, 7), (4, 4), (5, 2), (6, 6), (8, 1), (10, 3), (11, 7), (12, 5)], True, True),
    ('企業專班：台中',    '每月第二個星期日', 'EPIC 禪藝實相人文空間（台中）', '企業專班：台中',
     [(3, 14), (4, 11), (5, 9), (6, 13), (8, 8), (10, 10), (11, 14), (12, 12)], True, True),
    ('好好讀楞嚴：板橋',  '每月第二週週六', '善覺講堂', '好好讀楞嚴（善覺）',
     [(3, 13), (4, 10), (5, 8), (6, 12), (10, 9), (11, 13), (12, 11)], True, True),
    ('一日禪（新加坡）',  '依各列日期', '新加坡', '一日禪（新加坡）',
     [(3, 20), (4, 17), (5, 15), (7, 17), (9, 18), (10, 16), (11, 20), (12, 18)], True, True),
    ('楞嚴經（新加坡）',  '依各列日期', '新加坡', '楞嚴經（新加坡）',
     [(3, 21), (4, 18), (5, 16), (7, 18), (9, 19), (10, 17), (11, 21), (12, 19)], True, True),
    # 以下非見輝法師授課，但月曆同樣與課程場次同步
    ('英文一日禪（新加坡）', '每月第一個星期六', '新加坡', '英文一日禪（新加坡）',
     [(3, 13), (4, 3), (5, 1), (6, 5), (8, 7), (10, 2), (11, 6), (12, 4)], False, True),
    ('英文楞嚴經（新加坡）', '每月第一個星期日', '新加坡', '英文楞嚴經（新加坡）',
     [(3, 14), (4, 4), (5, 2), (6, 6), (8, 1), (10, 3), (11, 7), (12, 5)], False, True),
    # 秋季兒童哲學班：2027/9/26–12/19 週日班（遇華嚴海會自動停課）
    ('兒童哲學班（秋季）', '每週日', '各分院', '兒童哲學班（秋季）',
     [(d.month, d.day) for d in daterange(datetime.date(2027, 9, 26), datetime.date(2027, 12, 19))
      if d.weekday() == 6], False, True),
]

# ---------- 2d. 個別調整（本次新增，依排程決議逐筆記錄） ----------
#  (課程, 動作, 原日期, 新日期, 原因)
ADJUSTMENTS = [
    ('好好讀楞嚴：板橋',   'cancel', (3, 13), None,    '3/13 三案重疊，以新加坡禪七為主'),
    ('英文一日禪（新加坡）', 'cancel', (3, 13), None,    '3/13 三案重疊，以新加坡禪七為主'),
    ('企業專班：台中',     'cancel', (3, 14), None,    '3 月適逢新加坡禪七（3/13–3/21），法師時間重疊'),
    ('一日禪（新加坡）',   'cancel', (3, 20), None,    '合併至新加坡禪七'),
    ('英文楞嚴經（新加坡）', 'cancel', (3, 14), None,    '合併至新加坡禪七'),
    ('楞嚴經（新加坡）',   'cancel', (3, 21), None,    '合併至新加坡禪七'),
    ('英文一日禪（新加坡）', 'cancel', (4, 3),  None,    '清明法會放蒙山（台中）'),
    ('企業專班：台中',     'move',   (8, 8),  (8, 15), '8/8 彰化供佛齋僧大會'),
    ('好好讀楞嚴：板橋',   'add',    None,    (8, 14), '善覺講堂加排'),
    ('好好讀楞嚴：板橋',   'add',    None,    (7, 10), '7 月場次補排'),  # 7/3、7/4 適逢夏安居禪七第二梯次，不補排
    ('企業專班：台中',     'add',    None,    (7, 11), '7 月場次補排'),
    ('企業專班：台北',     'add',    None,    (9, 4),  '9 月場次補排'),
    ('企業專班：高雄',     'add',    None,    (9, 5),  '9 月場次補排'),
    ('好好讀楞嚴：板橋',   'add',    None,    (9, 11), '9 月場次補排'),
    ('企業專班：台中',     'add',    None,    (9, 12), '9 月場次補排'),
]
TIME_OVERRIDE = {('好好讀楞嚴：板橋', (5, 8)): '14:00'}   # 單場上課時間調整

_dates = {c[0]: list(c[4]) for c in COURSE_DEFS}
adj_log = []
for name, act, old, new_, why in ADJUSTMENTS:
    ds = _dates[name]
    if act in ('cancel', 'move'):
        assert old in ds, (name, act, old)
        ds.remove(old)
    if act in ('move', 'add'):
        assert new_ not in ds, (name, act, new_)
        ds.append(new_)
    adj_log.append((name, act, old, new_, why))
for ds in _dates.values():
    ds.sort()

# ---------- 2e. 套用至月曆，並建立授課資料 ----------
backfilled, dropped = [], []
course_rows = []     # (課程, 規則, 地點, [保留日期], [封鎖取消日期], 見輝法師)
for name, rule, place, cal_label, _base, is_teacher, sync in COURSE_DEFS:
    keep, blocked = [], []
    for m, d in _dates[name]:
        dt = datetime.date(2027, m, d)
        (blocked if BLOCK_START <= dt <= BLOCK_END else keep).append(dt)
    want = {(d.month, d.day) for d in keep}
    for m in range(1, 13):                       # 不在清單上的場次一律自月曆移除
        for d, v in cells[m].items():
            hit = [e for e in v['ev'] if e == cal_label or e.startswith(cal_label + ' ')]
            if hit and (m, d) not in want:
                for e in hit:
                    v['ev'].remove(e)
                dropped.append((datetime.date(2027, m, d), cal_label))
    if sync:                                     # 清單上缺的場次回補月曆
        for dt in keep:
            t = TIME_OVERRIDE.get((name, (dt.month, dt.day)))
            label = cal_label + (' ' + t if t else '')
            ev = cells[dt.month][dt.day]['ev']
            if label not in ev:
                for e in list(ev):
                    if e == cal_label or e.startswith(cal_label + ' '):
                        ev.remove(e)
                ev.append(label)
                backfilled.append((dt, label))
    course_rows.append((name, rule, place, keep, blocked, is_teacher))
teacher_rows = [r[:5] for r in course_rows if r[5]]

for name, act, old, new_, why in adj_log:
    d = lambda x: '%d/%d' % x if x else '—'
    print('調整：%-18s %-6s %-6s → %-6s  %s' % (name, act, d(old), d(new_), why))
for dt, lab in dropped:
    print('月曆移除：%s %s' % (dt.strftime('2027/%-m/%-d'), lab))
for dt, lab in backfilled:
    print('月曆新增：%s %s' % (dt.strftime('2027/%-m/%-d'), lab))
print('%s授課合計 %d 堂' % (TEACHER, sum(len(r[3]) for r in teacher_rows)))

# ---------- 3. 活動總表 ----------
COURSES = [
    ('兒童哲學班（春季）', '共 16 堂', '2027/3/7、2027/3/14、2027/3/21、2027/3/28、2027/4/4、2027/4/11、2027/4/18、2027/4/25、2027/5/2、2027/5/9、2027/5/16、2027/5/23、2027/5/30、2027/6/6、2027/6/13、2027/6/20'),
    ('兒童哲學班（秋季）', '共 12 堂', '2027/9/26、2027/10/3、2027/10/10、2027/10/17、2027/10/24、2027/10/31、2027/11/7、2027/11/14、2027/11/21、2027/11/28、2027/12/5、2027/12/12（原 14 堂，12/19、12/26 因華嚴海會取消）'),
    ('春季班', '共 17 週', '2027/3/2–2027/6/18（週一至週五晚上，首週自週二起）'),
    ('暑期班', '共 5 週',  '2027/7/12–2027/8/13（週一至週五晚上）'),
    ('秋季班', '原共 15 週', '2027/9/13–2027/12/24（週一至週五晚上）；12/13 起因華嚴海會停課，實際上課 2027/9/13–2027/12/10，共 13 週'),
    ('企業專班：台北', '共 8 堂', '2027/3/6、2027/4/3、2027/5/1、2027/6/5、2027/8/7、2027/10/2、2027/11/6、2027/12/4'),
    ('企業專班：高雄', '共 8 堂', '2027/3/7、2027/4/4、2027/5/2、2027/6/6、2027/8/1、2027/10/3、2027/11/7、2027/12/5'),
    ('企業專班：台中', '共 8 堂', '2027/3/14、2027/4/11、2027/5/9、2027/6/13、2027/8/8、2027/10/10、2027/11/14、2027/12/12'),
    ('好好讀楞嚴：板橋', '共 7 堂', '2027/3/13、2027/4/10、2027/5/8、2027/6/12、2027/10/9、2027/11/13、2027/12/11'),
    ('英文一日禪（新加坡）', '共 8 堂', '2027/3/13、2027/4/3、2027/5/1、2027/6/5、2027/8/7、2027/10/2、2027/11/6、2027/12/4'),
    ('英文楞嚴經（新加坡）', '共 8 堂', '2027/3/14、2027/4/4、2027/5/2、2027/6/6、2027/8/1、2027/10/3、2027/11/7、2027/12/5'),
    ('一日禪（新加坡）', '共 7 堂', '2027/3/20、2027/4/17、2027/5/15、2027/7/17、2027/9/18、2027/10/16、2027/11/20（原 8 堂，12/18 因華嚴海會取消）'),
    ('楞嚴經（新加坡）', '共 7 堂', '2027/3/21、2027/4/18、2027/5/16、2027/7/18、2027/9/19、2027/10/17、2027/11/21（原 8 堂，12/19 因華嚴海會取消）'),
]

# 見輝法師 6 門課改由 teacher_rows 產生（取消場次自動扣除），並加入八關齋戒
_tc = {r[0]: r[:5] for r in course_rows}
_adj_by_course = {}
for name, act, o, n_, why in adj_log:
    if act == 'cancel':
        txt = '%d/%d 取消（%s）' % (o[0], o[1], why)
    elif act == 'move':
        txt = '%d/%d 改至 %d/%d（%s）' % (o[0], o[1], n_[0], n_[1], why)
    else:
        txt = '%d/%d 加排（%s）' % (n_[0], n_[1], why)
    _adj_by_course.setdefault(name, []).append(txt)
def _notes(name):
    out = list(_adj_by_course.get(name, []))
    blocked = _tc[name][4] if name in _tc else []
    out += ['%s 取消（華嚴海會）' % d.strftime('%-m/%-d') for d in blocked]
    for (cn, md), t in TIME_OVERRIDE.items():
        if cn == name:
            out.append('%d/%d 改為 %s 上課' % (md[0], md[1], t))
    return out
def _course_row(name, n, dates):
    if name not in _tc:
        return (name, n, dates)
    _, _, _, keep, _blocked = _tc[name]
    txt = '、'.join(d.strftime('2027/%-m/%-d') for d in keep)
    notes = _notes(name)
    if notes:
        txt += '　※ ' + '；'.join(notes)
    return (name, '共 %d 堂' % len(keep), txt)
COURSES = [_course_row(*c) for c in COURSES]
_bg_skip = '、'.join('%d/%d（%s）' % (d.month, d.day, r[0]) for d, r in baguan_skipped)
COURSES.append((BAGUAN, '共 %d 場' % len(baguan_held),
                '%s　※ 每月第三個星期日；停辦：%s' % (fmt(baguan_held), _bg_skip)))

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
    ('2027/9/13–2027/12/10\n（原排至 12/24）', '星期一～星期五（晚上）', '秋季班', '一般課程', '待確認', '各分院',
     '原共 15 週；12/13 起因華嚴海會前置作業及法會停課，實際上課 9/13–12/10，共 13 週'),
    ('2027/9/28', '星期二', '教師禪修營', '禪七／禪修', '待確認', '待確認', '新增活動'),
    ('2027/12/13–2027/12/14', '%s～%s' % (wd(PREP_START), wd(PREP_END)), '華嚴海會前置作業', '法會／大型活動', '待確認', '各分院',
     '共 2 天；華嚴海會前置準備，不排其他活動'),
    ('2027/12/15–2027/12/26', '%s～%s' % (wd(HY_START), wd(HY_END)), '華嚴海會', '法會／大型活動', '待確認', '各分院',
     '共 12 天（本次更新）；原「歲末大型華嚴法會」2027/12/25–2028/1/2 調整為本案；本期間為重大活動，原有之其他活動均已取消'),
    ('%s' % fmt(baguan_held), '每月第三個星期日', BAGUAN, '法會／大型活動', '待確認', '各分院',
     '共 %d 場；配合全年度開課行事曆，遇華嚴海會、清明法會等法會之月份停辦。停辦：%s'
     % (len(baguan_held), _bg_skip)),
    ('2027/3/7 起每週日', '每週日', '兒童哲學班（春季）', '兒童教育', '待確認', '各分院', '共 16 堂；日期詳見上方課程場次表'),
    ('2027/9/26 起每週日', '每週日', '兒童哲學班（秋季）', '兒童教育', '待確認', '各分院', '共 12 堂（原 14 堂；12/19、12/26 因華嚴海會取消）；日期詳見上方課程場次表'),
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
    ('2027/1/10、2027/2/14、2027/3/14、2027/4/11、2027/5/9、2027/6/13、2027/7/11、2027/8/8、2027/9/12、2027/10/10、2027/11/14、2027/12/12',
     '依各列日期', '企業專班：台中', '企業專班', '待確認', 'EPIC 禪藝實相人文空間（台中）', '每月第二個星期日'),
    ('2027/1/9、2027/2/13、2027/3/13、2027/4/10、2027/5/8、2027/6/12、2027/7/10、2027/8/14、2027/9/11、2027/10/9、2027/11/13、2027/12/11',
     '依各列日期', '好好讀楞嚴（善覺）', '一般課程', '待確認', '善覺講堂', '每月第二週週六'),
]

# 企業專班與好好讀楞嚴的「活動明細」改採「一、課程場次」版本（8 堂／7 堂，3 月起）
_EV2TC = {'企業專班：台北': '企業專班：台北', '企業專班：高雄': '企業專班：高雄',
          '企業專班：台中': '企業專班：台中', '好好讀楞嚴（善覺）': '好好讀楞嚴：板橋',
          '兒童哲學班（秋季）': '兒童哲學班（秋季）'}
def _event_row(row):
    key = _EV2TC.get(row[2])
    if key is None:
        return row
    _, rule, place, keep, _b = _tc[key]
    note = '%s；共 %d 堂' % (rule, len(keep))
    if key == '兒童哲學班（秋季）':
        note = '每週日；2027/9/26–2027/12/19；共 %d 堂' % len(keep)
    extra = _notes(key)
    if extra:
        note += '；' + '；'.join(extra)
    return ('、'.join(d.strftime('2027/%-m/%-d') for d in keep),
            '依各列日期', row[2], row[3], row[4], place, note)
EVENTS = [_event_row(r) for r in EVENTS]


CAL_COL_WIDTH = 16          # 月曆欄寬（Excel 字元單位）
def vis_lines(text):
    """估算一段文字在月曆格內換行後佔用的行數。"""
    w = sum(2 if ord(ch) > 0x2E7F else 1 for ch in text)
    return max(1, -(-w // CAL_COL_WIDTH))

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
    ws.column_dimensions[get_column_letter(i)].width = CAL_COL_WIDTH
r = 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
style(ws.cell(r, 1, '2027 年全年行事曆（更新版）'), size=18, bold=True, h='left', border=False)
ws.row_dimensions[r].height = 30; r += 1
style(ws.cell(r, 1, '每日顯示農曆日期；本版更新：華嚴海會 2027/12/15–2027/12/26（共 12 天），前置作業 2027/12/13–2027/12/14（共 2 天）；該期間原排定之其他活動均已取消。'),
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
            maxlines = max(maxlines, sum(vis_lines(t) for t in lines))
            c.value = '\n'.join(lines)
            dt = datetime.date(2027, m, d)
            hl = HY_START <= dt <= HY_END
            pp = PREP_START <= dt <= PREP_END
            style(c, bg=HILITE if hl else (PREPBG if pp else None))
            if hl or pp:
                c.font = Font(name=FONT, size=10, bold=True,
                              color="9C0006" if hl else "7F6000")
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
for i in range(2, 12, 2):            # 每 2 個月換頁
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
    hl = row[2] in (NEW_EVENT, PREP_EVENT)
    bg_row = row[2] == BAGUAN
    for i, val in enumerate(row):
        c = style(ws2.cell(r, i + 1, val), bg=HILITE if hl else (BAGUANBG if bg_row else None))
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

# --- 工作表 3：見輝法師授課 ---
ws5 = wb.create_sheet('見輝法師授課')
for col, w in zip('ABCDE', (26, 20, 32, 12, 54)):
    ws5.column_dimensions[col].width = w
r = 1
style(ws5.cell(r, 1, '%s 2027 年授課一覽' % TEACHER), size=16, bold=True, border=False)
ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
ws5.row_dimensions[r].height = 26; r += 1
style(ws5.cell(r, 1, '合計 %d 堂；日期落在華嚴海會封鎖區間（12/13–12/26）者已取消。'
                     % sum(len(x[3]) for x in teacher_rows)),
      size=9, color="7F7F7F", border=False)
ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
r += 2

style(ws5.cell(r, 1, '一、課程總覽'), size=12, bold=True, border=False); r += 1
for i, h in enumerate(['課程', '開課規則', '地點', '堂數', '日期']):
    style(ws5.cell(r, i + 1, h), bold=True, color="FFFFFF", bg=NAVY, h='center', v='center')
ws5.row_dimensions[r].height = 22; r += 1
for name, rule, place, keep, drop in teacher_rows:
    txt = '、'.join(d.strftime('2027/%-m/%-d') for d in keep)
    if drop:
        txt += '　（取消：%s）' % '、'.join(d.strftime('%-m/%-d') for d in drop)
    for i, val in enumerate([name, rule, place, '%d 堂' % len(keep), txt]):
        style(ws5.cell(r, i + 1, val), bg=CREAM, h='center' if i == 3 else 'left')
    ws5.row_dimensions[r].height = 32; r += 1
r += 1

style(ws5.cell(r, 1, '二、授課行程（依日期排序）'), size=12, bold=True, border=False); r += 1
hdr5 = r
for i, h in enumerate(['日期', '星期', '課程', '地點', '備註']):
    style(ws5.cell(r, i + 1, h), bold=True, color="FFFFFF", bg=NAVY, h='center', v='center')
ws5.row_dimensions[r].height = 22; r += 1
sched = sorted(((d, name, place) for name, rule, place, keep, _ in teacher_rows for d in keep),
               key=lambda x: x[0])
prev_month = None
for d, name, place in sched:
    if prev_month is not None and d.month != prev_month:
        ws5.row_dimensions[r].height = 6; r += 1        # 月份之間留空行
    prev_month = d.month
    same_day = [n for dd, n, _ in sched if dd == d and n != name]
    note = '同日另有：' + '、'.join(same_day) if same_day else ''
    for i, val in enumerate([d.strftime('2027/%-m/%-d'), wd(d), name, place, note]):
        style(ws5.cell(r, i + 1, val), h='center' if i in (0, 1) else 'left')
    ws5.row_dimensions[r].height = 20; r += 1
ws5.freeze_panes = 'A%d' % (hdr5 + 1)
ws5.page_margins.left = ws5.page_margins.right = 0.3
ws5.page_margins.top = ws5.page_margins.bottom = 0.35
ws5.page_setup.orientation = 'landscape'
ws5.page_setup.paperSize = ws5.PAPERSIZE_A4
ws5.sheet_properties.pageSetUpPr.fitToPage = True
ws5.page_setup.fitToWidth = 1
ws5.page_setup.fitToHeight = 0
ws5.print_area = 'A1:E%d' % (r - 1)
ws5.print_title_rows = '%d:%d' % (hdr5, hdr5)

# --- 工作表 3：活動重疊檢查 ---
ws3 = wb.create_sheet('活動重疊檢查')
for col, w in zip('ABCD', (16, 10, 12, 76)):
    ws3.column_dimensions[col].width = w
r = 1
style(ws3.cell(r, 1, '2027 年活動重疊檢查'), size=16, bold=True, border=False)
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4); ws3.row_dimensions[r].height = 26; r += 1
style(ws3.cell(r, 1, '列出同一日安排 2 項（含）以上活動之日期。華嚴海會前置作業及法會期間（12/13–12/26）之衝突活動已全數取消，故該期間不再出現於本表。'),
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
        hl = BLOCK_START <= dt <= BLOCK_END
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
    ('更新項目', '華嚴海會（含前置作業）'),
    ('前置作業', '2027/12/13（%s）– 2027/12/14（%s），共 2 天' % (wd(PREP_START), wd(PREP_END))),
    ('法會期間', '2027/12/15（%s）– 2027/12/26（%s），共 12 天' % (wd(HY_START), wd(HY_END))),
    ('封鎖區間', '2027/12/13 – 2027/12/26，合計 14 天'),
    ('原排程', '歲末大型華嚴法會：2027/12/25 – 2028/1/2'),
    ('異動內容 1', '刪除原「歲末大型華嚴法會」2027/12/25–2028/1/2。'),
    ('異動內容 2', '新增「華嚴海會前置作業」12/13–12/14（2 天）與「華嚴海會」12/15–12/26（12 天）。'),
    ('異動內容 3', '華嚴海會為重大活動，12/13–12/26 期間原排定之所有活動全數取消，該 14 天僅保留前置作業與華嚴海會。'),
    ('取消活動明細', '\n'.join('%s：%s（共 %d 場／日）' % (e, '、'.join(ds), len(ds))
                           for e, ds in cancel_summary.items())),
    ('連帶調整', '秋季班原 9/13–12/24 共 15 週，12/13 起停課，實際上課 9/13–12/10 共 13 週；'
               '兒童哲學班（秋季）14 堂→12 堂（取消 12/19、12/26）；'
               '一日禪（新加坡）8 堂→7 堂（取消 12/18）；楞嚴經（新加坡）8 堂→7 堂（取消 12/19）。'),
    ('未受影響', '好好讀楞嚴（善覺）12/11、企業專班：台中 12/12、兒童哲學班（秋季）12/12 均在封鎖區間之前，維持原排程。'),
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

r += 1
style(ws4.cell(r, 1, '本次個別調整明細'), size=12, bold=True, border=False)
ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
ws4.row_dimensions[r].height = 24; r += 1
_ACT = {'cancel': '取消', 'move': '改期', 'add': '加排'}
for name, act, o, n_, why in adj_log:
    when = ('%d/%d → %d/%d' % (o + n_)) if act == 'move' else ('%d/%d' % (o or n_))
    style(ws4.cell(r, 1, '%s　%s　%s' % (_ACT[act], name, when)), bold=True, bg=HEADBG)
    style(ws4.cell(r, 2, why))
    ws4.row_dimensions[r].height = 20; r += 1
for (cn, md), t in TIME_OVERRIDE.items():
    style(ws4.cell(r, 1, '時間　%s　%d/%d' % (cn, md[0], md[1])), bold=True, bg=HEADBG)
    style(ws4.cell(r, 2, '改為 %s 上課' % t)); ws4.row_dimensions[r].height = 20; r += 1
style(ws4.cell(r, 1, '取消　%s　7/18' % BAGUAN), bold=True, bg=HEADBG)
style(ws4.cell(r, 2, '兒童夏令營 第2梯')); ws4.row_dimensions[r].height = 20; r += 1
style(ws4.cell(r, 1, '拆分　浴佛節'), bold=True, bg=HEADBG)
style(ws4.cell(r, 2, '5/8 浴佛節法會；5/9–5/16 浴佛週')); ws4.row_dimensions[r].height = 20; r += 1
ws4.print_area = 'A1:B%d' % (r - 1)

out = '2027行事曆_華嚴海會更新版.xlsx'
wb.save(out)
print('已輸出', out, '；重疊日數：', n_overlap)
