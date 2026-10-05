import re, collections, json, math, calendar
src = open('b1.html', encoding='utf-8').read()
W=[]
for m in re.finditer(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', src):
    W.append(dict(x0=float(m.group(1)),y0=float(m.group(2)),x1=float(m.group(3)),y1=float(m.group(4)),t=m.group(5)))
W=[w for w in W if w['y0']<4840]
X0=74.5; CW=114.0
for w in W: w['col']=min(6,max(0,int(math.floor((w['x0']-X0+2)/CW))))
bycol=collections.defaultdict(list)
for w in W: bycol[w['col']].append(w)
cols={}
for c,ws in bycol.items():
    ws.sort(key=lambda w:(w['y0'],w['x0']))
    lines=[];cur=[]
    for w in ws:
        if cur and abs(w['y0']-cur[0]['y0'])>3: lines.append(cur);cur=[]
        cur.append(w)
    if cur: lines.append(cur)
    out=[]
    for ln in lines:
        s='';prev=None
        for w in ln:
            if prev is not None:
                g=w['x0']-prev
                if g>6: s+='□'*int(round(g/10.0))
            s+=w['t'];prev=w['x1']
        out.append((ln[0]['y0'],s))
    cols[c]=out
# month label anchors in col 0
mlab=[(y,int(s)) for y,s in cols[0] if re.fullmatch(r'\d{1,2}',s) and 1<=int(s)<=12]
mlab.sort()
assert [m for _,m in mlab]==list(range(1,13)), mlab
bounds=[y for y,_ in mlab]+[1e9]
def month_of(y):
    for i in range(12):
        if bounds[i]<=y<bounds[i+1]: return i+1
    return None
SKIP={'三','四','五','六','。','「活動重疊檢查」'}|{str(i) for i in range(1,13)}
anchors=[]
for c in range(7):
    cur=None
    for y,s in cols[c]:
        m=re.match(r'^(\d+)（農(\d+)/(\d+)）$', s)
        if m:
            cur={'y':y,'col':c,'day':int(m.group(1)),'lunar':m.group(2)+'/'+m.group(3),'ev':[]}
            anchors.append(cur); continue
        if s in SKIP or s.startswith('年全年') or s.startswith('每□顯'):
            cur=None; continue
        if cur: cur['ev'].append(s)
res=collections.defaultdict(dict)
for a in anchors:
    mo=month_of(a['y'])
    assert a['day'] not in res[mo], (mo,a['day'])
    res[mo][a['day']]=dict(lunar=a['lunar'], ev=a['ev'], col=a['col'])
bad=0
for m in range(1,13):
    nd=calendar.monthrange(2027,m)[1]
    miss=[d for d in range(1,nd+1) if d not in res[m]]
    if miss: print('month',m,'missing',miss); bad+=1
    for d,v in res[m].items():
        if (calendar.weekday(2027,m,d)+1)%7 != v['col']: print('COLMISS',m,d,v['col']); bad+=1
print('problems',bad)
json.dump({str(m):{str(d):{'lunar':v['lunar'],'ev':v['ev']} for d,v in sorted(res[m].items())} for m in sorted(res)}, open('cells.json','w'), ensure_ascii=False, indent=1)
for m in [12]:
    for d in sorted(res[m]): print(m,d,res[m][d]['lunar'],res[m][d]['ev'])
