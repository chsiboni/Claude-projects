"""רשימת דיוור להשקה — שורה לאיש קשר (טלפון + מייל), גיליון הנהלה וגיליון כל המשתמשים.
היקף: לקוחות משלמים מדוח ספטמבר (out/clients_status_2026-09.json, בלי לא-פעילים).
מקורות: dump בורדים (חברות וסניפים, סניפים, אנשי קשר), vendor.json (טופס המילוי והסטטוסים של החברה החיצונית), בורד גבייה.
שורות מהאימות הטלפוני קודמות לשאר ומסומנות בכחול. הרצה: python3 src/mailing_list_2026.py (נתיבים בראש הקובץ).
"""
import gzip,json,re,collections as C,sys
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
V='/tmp/claude-0/-home-user-Claude-projects/9bedc6a6-f60e-5cce-a499-d0c73e5871ea/scratchpad/vendor/'
B=json.load(gzip.open(V+'dl/boards2.json.gz','rt',encoding='utf-8')); CON=json.load(gzip.open(V+'dl/contacts.json.gz','rt',encoding='utf-8'))['items']
VEN=json.load(open(V+'vendor.json',encoding='utf-8'))
SEP=json.load(open('/home/user/Claude-projects/fizikal-crm/out/clients_status_2026-09.json',encoding='utf-8'))
g=json.load(open('/root/.claude/projects/-home-user-Claude-projects/9bedc6a6-f60e-5cce-a499-d0c73e5871ea/tool-results/mcp-monday_com-all_api_read-1793692436635.txt'))
GV={}
for it in g['boards'][0]['items_page']['items']:
    cv={c['id']:(c['text'] or '').strip() for c in it['column_values']}
    if cv['text_mm69gg7g']: GV[cv['text_mm69gg7g']]={'phone':cv['phone_mm699f4'],'email':cv['email_mm6k2wy6'],'name':it['name']}
norm=lambda s:re.sub(r'\s+',' ',(s or '')).strip()
def ph(p):
    p=re.sub(r'\D','',p or '')
    if p.startswith('972'): p='0'+p[3:]
    if len(p)==9 and p[0] in '5': p='0'+p
    if len(p)==8 and p[0] in '2345689': p='0'+p
    return p if re.fullmatch(r'0\d{8,9}',p) else ''
mob=lambda p: re.fullmatch(r'05\d{8}',p) is not None
def em(e):
    e=(e or '').strip().lower()
    return e if re.fullmatch(r'[^@\s,;]+@[^@\s,;]+\.[a-z]{2,}',e) else ''
# scope: September paying clients
scope={c['sap']:c for c in SEP if c['status']!='לא פעיל'}
print('scope clients',len(scope),dict(C.Counter(c['status'] for c in scope.values())))
# company codes per SAP
BR=C.defaultdict(list); SAPB=C.defaultdict(set)
for b in B['branches']:
    cc=b['text_mm6v5fq7']; BR[cc].append(b)
    if b['text_mm6vwm4f']: SAPB[b['text_mm6vwm4f']].add(cc)
codes={}
for co in B['companies']:
    sap=co['text_mm6ccg01']; cs={s['text_mm76nbjb'] for s in co.get('subs',[]) if s.get('text_mm76nbjb')} or SAPB.get(sap,set())
    codes[sap]=cs
for sap in scope:
    if sap not in codes: codes[sap]=SAPB.get(sap,set())
cc2sap={}
for sap,cs in codes.items():
    if sap in scope:
        for cc in cs: cc2sap.setdefault(cc,sap)
RANK={'ניהול חברה':3,'הרשאת על':3,'מנהל':2,'מנהל משמרת':1}
def is_mg(c): r=c['dropdown_mm6v5453']; gr=c['dropdown_mm6v5svk']; return r in('ניהול חברה','הרשאת על') or (r=='מנהל' and gr in('הנהלה','הכל','מנהל'))
def role_txt(c):
    r=c['dropdown_mm6v5453']; gr=c['dropdown_mm6v5svk']
    return {'ניהול חברה':'ניהול חברה','הרשאת על':'הרשאת על'}.get(r, 'מנהל (הנהלה)' if (r=='מנהל' and gr in('הנהלה','הכל','מנהל')) else (r or 'משתמש'))
bname={}
for b in B['branches']: bname[(b['text_mm6v5fq7'],b['text_mm6vxmyx'])]=b['name']
def branch_of(c):
    cc=c['text_mm6v6myv']; bns=[x.strip() for x in re.split(r'[,،]',c['dropdown_mm6vwz6r']) if x.strip()]
    names=[bname.get((cc,bn)) for bn in bns]; names=[n for n in names if n]
    if not names: names=[x.strip() for x in re.split(r'[,،]',c['text_mm6vw0jh']) if x.strip()]
    return ', '.join(names[:3])+(' ועוד' if len(names)>3 else '')
HEAD=['שם','מספר טלפון','סוג','מייל','רמת לקוח','חברה','מספר לקוח','סניף','דרג','מקור','אומת בתאריך','סטטוס שיחה (חברה חיצונית)']
call_status={d.get('מספר לקוח',''):(d.get('סטטוס שיחה',''),d.get('הערות','')) for d in VEN['calls'] if d.get('מספר לקוח')}
mg_rows=[]; all_rows=[]
seen_mg=C.defaultdict(set); seen_all=C.defaultdict(set)
def add(target,seen,sap,name,phone,email,branch,rank,src,verified=''):
    c=scope[sap]; key=(phone or email or name.lower())
    if key in seen[sap]: return False
    seen[sap].add(key)
    cs=call_status.get(sap,('',''))
    target.append([name,phone,'נייד' if mob(phone) else ('נייח' if phone else ''),email,c['tier'],c['name'],sap,branch,rank,src,verified,(cs[0]+(' · '+cs[1] if cs[1] else '')).strip(' ·')])
    return True
# 1) vendor-verified contacts first (freshest)
vcount=0
for f in VEN['form']:
    sap=f.get('מספר לקוח','')
    if sap not in scope: continue
    nm=norm(f.get('שם מלא')); p=ph(f.get('טלפון')) or ph(f.get('טלפון עסקי')); e=em(f.get('מייל')) or em(f.get('מייל עסקי'))
    if not nm and not (p or e): continue
    rank=f.get('תפקיד') or 'עסק'
    nm=nm or scope[sap]['name']
    ver=f.get('אומת בתאריך') or f.get('') or '2026-10'
    if add(mg_rows,seen_mg,sap,nm,p,e,norm(f.get('שם סניף')),rank,'אימות טלפוני (חברה חיצונית)',ver): vcount+=1
    add(all_rows,seen_all,sap,nm,p,e,norm(f.get('שם סניף')),rank,'אימות טלפוני (חברה חיצונית)',ver)
print('vendor-verified rows',vcount)
# 2) board contacts
for c in CON:
    sap=cc2sap.get(c['text_mm6v6myv'])
    if not sap: continue
    p=ph(c['text_mm6vqqpn']); e=em(c['text_mm6vwyzq'])
    if not p and not e: continue
    if is_mg(c): add(mg_rows,seen_mg,sap,norm(c['name']),p,e,branch_of(c),role_txt(c),'מערכת פיזיקל (ייצוא 03/09)')
    add(all_rows,seen_all,sap,norm(c['name']),p,e,branch_of(c),role_txt(c),'מערכת פיזיקל (ייצוא 03/09)')
# 3) clients with no management row: branch phone / collections
nomg=[s for s in scope if not seen_mg[s]]
for sap in nomg:
    done=False
    for cc in codes.get(sap,()):
        for b in BR.get(cc,[]):
            p=ph(b['text_mm6vnk3m']); e=em(b['dropdown_mm6vfdrh'])
            if (p or e) and b['color_mm6vn3xc']=='פעיל': done|=add(mg_rows,seen_mg,sap,b['name'],p,e,b['name'],'טלפון הסניף','בורד סניפים')
    if not done and sap in GV:
        p=ph(GV[sap]['phone']); e=em(GV[sap]['email'])
        if p or e: done|=add(mg_rows,seen_mg,sap,GV[sap]['name'],p,e,'','כספים (גבייה)','בורד גבייה')
    if not done: add(mg_rows,seen_mg,sap,scope[sap]['name'],'','','', 'אין פרטי קשר','')
TIERS=['VIP','גדול','בינוני','קטן','מוב']
order={t:i for i,t in enumerate(TIERS)}
mg_rows.sort(key=lambda r:(order.get(r[4],9),r[5],r[9]!='אימות טלפוני (חברה חיצונית)',r[0]))
all_rows.sort(key=lambda r:(order.get(r[4],9),r[5],r[0]))
# stats
def stats(rows,label):
    cl=C.defaultdict(lambda:C.Counter())
    for r in rows:
        if r[8]=='אין פרטי קשר': continue
        cl[r[4]]['rows']+=1; cl[r[4]]['mob']+=mob(r[1]); cl[r[4]]['em']+=bool(r[3]); cl[r[4]]['both']+=(mob(r[1]) and bool(r[3]))
    tot=C.Counter()
    for t in cl: tot.update(cl[t])
    print(label,dict(tot),{t:dict(v) for t,v in cl.items()})
    return cl,tot
s_mg=stats(mg_rows,'MGMT'); s_all=stats(all_rows,'ALL')
cov=C.Counter(); covm=C.Counter()
for sap,c in scope.items():
    cov[c['tier']]+=1
    if any(mob(r[1]) for r in mg_rows if r[6]==sap): covm[c['tier']]+=1
print('clients with a management mobile by tier',{t:(covm[t],cov[t]) for t in TIERS},'total',sum(covm.values()),'/',len(scope))
# write
F='Arial'; wb=openpyxl.Workbook(); wb.remove(wb.active)
thin=Side(style='thin',color='C9D1E0'); BD=Border(top=thin,bottom=thin,left=thin,right=thin)
W=[24,14,7,30,10,30,11,22,16,26,12,34]
for title,rows in (('הנהלה',mg_rows),('כל המשתמשים',all_rows)):
    ws=wb.create_sheet(title); ws.sheet_view.rightToLeft=True
    for i,(h,w) in enumerate(zip(HEAD,W),1):
        c=ws.cell(row=1,column=i,value=h); c.font=Font(name=F,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F2A44'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BD
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    for r,row in enumerate(rows,2):
        for i,v in enumerate(row,1):
            c=ws.cell(row=r,column=i,value=v or None); c.font=Font(name=F,size=10); c.border=BD
            if i in(2,7): c.number_format='@'
            if i==10 and v.startswith('אימות'): c.fill=PatternFill('solid',fgColor='E3F2FD')
    ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:{openpyxl.utils.get_column_letter(len(HEAD))}{len(rows)+1}'
ws=wb.create_sheet('סיכום',0); ws.sheet_view.rightToLeft=True
ws['A1']='רשימת דיוור להשקה — Fizikal, 03/11/2026'; ws['A1'].font=Font(name=F,bold=True,size=13)
ws['A2']='היקף: לקוחות משלמים לפי דוח ההכנסות ספטמבר 2026 (פעילים, חדשים ובסיכון; בלי לא-פעילים). מקורות: משתמשי מערכת פיזיקל (ייצוא 03/09), אימות טלפוני של החברה החיצונית (אוקטובר), בורד סניפים ובורד גבייה. שורות מהאימות הטלפוני מסומנות בכחול וקודמות לשאר.'
ws['A2'].alignment=Alignment(wrap_text=True); ws.merge_cells('A2:H2'); ws.row_dimensions[2].height=45
hdr=['רמה','לקוחות','לקוחות עם נייד של מנהל','שורות הנהלה','הנהלה עם נייד','הנהלה עם מייל','שורות כל המשתמשים','משתמשים עם נייד','משתמשים עם מייל']
for i,h in enumerate(hdr,1):
    c=ws.cell(row=4,column=i,value=h); c.font=Font(name=F,bold=True); c.fill=PatternFill('solid',fgColor='EEF1F6'); c.border=BD; c.alignment=Alignment(wrap_text=True,horizontal='center')
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=16
r=5
for t in TIERS+['סה"כ']:
    if t=='סה"כ': m=s_mg[1]; a=s_all[1]; n=len(scope); nm=sum(covm.values())
    else: m=s_mg[0][t]; a=s_all[0][t]; n=cov[t]; nm=covm[t]
    for i,v in enumerate([t,n,nm,m['rows'],m['mob'],m['em'],a['rows'],a['mob'],a['em']],1):
        c=ws.cell(row=r,column=i,value=v); c.font=Font(name=F,bold=(t=='סה"כ')); c.border=BD
    r+=1
out='/home/user/Claude-projects/fizikal-crm/out/Fizikal_רשימת_דיוור_השקה_2026-11.xlsx'; wb.save(out); print('saved',out,len(mg_rows),len(all_rows))
