"""רשימת לקוחות לאימות פרטי קשר (חברה חיצונית מתקשרת) — גיליון לכל רמת לקוח + לא מסווג.
קלט: dump של ארבעה בורדים (חברות וסניפים, סניפים, אנשי קשר, onboarding) כ-JSON.gz שנוצר ב-execute_code של monday.
הנהלה בלבד: ניהול חברה / הרשאת על / מנהל בקבוצת הנהלה-הכל-מנהל. שורה = אדם בסניף.
לא מסווג = אונבורדינג שעלה לאוויר ממאי 2026 ולא בדוח ההכנסות + חברות פעילות במערכת בלי לקוח משלם.
הרצה: python3 src/call_list_2026.py <dump.json.gz>
"""
import sys
import gzip,json,collections as C,re,sys
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.datavalidation import DataValidation
D=json.load(gzip.open(sys.argv[1],'rt',encoding='utf-8'))
comp,br,onb,con=D['companies'],D['branches'],D['onboarding'],D['contacts']
norm=lambda s:(s or '').strip()
# ---- branches index
BR={}; BRC=C.defaultdict(list)
for b in br:
    cc,bn=norm(b['text_mm6v5fq7']),norm(b['text_mm6vxmyx'])
    b['_cc'],b['_bn']=cc,bn
    if cc: BR[(cc,bn)]=b; BRC[cc].append(b)
# ---- contacts index by company code
CON=C.defaultdict(list)
for c in con:
    cc=norm(c['text_mm6v6myv']); c['_cc']=cc
    if cc: CON[cc].append(c)
print('contacts without company code',sum(1 for c in con if not c['_cc']))
def is_mgmt(c):
    r=norm(c['dropdown_mm6v5453']); g=norm(c['dropdown_mm6v5svk'])
    return r in('ניהול חברה','הרשאת על') or (r=='מנהל' and g in('הנהלה','הכל','מנהל'))
# ---- role mapping
def role(c):
    r=norm(c['dropdown_mm6v5453']); g=norm(c['dropdown_mm6v5svk'])
    if r in('ניהול חברה','הרשאת על'): return 'הנהלה אחרת'
    if r=='מנהל משמרת': return 'מנהל משמרת'
    if r=='מנהל':
        if g in('הנהלה','הכל','מנהל'): return 'מנהל סניף'
        if g in('קבלה','דלפק'): return 'קבלה'
        return 'אחר'
    return 'אחר'
TIERS=['VIP','גדול','בינוני','קטן','מוב']
def tier_of(g):
    g=g.replace(' - לא פעילים','').strip()
    return {'VIP':'VIP','גדולים':'גדול','בינוניים':'בינוני','קטנים':'קטן','מוב':'מוב'}.get(g,g)
HEAD=['שם חברה','ח.פ','שם לחשבונית','מספר חברה','מספר סניף','שם סניף','סטטוס לקוח','שם מלא','תפקיד','ת.ז','מייל','טלפון','אומת בתאריך','מספר לקוח (הנה"ח)','הרשאה במערכת','קבוצת עובדים','מקור']
rows=C.defaultdict(list); stats=C.Counter(); linked_cc=set(); linked_sap=set()
def people_rows(cc,bn,base,src):
    out=[]; seen=set()
    for c in CON.get(cc,[]):
        if not is_mgmt(c): continue
        bns=[x.strip() for x in re.split(r'[,،]',c['dropdown_mm6vwz6r']) if x.strip()]
        bnames=[x.strip() for x in re.split(r'[,،]',c['text_mm6vw0jh']) if x.strip()]
        if bns:
            if bn not in bns: continue
        elif bnames:
            if norm(base[5]) not in bnames: continue
        elif bn!=min((b['_bn'] for b in BRC.get(cc,[])),default=bn): continue
        key=(c['id'],bn)
        if key in seen: continue
        seen.add(key)
        out.append(base+[norm(c['name']),role(c),norm(c['text_mm6v31b9']),norm(c['text_mm6vwyzq']),norm(c['text_mm6vqqpn']),'',base_sap[0],norm(c['dropdown_mm6v5453']),norm(c['dropdown_mm6v5svk']),src])
    return out
for co in comp:
    sap=norm(co['text_mm6ccg01']); hp=norm(co['text_mm5ebgdf']); st='לא פעיל' if co['color_mm6cyme4']=='לא פעיל' else 'פעיל'
    tier=tier_of(co['group']); name=norm(co['name']); base_sap=[sap]
    branches=[]
    for s in co.get('subitems',[]):
        cc,bn=norm(s.get('text_mm76nbjb')),norm(s.get('text_mm76sk79'))
        if cc: branches.append(BR.get((cc,bn)) or {'_cc':cc,'_bn':bn,'name':s['name'],'text_mm6vx83h':norm(s.get('text_mm76hz44')),'text_mm6v1yfx':''})
    if not branches:
        branches=[b for b in br if norm(b['text_mm6vwm4f'])==sap]
    if not branches: stats['clients_no_branch']+=1
    for b in branches: linked_cc.add(b['_cc'])
    linked_sap.add(sap)
    n0=len(rows[tier])
    if not branches:
        rows[tier].append([name,hp,name,'','','',st,'','','','','','',sap,'','','בורד חברות'])
        continue
    for b in branches:
        base=[name,hp or norm(b.get('text_mm6vx83h')),norm(b.get('text_mm6v1yfx')) or name,b['_cc'],b['_bn'],norm(b.get('name')),st]
        pr=people_rows(b['_cc'],b['_bn'],base,'בורד חברות')
        if not pr: pr=[base+['','','','','','',sap,'','','בורד חברות']]; stats['branches_no_people']+=1
        else: stats['people_rows']+=len(pr)
        rows[tier]+=pr
    if len(rows[tier])==n0: stats['clients_no_rows']+=1
# ---- לא מסווג: onboarding since May not linked to a paying client, plus active branches with no paying client
LIVE={'אונבורדינג הסתיים','עלה לאוויר- ליווי חודש 1','ליווי חודש 2','באוויר - ממתין לסליקה','הדרכה 1','הדרכה 2','הדרכה עיכובים','ממתין להדרכה','שלבי הסבה (אופציונלי)','ממתין לסליקה (מכירות)','סיכום מכירה'}
onb_may=[o for o in onb if o['date_mm2fvzsy']>='2026-05-01' or (not o['date_mm2fvzsy'] and o['status'] in LIVE and o['date4']>='2026-05-01')]
xc=C.Counter(); unl=[]
for o in onb_may:
    cc=norm(o['text_mm2bb1bq'])
    if o['status']=='ביטל עסקה': xc['cancelled']+=1; continue
    if cc and cc in linked_cc: xc['linked_paying']+=1; continue
    if cc and cc in BRC: xc['in_system_not_paying']+=1; src='אונבורדינג מאי+ (במערכת, לא בדוח הכנסות)'
    elif cc: xc['code_unknown']+=1; src='אונבורדינג מאי+ (מספר חברה לא נמצא בבורד הסניפים)'
    else: xc['no_code']+=1; src='אונבורדינג מאי+ (בלי מספר חברה)'
    unl.append((o,cc,src))
seen_cc=set()
for o,cc,src in unl:
    name=norm(o['name'])
    if cc and cc in BRC:
        seen_cc.add(cc)
        for b in BRC[cc]:
            if b['color_mm6vn3xc']!='פעיל': continue
            base=[norm(b['text_mm6vcn1r']) or name,norm(b['text_mm6vx83h']),norm(b['text_mm6v1yfx']) or name,cc,b['_bn'],norm(b['name']),'פעיל']
            pr=people_rows(cc,b['_bn'],base,src) if (base_sap:=['']) else []
            rows['לא מסווג']+=pr or [base+['','','','','','','','','',src]]
    else:
        rows['לא מסווג'].append([name,'',name,cc,'',norm(o['text_mm2ba4y1']),'פעיל','','','','','','','','','',src])
# active branches of companies with no paying client (and not already added)
act_cc=C.defaultdict(list)
for b in br:
    if b['color_mm6vn3xc']=='פעיל' and b['_cc'] and b['_cc'] not in linked_cc and b['_cc'] not in seen_cc: act_cc[b['_cc']].append(b)
for cc,bs in act_cc.items():
    for b in bs:
        base=[norm(b['text_mm6vcn1r']),norm(b['text_mm6vx83h']),norm(b['text_mm6v1yfx']),cc,b['_bn'],norm(b['name']),'פעיל']
        base_sap=['']
        pr=people_rows(cc,b['_bn'],base,'פעיל במערכת, לא בדוח הכנסות')
        rows['לא מסווג']+=pr or [base+['','','','','','','','','','פעיל במערכת, לא בדוח הכנסות']]
stats['unclassified_companies_from_system']=len(act_cc)
# ---- collections board: SAP cards billed today but absent from the August list (and not a known family ח.פ)
import os
cm=os.path.join(os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else '.','collections_missing.json')
try: COLM=json.load(open(cm,encoding='utf-8'))
except Exception: COLM=[]
CHAINS=['פרופיט','פיטפורלס','גרייט שייפ','סוויפט','מטי מועדוני','מגדלי הים התיכון','בארבור','אייקון פיטנס']
for x in COLM:
    chain=next((c for c in CHAINS if c in x['name']),'')
    src='בורד גביה, לא בדוח הכנסות אוגוסט'+(f' (כנראה סניף של רשת {chain} שכבר ברשימה)' if chain else '')
    extra=' · '.join(v for v in [x.get('phone',''),x.get('email','')] if v)
    if extra: src+=' · פרטי קשר בגביה: '+extra
    rows['לא מסווג'].append([norm(x['name']),norm(x.get('cc','')),norm(x['name']),'','','','פעיל','','','','','','',x['sap'],'','',src])
stats['from_collections']=len(COLM)

print('onboarding since May:',len(onb_may),dict(xc)); print(dict(stats))
for k in TIERS+['לא מסווג']: print(k,len(rows[k]),'rows',len({(r[0],r[3]) for r in rows[k]}),'companies')
# ---- write
F='Arial'; wb=openpyxl.Workbook(); wb.remove(wb.active)
thin=Side(style='thin',color='C9D1E0'); B=Border(top=thin,bottom=thin,left=thin,right=thin)
widths=[28,12,26,10,9,24,11,24,13,12,28,13,13,13,13,13,34]
for k in TIERS+['לא מסווג']:
    ws=wb.create_sheet(k); ws.sheet_view.rightToLeft=True
    for i,(h,w) in enumerate(zip(HEAD,widths),1):
        c=ws.cell(row=1,column=i,value=h); c.font=Font(name=F,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F2A44' if i<=13 else '6B7280'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=B
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws.row_dimensions[1].height=30
    for r,row in enumerate(rows[k],2):
        for i,v in enumerate(row,1):
            c=ws.cell(row=r,column=i,value=v if v!='' else None); c.font=Font(name=F,size=10); c.border=B
            if i in(2,4,5,10,12,14): c.number_format='@'
    n=len(rows[k])+1
    dv=DataValidation(type='list',formula1='"בעלים,מנהל סניף,הנהלה אחרת,קבלה,מנהל משמרת,אחר"',allow_blank=True); ws.add_data_validation(dv); dv.add(f'I2:I{max(n,2)+500}')
    dv2=DataValidation(type='list',formula1='"פעיל,לא פעיל"',allow_blank=True); ws.add_data_validation(dv2); dv2.add(f'G2:G{max(n,2)+500}')
    ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:{openpyxl.utils.get_column_letter(len(HEAD))}{max(n,2)}'
ws=wb.create_sheet('הסבר',0); ws.sheet_view.rightToLeft=True; ws.column_dimensions['A'].width=110
lines=["רשימת לקוחות לאימות פרטי קשר — Fizikal, נתוני 12/10/2026",
"גיליון לכל רמת לקוח (לפי ריטיינר אוגוסט 2026): VIP · גדול · בינוני · קטן · מוב, ועוד 'לא מסווג'.",
"שורה = אדם בסניף. אדם שעובד בכמה סניפים מופיע בכמה שורות עם אותה ת.ז.",
"13 העמודות הראשונות הן השדות שסיכמנו. 4 העמודות האפורות בסוף הן עזר: מספר הלקוח בהנה\"ח (המפתח לטעינה חזרה), ההרשאה והקבוצה כפי שרשומות היום במערכת, ומקור השורה.",
"האנשים ברשימה הם ההנהלה בלבד, לפי ההרשאות במערכת פיזיקל (ייצוא 03/09/2026): ניהול חברה, הרשאת על, ומנהל בקבוצת הנהלה. שאר העובדים לא נכללו. 'תפקיד' תורגם מההרשאה: ניהול חברה/הרשאת על → הנהלה אחרת · מנהל+הנהלה → מנהל סניף · מנהל+קבלה → קבלה · מנהל משמרת → מנהל משמרת · השאר → אחר. המתקשר מאמת ומתקן.",
"'לא מסווג': (1) לקוחות מאונבורדינג שעלו לאוויר ממאי 2026 ולא נמצאו בדוח ההכנסות, (2) כל חברה שפעילה במערכת אבל לא משלמת לפי דוח ההכנסות, (3) כרטיסי לקוח מבורד הגבייה שאינם בדוח אוגוסט. עמודת 'מקור' אומרת מאיפה כל שורה.",
"'אומת בתאריך' ריק = השורה לא עברה שיחה."]
for i,t in enumerate(lines,1):
    c=ws.cell(row=i,column=1,value=t); c.font=Font(name=F,bold=(i==1),size=11 if i>1 else 13); c.alignment=Alignment(wrap_text=True,vertical='top',horizontal='right'); ws.row_dimensions[i].height=34 if i>1 else 22
out='/home/user/Claude-projects/fizikal-crm/out/Fizikal_רשימת_לקוחות_לאימות_2026-10.xlsx'; wb.save(out); print('saved',out)
json.dump({'rows':{k:len(v) for k,v in rows.items()},'stats':dict(stats),'onb':dict(xc)},open('stats.json','w'),ensure_ascii=False)
