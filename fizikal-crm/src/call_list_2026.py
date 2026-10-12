"""רשימת לקוחות לאימות פרטים ע"י חברה חיצונית — שורה ללקוח, גיליון לרמה + לא מסווג.
הנציג מגיע למנהל הפעילות (עמודות העזר: למי לבקש, טלפונים) והמנהל ממלא את 13 השדות הריקים.
קלט: dump JSON.gz של ארבעה בורדים + collections.json / collections_missing.json מבורד הגבייה.
הרצה: python3 src/call_list_2026.py <dump.json.gz>
"""
import gzip,json,collections as C,re,sys
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.datavalidation import DataValidation
D=json.load(gzip.open(sys.argv[1] if len(sys.argv)>1 else '/tmp/claude-0/-home-user-Claude-projects/9bedc6a6-f60e-5cce-a499-d0c73e5871ea/scratchpad/dl/all.json.gz','rt',encoding='utf-8'))
COLM=json.load(open('collections_missing.json',encoding='utf-8')); COL=json.load(open('collections.json',encoding='utf-8'))
comp,br,onb,con=D['companies'],D['branches'],D['onboarding'],D['contacts']
norm=lambda s:(s or '').strip()
mob=lambda p: re.fullmatch(r'05\d{8}',re.sub(r'\D','',p or '')) is not None
clean=lambda p: re.sub(r'\D','',p or '')
BRC=C.defaultdict(list); SAPB=C.defaultdict(list)
for b in br:
    b['_cc']=norm(b['text_mm6v5fq7']); b['_bn']=norm(b['text_mm6vxmyx'])
    if b['_cc']: BRC[b['_cc']].append(b)
    if norm(b['text_mm6vwm4f']): SAPB[norm(b['text_mm6vwm4f'])].append(b)
CON=C.defaultdict(list)
for c in con: CON[norm(c['text_mm6v6myv'])].append(c)
COLSAP={x['sap']:x for x in COL}
def is_mgmt(c):
    r=norm(c['dropdown_mm6v5453']); g=norm(c['dropdown_mm6v5svk'])
    return r in('ניהול חברה','הרשאת על') or (r=='מנהל' and g in('הנהלה','הכל','מנהל'))
def rank(c):
    r=norm(c['dropdown_mm6v5453']); return 0 if r=='ניהול חברה' else 1 if r=='הרשאת על' else 2
def reach(ccs,sap=''):
    """who to ask for + phones: management mobiles first, then branch phones, then collections"""
    people=[]; seen=set()
    for cc in ccs:
        for c in sorted((c for c in CON.get(cc,[]) if is_mgmt(c)),key=rank):
            p=clean(c['text_mm6vqqpn'])
            if mob(p) and p not in seen: seen.add(p); people.append((norm(c['name']),p))
    phones=[f"{n} {p}" for n,p in people[:3]]
    if not phones:
        bp=[clean(b['text_mm6vnk3m']) for cc in ccs for b in BRC.get(cc,[]) if clean(b['text_mm6vnk3m'])]
        phones=[f"סניף {p}" for p in dict.fromkeys(bp).keys()][:3]
    if not phones and sap in COLSAP and clean(COLSAP[sap]['phone']): phones=[f"גבייה {clean(COLSAP[sap]['phone'])}"]
    ask=', '.join(n for n,_ in people[:3])
    return ask, ' · '.join(phones), len(people)
tier_of=lambda g:{'VIP':'VIP','גדולים':'גדול','בינוניים':'בינוני','קטנים':'קטן','מוב':'מוב'}.get(g.replace(' - לא פעילים','').strip(),g)
HELP=['מספר לקוח','שם הלקוח אצלנו','מצב אצלנו','למי לבקש','טלפונים להתקשרות']
FIELDS=['שם חברה','ח.פ','שם לחשבונית','מספר חברה','מספר סניף','שם סניף','סטטוס לקוח','שם מלא','תפקיד','ת.ז','מייל','טלפון','אומת בתאריך']
rows=C.defaultdict(list); linked=set(); stats=C.Counter()
for co in comp:
    sap=norm(co['text_mm6ccg01']); st='לא פעיל 3+ חודשים' if co['color_mm6cyme4']=='לא פעיל' else co['color_mm6cyme4']
    ccs=list(dict.fromkeys([norm(s.get('text_mm76nbjb')) for s in co.get('subitems',[]) if norm(s.get('text_mm76nbjb'))] or [b['_cc'] for b in SAPB.get(sap,[])]))
    linked.update(ccs)
    ask,ph,n=reach(ccs,sap)
    if not ph: stats['no_phone']+=1
    if not n: stats['no_mgmt']+=1
    rows[tier_of(co['group'])].append([sap,norm(co['name']),st,ask,ph])
# unclassified
LIVE={'אונבורדינג הסתיים','עלה לאוויר- ליווי חודש 1','ליווי חודש 2','באוויר - ממתין לסליקה','הדרכה 1','הדרכה 2','הדרכה עיכובים','ממתין להדרכה','שלבי הסבה (אופציונלי)','ממתין לסליקה (מכירות)','סיכום מכירה'}
seen_cc=set(); x=C.Counter()
for o in onb:
    if not (o['date_mm2fvzsy']>='2026-05-01' or (not o['date_mm2fvzsy'] and o['status'] in LIVE and o['date4']>='2026-05-01')): continue
    if o['status']=='ביטל עסקה': x['cancelled']+=1; continue
    cc=norm(o['text_mm2bb1bq'])
    if cc and cc in linked: x['paying']+=1; continue
    if cc and cc in BRC: seen_cc.add(cc); ask,ph,n=reach([cc]); src='אונבורדינג מאי+'
    else: ask,ph='',''; src='אונבורדינג מאי+, בלי מספר חברה' if not cc else 'אונבורדינג מאי+'
    x['added']+=1
    rows['לא מסווג'].append([cc or '',norm(o['name']),src,ask,ph])
act=C.defaultdict(list)
for b in br:
    if b['color_mm6vn3xc']=='פעיל' and b['_cc'] and b['_cc'] not in linked and b['_cc'] not in seen_cc: act[b['_cc']].append(b)
for cc,bs in act.items():
    ask,ph,n=reach([cc]); rows['לא מסווג'].append([cc,norm(bs[0]['text_mm6vcn1r']) or norm(bs[0]['name']),'פעיל במערכת, לא בדוח הכנסות',ask,ph])
CH=['פרופיט','פיטפורלס','גרייט שייפ','סוויפט','מטי מועדוני','מגדלי הים התיכון','בארבור','אייקון פיטנס']
for m in COLM:
    chain=next((c for c in CH if c in m['name']),'')
    ph=('גבייה '+clean(m['phone'])) if clean(m['phone']) else ''
    rows['לא מסווג'].append([m['sap'],norm(m['name']),'בורד גביה, לא בדוח אוגוסט'+(f' (כנראה סניף של רשת {chain})' if chain else ''),'',ph+((' · '+m['email']) if m.get('email') else '')])
print('onboarding',dict(x),'| unclassified from system',len(act),'| from collections',len(COLM)); print(dict(stats))
for k in ['VIP','גדול','בינוני','קטן','מוב','לא מסווג']: print(k,len(rows[k]),'rows; with phone',sum(1 for r in rows[k] if r[4]))
# write
F='Arial'; wb=openpyxl.Workbook(); wb.remove(wb.active)
thin=Side(style='thin',color='C9D1E0'); B=Border(top=thin,bottom=thin,left=thin,right=thin)
W=[11,30,22,26,40]+[26,12,26,10,9,22,12,24,14,12,28,13,13]
for k in ['VIP','גדול','בינוני','קטן','מוב','לא מסווג']:
    ws=wb.create_sheet(k); ws.sheet_view.rightToLeft=True
    for i,(h,w) in enumerate(zip(HELP+FIELDS,W),1):
        c=ws.cell(row=1,column=i,value=h); c.font=Font(name=F,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='6B7280' if i<=5 else '1F2A44'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=B
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws.row_dimensions[1].height=30
    for r,row in enumerate(rows[k],2):
        for i in range(1,len(HELP)+len(FIELDS)+1):
            v=row[i-1] if i<=5 else None
            c=ws.cell(row=r,column=i,value=v or None); c.font=Font(name=F,size=10,color='555555' if i<=5 else '000000'); c.border=B
            if i<=5: c.fill=PatternFill('solid',fgColor='F3F4F6')
            if i in(1,7,9,10,15,17): c.number_format='@'
    n=len(rows[k])+1; last=openpyxl.utils.get_column_letter(len(HELP)+len(FIELDS))
    dv=DataValidation(type='list',formula1='"בעלים,מנהל סניף,הנהלה אחרת,קבלה,מנהל משמרת,אחר"',allow_blank=True); ws.add_data_validation(dv); dv.add(f'N2:N{n+2000}')
    dv2=DataValidation(type='list',formula1='"פעיל,לא פעיל"',allow_blank=True); ws.add_data_validation(dv2); dv2.add(f'L2:L{n+2000}')
    ws.freeze_panes='F2'; ws.auto_filter.ref=f'A1:{last}{max(n,2)}'
ws=wb.create_sheet('הסבר',0); ws.sheet_view.rightToLeft=True; ws.column_dimensions['A'].width=115
L=["רשימת לקוחות לאימות פרטים — Fizikal, 12/10/2026",
"גיליון לכל רמת לקוח (לפי דוח ההכנסות אוגוסט 2026): VIP · גדול · בינוני · קטן · מוב. 'לא מסווג' = מי שלא בדוח אוגוסט: אונבורדינג שעלה לאוויר ממאי, חברות פעילות במערכת בלי תשלום, וכרטיסים מבורד הגבייה.",
"שורה = לקוח. 5 העמודות האפורות הן לנציג בלבד: מי הלקוח, מה מצבו אצלנו, למי לבקש ובאיזה טלפון. אין להעתיק מהן לשדות.",
"13 העמודות הכהות ריקות בכוונה. הנציג מגיע למי שמנהל את הפעילות, והמנהל ממלא: פרטי העסק ואנשי ההנהלה. אדם נוסף או סניף נוסף = שורה נוספת מתחת לאותו לקוח, עם אותו מספר לקוח בעמודה הראשונה.",
"תפקיד מרשימה: בעלים · מנהל סניף · הנהלה אחרת · קבלה · מנהל משמרת · אחר. סטטוס: פעיל / לא פעיל. 'אומת בתאריך' ריק = לא עברה שיחה.",
"'למי לבקש' = אנשי הנהלה שרשומים אצלנו (ניהול חברה, הרשאת על, מנהל בקבוצת הנהלה), עד שלושה, עם הנייד. כשאין, טלפון הסניף, ואם גם זה חסר, הטלפון מבורד הגבייה."]
for i,t in enumerate(L,1):
    c=ws.cell(row=i,column=1,value=t); c.font=Font(name=F,bold=(i==1),size=13 if i==1 else 11); c.alignment=Alignment(wrap_text=True,vertical='top',horizontal='right'); ws.row_dimensions[i].height=22 if i==1 else 36
out='/home/user/Claude-projects/fizikal-crm/out/Fizikal_רשימת_לקוחות_לאימות_2026-10.xlsx'; wb.save(out); print('saved')
