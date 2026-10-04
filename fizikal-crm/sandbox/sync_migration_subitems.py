# קוד ל-execute_code (סנדבוקס של monday MCP), לא להרצה מקומית.
# סנכרון אנשי קשר → סאב-אייטמים בבורד "מעבר לגרסה החדשה".
# מקור האמת: בורד אנשי קשר 5103469435 (לפי מספר סאפ) → אבות 5102863984 → סאב-אייטמים 5102865659.
# התאמה: מספר עובד → טלפון → שם. יוצר חסרים, מעדכן קיימים. 3 תהליכונים, תקציב 250 שניות.
# לא נוגע לעולם ב: color_mm6jm46b date_mm6je57s date_mm6j7fvm text_mm6jehn6 (אב) · color_mm6jkycq date_mm6j7e3f (סאב). אלה של בני אדם.
# ה-MCP חותך ב-60 שניות אבל הסנדבוקס ממשיך — לאמת בקריאה חוזרת. להריץ שוב עד ששתי ריצות רצופות לא מוסיפות כלום.
# 10/09/2026: הסריקה של אנשי הקשר מסוננת ל-SAP לא ריק — המכסה היומית היא לפי מורכבות, וסריקות מלאות של 10k שורות שרפו אותה.
import json, time, requests, re, threading, queue, collections as C
API='https://api.monday.com/v2'
CB='5103469435'; PB='5102863984'; SB='5102865659'
START=time.time(); BUDGET=250
TL=threading.local()
def S():
    if not hasattr(TL,'s'): TL.s=requests.Session()
    return TL.s
def gq(q,v=None):
    for a in range(6):
        try:
            r=S().post(API,json={'query':q,'variables':v or {}},timeout=60); b=r.json()
        except Exception:
            time.sleep(3*(a+1)); continue      # דף HTML/5xx חולף — לא JSON
        if 'errors' in b:
            m=json.dumps(b['errors'],ensure_ascii=False)[:250]
            if 'DAILY_LIMIT' in m: raise SystemExit('DAILY_LIMIT_EXCEEDED')
            if any(k in m for k in ('TooManyConcurrent','Rate','rate','Complexity','lock')): time.sleep(2*(a+1)); continue
            raise RuntimeError(m)
        return b['data']
    raise RuntimeError('retries exhausted')
def scan(board, cols, want_parent=False, only_col=None):
    # only_col: לסרוק רק שורות שבהן העמודה לא ריקה — חוסך ~2/3 מהמורכבות בבורד אנשי קשר (9,431 → ~3,300)
    pf=' parent_item{id}' if want_parent else ''
    qp=',query_params:{rules:[{column_id:"%s",compare_value:[""],operator:is_not_empty}]}'%only_col if only_col else ''
    out=[]; cur=None; first=True
    while True:
        if first:
            r=gq('query($b:[ID!],$c:[String!]){boards(ids:$b){items_page(limit:500%s){cursor items{id name%s column_values(ids:$c){id text}}}}}'%(qp,pf),{'b':[board],'c':cols})
            p=r['boards'][0]['items_page']; first=False
        else:
            r=gq('query($cur:String!,$c:[String!]){next_items_page(limit:500,cursor:$cur){cursor items{id name%s column_values(ids:$c){id text}}}}'%pf,{'cur':cur,'c':cols})
            p=r['next_items_page']
        for it in p['items']:
            d={'id':it['id'],'name':re.sub(r'\s+',' ',it['name'] or '').strip(),
               'cv':{c['id']:(c['text'] or '').strip() for c in it['column_values']}}
            if want_parent: d['parent']=(it['parent_item'] or {}).get('id')
            out.append(d)
        cur=p['cursor']
        if not cur: break
    return out
C_SAP='text_mm6vke74'; C_EMP='text_mm6v429b'; C_ROLE='dropdown_mm6v5453'
C_PH='text_mm6vqqpn'; C_EM='text_mm6vwyzq'; C_TZ='text_mm6v31b9'; C_BR='text_mm6vw0jh'
P_SAP='text_mm6j4ytj'
ROLE='text_mm6j4yqt'; PH='text_mm6jktf8'; EM='email_mm6jkzhn'; TZ='text_mm6jv77n'
BR='text_mm6jq7d5'; OK='color_mm6jvqmv'; EMP='text_mm6vra3x'
sap2item={p['cv'].get(P_SAP,''):p['id'] for p in scan(PB,[P_SAP]) if p['cv'].get(P_SAP,'')}
contacts=scan(CB,[C_SAP,C_EMP,C_ROLE,C_PH,C_EM,C_TZ,C_BR],only_col=C_SAP)
subs=scan(SB,[PH,BR,EMP],want_parent=True)
byp=C.defaultdict(list)
for s in subs: byp[s['parent']].append(s)
tgt=C.defaultdict(list)
for c in contacts:
    for s in [x.strip() for x in c['cv'].get(C_SAP,'').split(',') if x.strip()]:
        iid=sap2item.get(s)
        if iid: tgt[iid].append(c)
print('subitems:%d  target:%d rows over %d parents'%(len(subs),sum(len(v) for v in tgt.values()),len(tgt)),flush=True)
def esc(s): return json.dumps(str(s or ''),ensure_ascii=False)
def cvals(c):
    v=c['cv']; em=v.get(C_EM,''); tz=v.get(C_TZ,''); ph=v.get(C_PH,'')
    br=(v.get(C_BR,'').split(',')[0] or '').strip()
    d={ROLE:v.get(C_ROLE,''),PH:ph,TZ:tz,BR:br,OK:{'label':'כן' if (em and tz) else 'לא'},EMP:v.get(C_EMP,'')}
    if em: d[EM]={'email':em,'text':em}
    return json.dumps(json.dumps(d,ensure_ascii=False),ensure_ascii=False)
STAT=C.Counter(); LOCK=threading.Lock()
def work(q):
    while True:
        try: iid,rows=q.get_nowait()
        except queue.Empty: return
        if time.time()-START>BUDGET:
            with LOCK: STAT['deferred']+=1
            continue
        ex=list(byp.get(iid,[])); used=set(); upd=[]; cre=[]
        for c in rows:
            emp=c['cv'].get(C_EMP,''); ph=c['cv'].get(C_PH,''); nm=c['name'].lower()
            m=None
            if emp:
                for s in ex:
                    if s['id'] not in used and s['cv'].get(EMP)==emp: m=s; break
            if not m and ph:
                for s in ex:
                    if s['id'] not in used and s['cv'].get(PH)==ph: m=s; break
            if not m:
                for s in ex:
                    if s['id'] not in used and s['name'].lower()==nm: m=s; break
            if m: used.add(m['id']); upd.append('u%s: change_multiple_column_values(board_id:%s,item_id:%s,column_values:%s){id}'%(m['id'],SB,m['id'],cvals(c)))
            else: cre.append('c%d: create_subitem(parent_item_id:%s,item_name:%s,column_values:%s){id}'%(len(cre),iid,esc(c['name'] or 'ללא שם'),cvals(c)))
        try:
            for i in range(0,len(cre),10): gq('mutation{'+''.join(cre[i:i+10])+'}')
            for i in range(0,len(upd),20): gq('mutation{'+''.join(upd[i:i+20])+'}')
            with LOCK: STAT['created']+=len(cre); STAT['updated']+=len(upd); STAT['parents']+=1
        except SystemExit: raise
        except Exception as e:
            with LOCK: STAT['errors']+=1
            print('ERR',iid,str(e)[:100],flush=True)
need=[(iid,rows) for iid,rows in tgt.items() if len(byp.get(iid,[]))<len(rows)]
print('parents short:%d  rows missing:%d'%(len(need),sum(len(r)-len(byp.get(i,[])) for i,r in need)),flush=True)
Q=queue.Queue()
for iid,rows in sorted(need,key=lambda kv: len(kv[1])): Q.put((iid,rows))
th=[threading.Thread(target=work,args=(Q,),daemon=True) for _ in range(3)]
[t.start() for t in th]; [t.join() for t in th]
print('RESULT',dict(STAT),'elapsed %.0fs'%(time.time()-START))
