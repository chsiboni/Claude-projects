# קוד ל-execute_code (סנדבוקס של monday MCP), לא להרצה מקומית.
# בונה את "חברות וסניפים": עותק של בורד חברות 5102514504 (לשעבר "לקוחות") עם הסניפים כסאב-אייטמים.
# החלטת Chen 09/09/2026: לא נוגעים בבורד הקיים. עותק → הוא בודק → אישור → הישן לארכיון.
#
# vars:
#   MAP       — JSON: [[sap, company_code, branch_no], ...] (584, מ-out/branch_subitems_<tag>.json; ~12KB)
#   NEW_BOARD — ריק = לשכפל; אחרת מזהה של העותק להמשך ריצה
# פרטי הסניפים נקראים מבורד סניפים 5103467882 (1,679) — אין צורך להעביר אותם.
# אידמפוטנטי לפי (מס חברה, מס סניף) תחת אותו אב. ה-MCP חותך ב-60 שניות; הסנדבוקס ממשיך עד BUDGET.

import json, os, time, requests, threading, queue, collections as C

API = 'https://api.monday.com/v2'
SRC = '5102514504'; SAP_COL = 'text_mm6ccg01'
BR = '5103467882'
B = {'code': 'text_mm6v5fq7', 'branch_no': 'text_mm6vxmyx', 'hp': 'text_mm6vx83h', 'active': 'color_mm6vn3xc',
     'joined': 'date_mm6vqstg', 'phone': 'text_mm6vnk3m', 'email': 'dropdown_mm6vfdrh', 'address': 'text_mm6veycj',
     'level': 'color_mm6v5qfj', 'members': 'numeric_mm6vm8dh', 'credit': 'color_mm6v7fs1', 'app': 'color_mm6vzftb',
     'move': 'color_mm6vzt56'}
START = time.time(); BUDGET = 240
TL = threading.local()
def S():
    if not hasattr(TL, 's'): TL.s = requests.Session()
    return TL.s
def gq(q, v=None):
    for a in range(6):
        try:
            b = S().post(API, json={'query': q, 'variables': v or {}}, timeout=60).json()
        except Exception:
            time.sleep(3 * (a + 1)); continue
        if 'errors' in b:
            m = json.dumps(b['errors'], ensure_ascii=False)[:300]
            if 'DAILY_LIMIT' in m: raise SystemExit('DAILY_LIMIT_EXCEEDED')
            if any(k in m for k in ('TooManyConcurrent', 'Rate', 'rate', 'Complexity', 'lock')):
                time.sleep(2 * (a + 1)); continue
            raise RuntimeError(m)
        return b['data']
    raise RuntimeError('retries exhausted')

def scan(board, cols, want_parent=False):
    pf = ' parent_item{id}' if want_parent else ''
    out = []; cur = None; first = True
    while True:
        if first:
            r = gq('query($b:[ID!],$c:[String!]){boards(ids:$b){items_page(limit:500){cursor items{id name%s column_values(ids:$c){id text}}}}}' % pf, {'b': [board], 'c': cols})
            p = r['boards'][0]['items_page']; first = False
        else:
            r = gq('query($cur:String!,$c:[String!]){next_items_page(limit:500,cursor:$cur){cursor items{id name%s column_values(ids:$c){id text}}}}' % pf, {'cur': cur, 'c': cols})
            p = r['next_items_page']
        for it in p['items']:
            d = {'id': it['id'], 'name': (it['name'] or '').strip(), 'cv': {c['id']: (c['text'] or '').strip() for c in it['column_values']}}
            if want_parent: d['parent'] = (it['parent_item'] or {}).get('id')
            out.append(d)
        cur = p['cursor']
        if not cur: break
    return out

MAP = json.loads(os.environ['MAP'])

# --- 1. הבורד ---
new = os.environ.get('NEW_BOARD', '').strip()
if not new:
    r = gq('mutation{duplicate_board(board_id:%s,duplicate_type:duplicate_board_with_pulses,board_name:"חברות וסניפים"){board{id}}}' % SRC)
    new = r['duplicate_board']['board']['id']
    print('NEW_BOARD', new, flush=True)
    time.sleep(25)                        # השכפול אסינכרוני
parents = scan(new, [SAP_COL])
print('copy items', len(parents), flush=True)
sap2id = {p['cv'].get(SAP_COL, ''): p['id'] for p in parents if p['cv'].get(SAP_COL, '')}

# --- 2. בורד הסאב-אייטמים והעמודות ---
info = gq('query{boards(ids:[%s]){columns{id type settings_str}}}' % new)['boards'][0]['columns']
subcol = next((c for c in info if c['type'] == 'subtasks'), None)
if not subcol:
    r = gq('mutation{create_subitem(parent_item_id:%s,item_name:"__init__"){id board{id}}}' % parents[0]['id'])
    sub_board = r['create_subitem']['board']['id']
    gq('mutation{delete_item(item_id:%s){id}}' % r['create_subitem']['id'])
else:
    sub_board = str(json.loads(subcol['settings_str'])['boardIds'][0])
print('SUB_BOARD', sub_board, flush=True)

WANT = [('מס חברה', 'text'), ('מס סניף', 'text'), ('ח.פ', 'text'), ('פעיל בתוכנה', 'status'), ('כמות מתאמנים', 'numbers'),
        ('רמה', 'text'), ('אפליקציה', 'text'), ('MOVE', 'checkbox'), ('סולק', 'text'), ('הצטרפות', 'date'),
        ('טלפון', 'text'), ('מייל', 'text'), ('כתובת', 'text')]
cols = {c['title']: c for c in gq('query{boards(ids:[%s]){columns{id title type}}}' % sub_board)['boards'][0]['columns']}
for title, typ in WANT:
    if title not in cols:
        cols[title] = gq('mutation{create_column(board_id:%s,title:"%s",column_type:%s){id title type}}' % (sub_board, title, typ))['create_column']
CID = {t: cols[t]['id'] for t, _ in WANT}
print('COLUMNS', json.dumps(CID, ensure_ascii=False), flush=True)

# --- 3. נתוני הסניפים + מה כבר קיים ---
br = {(b['cv'][B['code']], b['cv'][B['branch_no']]): b for b in scan(BR, list(B.values()))}
existing = scan(sub_board, [CID['מס חברה'], CID['מס סניף']], want_parent=True)
have = {(e['parent'], e['cv'].get(CID['מס חברה'], ''), e['cv'].get(CID['מס סניף'], '')) for e in existing}
print('branches board', len(br), 'existing subitems', len(existing), flush=True)

def vals(b):
    v = b['cv']
    d = {CID['מס חברה']: v[B['code']], CID['מס סניף']: v[B['branch_no']], CID['ח.פ']: v[B['hp']],
         CID['פעיל בתוכנה']: {'label': 'כן' if v[B['active']] == 'פעיל' else 'לא'},
         CID['רמה']: v[B['level']], CID['אפליקציה']: v[B['app']], CID['סולק']: v[B['credit']],
         CID['טלפון']: v[B['phone']], CID['מייל']: v[B['email']], CID['כתובת']: v[B['address']]}
    if v[B['move']] == 'כן': d[CID['MOVE']] = {'checked': 'true'}
    if v[B['members']].replace('.', '').isdigit(): d[CID['כמות מתאמנים']] = v[B['members']]
    if v[B['joined']]: d[CID['הצטרפות']] = {'date': v[B['joined']]}
    d = {k: x for k, x in d.items() if x not in (None, '')}
    return json.dumps(json.dumps(d, ensure_ascii=False), ensure_ascii=False)

todo = C.defaultdict(list); miss = 0
for sap, code, bno in MAP:
    pid = sap2id.get(sap); b = br.get((code, bno))
    if not pid or not b: miss += 1; continue
    if (pid, code, bno) not in have: todo[pid].append(b)
print('to create', sum(len(v) for v in todo.values()), 'under', len(todo), 'parents; unmatched', miss, flush=True)

STAT = C.Counter(); LOCK = threading.Lock()
def work(q):
    while True:
        try: pid, rows = q.get_nowait()
        except queue.Empty: return
        if time.time() - START > BUDGET:
            with LOCK: STAT['deferred'] += len(rows)
            continue
        muts = ['c%d: create_subitem(parent_item_id:%s,item_name:%s,column_values:%s){id}' % (i, pid, json.dumps(b['name'] or 'סניף', ensure_ascii=False), vals(b)) for i, b in enumerate(rows)]
        try:
            for i in range(0, len(muts), 10): gq('mutation{' + ''.join(muts[i:i + 10]) + '}')
            with LOCK: STAT['created'] += len(rows)
        except SystemExit: raise
        except Exception as e:
            with LOCK: STAT['errors'] += 1
            print('ERR', pid, str(e)[:120], flush=True)
Q = queue.Queue()
for pid, rows in todo.items(): Q.put((pid, rows))
th = [threading.Thread(target=work, args=(Q,), daemon=True) for _ in range(3)]
[t.start() for t in th]; [t.join() for t in th]
print('RESULT', dict(STAT), 'elapsed %.0fs' % (time.time() - START))
