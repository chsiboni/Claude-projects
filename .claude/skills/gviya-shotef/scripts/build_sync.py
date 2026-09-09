#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Weekly gviya sync builder — report XLSX + live board dump → APPLY plan (dry run, no writes).
usage: python3 build_sync.py REPORT.xlsx BOARD.json OUT.json
Rules (approved 26.08–09.09):
  zero → בוצע תשלום + 'ירד לאפס · לאישור' · returned after בוצע תשלום → לקוח בחוב + 'חזר ⚠'
  partial payment → מצב יתרה 'תשלום חלקי' only (stage untouched, since 07.09)
  strict per-invoice shotef, only between לקוח בחוב ↔ חוב בתחום שוטף · refresh every card
  המתנה לתשלום with promised date passed and no balance drop → פניה שניה + 'הבטחה לא קוימה'
"""
import sys, types, json, datetime, calendar, re, collections
sys.modules['mon'] = types.SimpleNamespace(q=None, paged=None, cv=None)
sys.path.insert(0, '/tmp/gviya/fizikal-gviya'); import sync

REP, BOARD, OUT = sys.argv[1:4]
name, rep, asof = sync.read_report(REP)
board = json.load(open(BOARD))
TODAY = datetime.date.fromisoformat(asof)
D = TODAY.strftime('%d.%m.%y')
LB, SH, PAID, NONE, WAIT, P2 = 'לקוח בחוב', 'חוב בתחום שוטף', 'בוצע תשלום', 'ללא חוב', 'המתנה לתשלום', 'פניה שניה'
COL = dict(debt='numeric_mm69dbtr', prev='numeric_mm692w0a', age='text_mm6970hn', agedays='numeric_mm694bkc',
           asof='date_mm69pvs6', bal='color_mm69xcy2', stage='color_mm69ws8c', log='long_text_mm69pjd4',
           code='text_mm69gg7g', hp='text_mm691gkk', terms='text_mm69ew8e', promised='date_mm7170br')

def eom(d): return datetime.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])
def n_of(t):
    t = (t or '').strip()
    if not t or ('הו' in t and 'ק' in t): return 0
    m = re.search(r'(\d+)', t); return int(m.group(1)) if m else 0
def overdue_terms(rec):
    n = n_of(rec['terms'])
    return sum(v for l, k, v in rec['buckets'] if v > 0 and (k != 'month' or TODAY > eom(l) + datetime.timedelta(days=n)))
def money(v): return format(int(round(v)), ',')

plan, stats = [], collections.Counter()
for code, rec in sorted(rep.items(), key=lambda kv: -kv[1]['bal']):
    nb = rec['bal']; age, days = sync.debt_age(rec, asof)
    over = overdue_terms(rec) if nb > 0 else 0
    b = board.get(code)
    if b is None:
        if nb <= 0: stats['skip-new-zero'] += 1; continue
        stage = SH if over == 0 else LB; bal = 'תחום שוטף' if stage == SH else 'יש חוב'
        cv = {COL['code']: code, COL['hp']: rec['hp'], COL['terms']: rec['terms'], COL['debt']: nb, COL['prev']: 0,
              COL['stage']: {'label': stage}, COL['bal']: {'label': bal}, COL['age']: age or '', COL['agedays']: days or '',
              COL['asof']: {'date': asof}, COL['log']: '%s: נוצר מדוח %s · חוב %s · %s' % (D, D[:5], money(nb), stage)}
        plan.append({'key': 'C-' + code, 'action': 'create', 'name': rec['name'], 'column_values': cv, 'why': 'לקוח חדש בחוב → ' + stage})
        stats['create → ' + stage] += 1; continue
    ob = float(b['debt'] or 0); stage = b['stage']; bal = b['bal']; newstage = stage; newbal = bal; log = None
    # --- balance rules
    if nb <= 0 and ob > 0 and stage not in (PAID, NONE):
        newstage, newbal, log = PAID, 'ירד לאפס · לאישור', '%s: ירד לאפס (היה %s) → בוצע תשלום' % (D, money(ob))
    elif nb > 0 and ob <= 0 and stage == PAID:
        newstage, newbal, log = LB, 'חזר ⚠', '%s: חוב חזר אחרי בוצע תשלום · %s → לקוח בחוב' % (D, money(nb))
    elif nb > 0 and ob <= 0 and stage == NONE:
        newstage, newbal, log = LB, 'יש חוב', '%s: חוב חדש %s → לקוח בחוב' % (D, money(nb))
    elif nb > 0 and 0 < nb < ob:
        newbal, log = 'תשלום חלקי', '%s: תשלום חלקי %s → %s (השלב לא שונה)' % (D, money(ob), money(nb))
    elif nb > 0 and nb > ob > 0:
        log = '%s: סכום עודכן %s → %s' % (D, money(ob), money(nb))
    # --- promised date rule (before shotef; only if still in WAIT)
    if newstage == WAIT and b.get('promised') and nb > 0 and nb >= ob:
        pd = datetime.date.fromisoformat(b['promised'])
        if pd < TODAY:
            newstage = P2; log = '%s: הבטחה לא קוימה — מועד %s עבר, היתרה לא ירדה (%s) → פניה שניה' % (D, pd.strftime('%d.%m'), money(nb))
    # --- shotef (strict, per invoice, only LB <-> SH)
    if nb > 0 and newstage in (LB, SH):
        if newstage == LB and over == 0:
            newstage, newbal = SH, 'תחום שוטף'; log = (log + ' · ' if log else D + ': ') + 'כל החוב בתחום השוטף → חוב בתחום שוטף'
        elif newstage == SH and over > 0:
            newstage, newbal = LB, 'יש חוב'; log = (log + ' · ' if log else D + ': ') + '₪%s עברו את מועד הפירעון → לקוח בחוב' % money(over)
        elif newstage == SH and newbal != 'תחום שוטף':
            newbal = 'תחום שוטף'
    cv = {COL['debt']: nb, COL['prev']: ob, COL['age']: age or '', COL['agedays']: days or '', COL['asof']: {'date': asof}}
    if newstage != stage: cv[COL['stage']] = {'label': newstage}
    if newbal != bal: cv[COL['bal']] = {'label': newbal}
    if log: cv[COL['log']] = log
    kind = 'M' if (newstage != stage or newbal != bal) else 'F'
    plan.append({'key': '%s-%s' % (kind, code), 'action': 'update', 'item_id': b['id'], 'name': b['name'], 'cur': stage, 'target': newstage,
                 'ob': ob, 'nb': nb, 'over': over, 'terms': rec['terms'], 'age': age, 'column_values': cv, 'why': log or 'רענון'})
    if newstage != stage: stats['%s → %s' % (stage, newstage)] += 1
    elif newbal != bal: stats['bal %s → %s' % (bal, newbal)] += 1
    else: stats['refresh'] += 1

missing = [(c, v['name'], v['stage'], v['debt']) for c, v in board.items() if c not in rep]
json.dump(plan, open(OUT, 'w'), ensure_ascii=False, indent=1)
print('report', name, asof, '| plan items', len(plan))
for k, v in sorted(stats.items(), key=lambda kv: -kv[1]): print('  %-45s %d' % (k, v))
print('board cards NOT in report:', len(missing))
for m in missing[:20]: print('   ', m)
print('\n== movers ==')
for p in plan:
    if p['action'] == 'create' or p['key'].startswith('M-'):
        if p['action'] == 'create': print('  NEW  %-32s %10s  %s' % (p['name'][:32], money(p['column_values'][COL['debt']]), p['why']))
        else: print('  %-7s %-32s %10s→%-10s %-16s→%-16s over=%-8s %s' % (p['key'][2:], p['name'][:32], money(p['ob']), money(p['nb']), p['cur'], p['target'], money(p['over']), p['terms']))
