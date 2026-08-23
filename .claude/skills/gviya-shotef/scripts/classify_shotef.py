#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gviya-shotef · מסווג "חוב בתחום שוטף" ללוח הגבייה של פיזיקל.

קורא את דוח הגיול (XLSX של הכספים), מצליב מול היסטוריית הסנאפשוטים וקובץ
החזרי ההו"ק (אם ניתן), ומוציא שלוש רשימות:

  clean    — כל החוב בתוך חלון תנאי התשלום. מועמדים להעברה.
  suspect  — נראים בתחום השוטף אבל יש דפוס חוב-ממוחזר. דורשים החלטה של חן.
  already  — כבר בשלב "חוב בתחום שוטף" בלוח (אימות, לא פעולה).

שימוש:
  python3 classify_shotef.py --xlsx <גיול.xlsx> --board <board_dump.json> \
      [--sheet 16.08.26] [--history <תיקיית סנאפשוטים>] [--returns <הוק-חוזרות.xlsx>] \
      [--out shotef_apply.json]

board_dump.json — מיפוי {קוד לקוח: {"id","name","stage","debt"}} שנשלף מהלוח דרך ה-MCP.
הפלט --out הוא קובץ APPLY בפורמט שסוכן-הרקע יודע לבצע (item_id, target_index, update).

הסקריפט לא כותב למאנדיי לעולם. הוא רק מחשב.
"""
import argparse, calendar, collections, datetime, glob, json, os, re, sys
import openpyxl

MONTHS = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני","יולי","אוגוסט",
          "ספטמבר","אוקטובר","נובמבר","דצמבר"]
STAGE_TARGET, STAGE_TARGET_INDEX = "חוב בתחום שוטף", 10

# --------------------------------------------------------------- aging report
def read_report(path, sheet=None):
    """קריאה לפי שם כותרת בלבד — הפריסה נודדת בין שבועות (ראה HANDOFF §5)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    name = sheet or wb.sheetnames[-1]
    if name not in wb.sheetnames:
        sys.exit("אין גיליון %r. קיימים: %s" % (name, ", ".join(wb.sheetnames)))
    rows = list(wb[name].iter_rows(values_only=True))
    hdr = next(i for i, r in enumerate(rows[:8])
               if any(str(c).strip() == "שם לקוח" for c in r if c))
    cols = [str(c).strip() if c is not None else "" for c in rows[hdr]]
    ci = {n: j for j, n in enumerate(cols) if n}
    bal_j = ci.get("יתרה לתשלום", ci.get(" יתרה לתשלום "))
    if bal_j is None: sys.exit("לא נמצאה עמודת 'יתרה לתשלום' בגיליון %s" % name)
    aging = []
    for j, n in enumerate(cols):
        raw = rows[hdr][j]
        if isinstance(raw, datetime.datetime): aging.append((j, raw.date(), "month"))
        elif n in MONTHS:                      aging.append((j, n, "lastyear"))
        elif n == "שנים קודמות":                aging.append((j, n, "ancient"))
    out = {}
    for r in rows[hdr + 1:]:
        code = r[ci.get("קוד לקוח", -1)]
        if code is None or not str(code).strip().isdigit(): continue
        try: bal = float(r[bal_j]) if r[bal_j] not in (None, "") else 0.0
        except (TypeError, ValueError): bal = 0.0
        buckets = []
        for j, label, kind in aging:
            try: v = float(r[j]) if r[j] not in (None, "") else 0.0
            except (TypeError, ValueError): v = 0.0
            if v: buckets.append((label, kind, round(v)))
        out[str(code).strip()] = {
            "name": str(r[ci["שם לקוח"]] or "").strip(),
            "terms": str(r[ci.get("שוטף", -1)] or "").strip() if "שוטף" in ci else "",
            "bal": round(bal), "buckets": buckets}
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})", name.strip())
    asof = (datetime.date(int(m.group(3)) + (2000 if int(m.group(3)) < 100 else 0),
                          int(m.group(2)), int(m.group(1))).isoformat()
            if m else datetime.date.today().isoformat())
    return name, out, asof

# ----------------------------------------------------------------- terms math
def parse_terms(t):
    t = (t or "").strip()
    if not t: return 0, "ריק ⚠"
    if "הו" in t and "ק" in t: return 0, 'הו"ק'
    m = re.search(r"(\d+)", t)
    return (int(m.group(1)), t) if m else (0, t + " ⚠")

def within_terms(rec, asof):
    """True אם כל שקל חוב יושב בחודש שמועד הפירעון שלו (סוף חודש + N ימי שוטף)
       עוד לא עבר ביום הדוח. עמודות שנה-קודמת ושנים-קודמות תמיד באיחור."""
    today = datetime.date.fromisoformat(asof)
    n, label = parse_terms(rec["terms"])
    oldest = None
    for lbl, kind, v in rec["buckets"]:
        if v <= 0: continue
        if kind != "month": return False, label, None
        due = datetime.date(lbl.year, lbl.month,
              calendar.monthrange(lbl.year, lbl.month)[1]) + datetime.timedelta(days=n)
        if today > due: return False, label, None
        if oldest is None or lbl < oldest[0]: oldest = (lbl, due)
    return oldest is not None, label, oldest

# --------------------------------------------- §3 guards: recycled-debt traps
def load_history(hdir):
    hist = {}
    for f in sorted(glob.glob(os.path.join(hdir, "*.json"))):
        d = json.load(open(f)); hist[d["date"]] = d["clients"]
    return hist

def recycled_flags(code, hist):
    """שני דפוסים של חוב ממוחזר שנראה 'שוטף' בגיול:
       1. ירידה לאפס וחזרה באותו סכום (זיכוי + הפקה מחדש — בוטיק פיטנס, כפ"ס).
       2. הפיגור ירד בלי שהיתרה ירדה (חשבונית נדדה לחודש הנוכחי — בת ציון).
       זה בדיוק ה-clamp של קוביית התשלומים, בשימוש הפוך."""
    dates = sorted(hist)
    rows = [hist[d].get(code) for d in dates]
    flags = []
    pos = [(i, r[0]) for i, r in enumerate(rows) if r and r[0] > 0]
    for (i, a), (j, b) in zip(pos, pos[1:]):
        if j - i > 1 and all((rows[k] or [0])[0] <= 0 for k in range(i + 1, j)) and a == b:
            flags.append("ירד לאפס וחזר באותו סכום (₪%s, %s→%s)" % (format(a, ","), dates[i], dates[j]))
    for k in range(len(rows) - 1):
        a, b = rows[k], rows[k + 1]
        if a and b and len(a) > 1 and len(b) > 1 and a[1] - b[1] > 0 and a[0] - b[0] <= 0:
            flags.append("הפיגור ירד בלי שהיתרה ירדה (%s→%s) — חשבונית הופקה מחדש" % (dates[k], dates[k + 1]))
    return flags

def returns_codes(path):
    """קורא את קובץ החזרי ההו"ק של יסמין ומחזיר את קודי הלקוח שמופיעים בו.
       גשר: לומדים חשבון-בנק→קוד מהשורות שיש בהן קוד 20xxxx (אפס התנגשויות, נבדק),
       ומחילים על שורות בלי קוד. שמות בקובץ משובשים — לעולם לא מפתח."""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))
    S = lambda c: "" if c is None else str(c).strip()
    acct2code, accts = {}, set()
    for r in rows:
        cells = [S(c) for c in r]
        code = next((c for c in cells[:2] if re.fullmatch(r"20\d{4}", c)), "")
        acct = next((cells[k + 1] for k, c in enumerate(cells)
                     if c == "ILS" and k + 1 < len(cells) and cells[k + 1].isdigit()), "")
        if not acct:
            digs = [c for c in cells[4:8] if c.isdigit() and len(c) >= 4]
            acct = digs[0] if digs else ""
        if acct:
            accts.add(acct)
            if code: acct2code[acct] = code
    return {acct2code[a] for a in accts if a in acct2code}

# ------------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True); ap.add_argument("--sheet")
    ap.add_argument("--board", required=True)
    ap.add_argument("--history"); ap.add_argument("--returns")
    ap.add_argument("--out", default="shotef_apply.json")
    a = ap.parse_args()

    sheet, rep, asof = read_report(a.xlsx, a.sheet)
    board = json.load(open(a.board))
    hist = load_history(a.history) if a.history else {}
    ret = returns_codes(a.returns) if a.returns else set()
    today = datetime.date.fromisoformat(asof)

    clean, suspect, already = [], [], []
    for code, r in sorted(rep.items(), key=lambda kv: -kv[1]["bal"]):
        if r["bal"] <= 0: continue
        ok, terms, oldest = within_terms(r, asof)
        if not ok: continue
        b = board.get(code, {})
        row = {"code": code, "name": r["name"], "bal": r["bal"], "terms": terms,
               "stage": b.get("stage", "— לא בלוח —"), "item_id": b.get("id"),
               "oldest": oldest[0].strftime("%m/%Y") if oldest else "",
               "due": oldest[1].strftime("%d.%m.%y") if oldest else "",
               "days_left": (oldest[1] - today).days if oldest else 0, "flags": []}
        if code in ret: row["flags"].append('בקובץ החזרי ההו"ק — ההו"ק חזרה והחוב הופק מחדש')
        if hist: row["flags"] += recycled_flags(code, hist)
        if b.get("stage") == STAGE_TARGET: already.append(row)
        elif row["flags"]: suspect.append(row)
        else: clean.append(row)

    plan = []
    for row in clean:
        if not row["item_id"]: continue
        upd = ["📅 **הועבר ל\"חוב בתחום שוטף\"** — כל החוב בתוך תנאי התשלום",
               "חוב: ₪%s · תנאים: %s" % (format(row["bal"], ","), row["terms"]),
               "החשבונית הישנה ביותר: %s → פירעון %s (בעוד %d ימים מיום הדוח)"
               % (row["oldest"], row["due"], row["days_left"]),
               "נבדק מול היסטוריית הסנאפשוטים%s — לא זוהה דפוס חוב ממוחזר"
               % (' וקובץ החזרי ההו"ק' if ret else ""),
               "", "_חישוב תחום-שוטף · דוח גיול %s · הופק %s_"
               % (sheet, datetime.date.today().strftime("%d.%m.%y"))]
        plan.append({"code": row["code"], "item_id": row["item_id"], "name": row["name"],
                     "cur": row["stage"], "target": STAGE_TARGET,
                     "target_index": STAGE_TARGET_INDEX, "update": "\n".join(upd)})
    json.dump(plan, open(a.out, "w"), ensure_ascii=False, indent=1)

    money = lambda n: "₪" + format(int(n), ",")
    print("=" * 74)
    print("תחום שוטף · גיליון %s (תאריך דוח %s)" % (sheet, asof))
    print("=" * 74)
    print("\n✅ נקיים — מועמדים להעברה: %d · %s" % (len(clean), money(sum(r['bal'] for r in clean))))
    for r in clean:
        print("   %-7s %-34s %10s  %-9s פירעון %s (%d י׳)  [%s]"
              % (r["code"], r["name"][:34], money(r["bal"]), r["terms"], r["due"], r["days_left"], r["stage"]))
    print("\n🚫 חשודים כחוב ממוחזר — דורשים החלטה: %d · %s" % (len(suspect), money(sum(r['bal'] for r in suspect))))
    for r in suspect:
        print("   %-7s %-34s %10s  → %s" % (r["code"], r["name"][:34], money(r["bal"]), "; ".join(r["flags"])))
    print("\n✔ כבר בשלב הנכון: %d · %s" % (len(already), money(sum(r['bal'] for r in already))))
    for r in already:
        print("   %-7s %-34s %10s" % (r["code"], r["name"][:34], money(r["bal"])))
    print("\nתוכנית ביצוע (%d העברות) נכתבה ל-%s — לא בוצע דבר במאנדיי." % (len(plan), a.out))
    if not hist: print("⚠ רץ בלי --history — בדיקת חוב-ממוחזר לפי סנאפשוטים לא בוצעה.")
    if not ret:  print('⚠ רץ בלי --returns — הצלבה מול קובץ החזרי ההו"ק לא בוצעה.')

if __name__ == "__main__":
    main()
