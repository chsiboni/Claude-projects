#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gviya-shotef · עובר על לוח הגבייה ובודק מי בתחום השוטף.

הכל מהלוח עצמו — בלי קובץ גיול. הסנכרון כבר מטביע בכל כרטיס את
`משך חוב` (החודש הישן ביותר שיש בו כסף) ואת `תנאי תשלום`, וזה כל
מה שצריך: חוב הוא "בתחום שוטף" אם מועד הפירעון של החשבונית הישנה
ביותר (סוף החודש שלה + N ימי שוטף) עוד לא עבר.

שימוש:
  python3 classify_shotef.py --board board_full.json [--out shotef_apply.json] [--today YYYY-MM-DD]

board_full.json — מיפוי {קוד: {"id","name","stage","debt","terms","age","report"}}
שנשלף מהלוח דרך ה-MCP (העמודות: text_mm69gg7g, color_mm69ws8c, numeric_mm69dbtr,
text_mm69ew8e, text_mm6970hn, date_mm69pvs6).

הסקריפט לא כותב למאנדיי לעולם. הוא רק מחשב ומדפיס, וכותב תוכנית ביצוע ל---out.
"""
import argparse, calendar, datetime, json, re, sys

STAGE_TARGET, STAGE_TARGET_INDEX = "חוב בתחום שוטף", 10

def parse_terms(t):
    t = (t or "").strip()
    if not t: return 0, "ריק ⚠"
    if "הו" in t and "ק" in t: return 0, 'הו"ק'
    m = re.search(r"(\d+)", t)
    return (int(m.group(1)), t) if m else (0, t + " ⚠")

def oldest_month(age):
    """'07/26' → date. 'דצמבר (שנה קודמת)' / 'שנים קודמות' → None (באיחור בהגדרה)."""
    m = re.fullmatch(r"(\d{1,2})/(\d{2})", (age or "").strip())
    if not m: return None
    return datetime.date(2000 + int(m.group(2)), int(m.group(1)), 1)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--board", required=True)
    ap.add_argument("--out", default="shotef_apply.json")
    ap.add_argument("--today", help="לתיקוף בלבד; ברירת מחדל היום")
    a = ap.parse_args()
    today = datetime.date.fromisoformat(a.today) if a.today else datetime.date.today()
    board = json.load(open(a.board))

    clean, already, overdue_in_stage = [], [], []
    for code, b in sorted(board.items(), key=lambda kv: -float(kv[1]["debt"] or 0)):
        debt = float(b["debt"] or 0)
        if debt <= 0: continue
        n, terms = parse_terms(b["terms"])
        mo = oldest_month(b["age"])
        within = False; due = None
        if mo:
            due = datetime.date(mo.year, mo.month,
                  calendar.monthrange(mo.year, mo.month)[1]) + datetime.timedelta(days=n)
            within = today <= due
        row = {"code": code, "item_id": b["id"], "name": b["name"], "bal": round(debt),
               "terms": terms, "age": b["age"], "stage": b["stage"], "report": b["report"],
               "due": due.strftime("%d.%m.%y") if due else "",
               "days_left": (due - today).days if due else None}
        if b["stage"] == STAGE_TARGET:
            (already if within else overdue_in_stage).append(row)
        elif within:
            clean.append(row)

    plan = []
    for r in clean:
        upd = ["📅 **הועבר ל\"חוב בתחום שוטף\"** — החוב בתוך תנאי התשלום",
               "חוב: ₪%s · תנאים: %s · חשבונית ישנה ביותר: %s" % (format(r["bal"], ","), r["terms"], r["age"]),
               "מועד פירעון: %s (בעוד %d ימים)" % (r["due"], r["days_left"]),
               "", "_חישוב תחום-שוטף מנתוני הלוח (עודכן מדוח %s) · הופק %s_"
               % (r["report"] or "?", today.strftime("%d.%m.%y"))]
        plan.append({"code": r["code"], "item_id": r["item_id"], "name": r["name"],
                     "cur": r["stage"], "target": STAGE_TARGET,
                     "target_index": STAGE_TARGET_INDEX, "update": "\n".join(upd)})
    json.dump(plan, open(a.out, "w"), ensure_ascii=False, indent=1)

    money = lambda v: "₪" + format(int(v), ",")
    print("=" * 74)
    print("תחום שוטף · מנתוני הלוח · נכון ל-%s" % today.strftime("%d.%m.%y"))
    print("=" * 74)
    print("\n✅ בתחום השוטף ולא בשלב — מועמדים להעברה: %d · %s"
          % (len(clean), money(sum(r["bal"] for r in clean))))
    for r in clean:
        print("   %-7s %-34s %10s  %-9s ישן: %-6s פירעון %s (%d י׳)  [%s]"
              % (r["code"], r["name"][:34], money(r["bal"]), r["terms"], r["age"], r["due"], r["days_left"], r["stage"]))
    print("\n✔ כבר בשלב ועדיין בשוטף: %d · %s" % (len(already), money(sum(r["bal"] for r in already))))
    for r in already:
        print("   %-7s %-34s %10s  פירעון %s" % (r["code"], r["name"][:34], money(r["bal"]), r["due"]))
    if overdue_in_stage:
        print("\n⏰ בשלב 'חוב בתחום שוטף' אבל החוב כבר עבר את הפירעון — להחזיר לטיפול? (החלטת חן/נציגה):")
        for r in overdue_in_stage:
            print("   %-7s %-34s %10s  %-9s ישן: %-6s פירעון עבר: %s"
                  % (r["code"], r["name"][:34], money(r["bal"]), r["terms"], r["age"], r["due"] or "—"))
    print("\nתוכנית ביצוע (%d העברות) נכתבה ל-%s — לא בוצע דבר במאנדיי." % (len(plan), a.out))
    print("⚠ תזכורת §3: חוב שהופק מחדש (הו\"ק שחזרה, זיכוי+חשבונית חדשה) נראה 'שוטף'")
    print("  בנתוני הגיול. אם לקוח ברשימה מוכר כחוזר כרוני — לעצור ולשאול את חן.")

if __name__ == "__main__":
    main()
