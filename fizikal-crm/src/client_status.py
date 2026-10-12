#!/usr/bin/env python3
"""
Fizikal — סטטוס פעיל/לא-פעיל של הלקוחות.

מרחיב את צינור ההכנסות מחודש בודד להיסטוריה: לכל לקוח הריטיינר בכל חודש בדוח,
וממנו נגזרים הסטטוס, כמה חודשים הוא לא שילם, ומתי שילם לאחרונה.

    python3 src/client_status.py --file 'data/הכנסות יולי.xlsx' --month 2026-07

למה זה סקריפט נפרד ולא הרחבה של revenue_pipeline: הצינור ההוא עונה על
"מי הלקוחות שלנו החודש" וזורק ריטיינר 0. כאן דווקא הלקוחות האלה הם העיקר —
מי שהפסיק לשלם הוא מה שאנחנו מחפשים.
"""

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

from revenue_pipeline import (CATS, OUT, ROOT, canonical_name, category_of, digits,
                             family_key, load_config, open_sheet, resolve_hp_by_sap, s,
                             num, tier_of)

MONTH_HE = {1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל", 5: "מאי", 6: "יוני",
            7: "יולי", 8: "אוגוסט", 9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר"}


def month_label(key):
    year, month = key.split("-")
    return f"{MONTH_HE[int(month)]} {year}"


def build_history(path, month, rev, nets, inactive_after):
    """
    לכל לקוח: הריטיינר בכל חודש בדוח, ופירוט סעיפי ההכנסה של חודש היעד.

    נכלל כל מי שהיה לו ריטיינר באיזה שהוא חודש — לא רק בחודש היעד. לקוח שהפסיק
    לשלם צריך להישאר בבורד עם סטטוס, ולא להיעלם ממנו.
    """
    layout = rev["layout"]
    months = [m for m in sorted(layout["month_columns"]) if m <= month]
    if month not in layout["month_columns"]:
        raise SystemExit(f"החודש {month} לא מוגדר בקונפיג. זמינים: {', '.join(layout['month_columns'])}")

    sap_re = re.compile(rev["client_sap_pattern"])
    skip_names, skip_saps = set(rev["skip_names"]), set(rev["skip_saps"])
    skip_details = set(rev["skip_details"])

    ws = open_sheet(path)
    rows = [r for r in ws.iter_rows(min_row=layout["first_data_row"], values_only=True) if any(r)]
    hp_by_sap, _, _ = resolve_hp_by_sap(rows, layout)

    fam = defaultdict(lambda: {
        "hp": "", "saps": set(), "names": Counter(),
        "ret": dict.fromkeys(months, 0.0),          # ריטיינר לכל חודש
        "cat": dict.fromkeys(CATS, 0.0),            # סעיפי חודש היעד
    })

    for row in rows:
        sap, name, detail = s(row[layout["col_sap"]]), s(row[layout["col_name"]]), s(row[layout["col_detail"]])
        if not sap_re.fullmatch(sap) or sap in skip_saps or not name:
            continue
        if name in skip_names or any(name.startswith(p) for p in rev["skip_name_prefixes"]):
            continue
        if detail in skip_details:
            continue

        hp = hp_by_sap.get(sap, "")
        kind, key = family_key(name, hp, sap, nets)
        f = fam[(kind, key)]
        f["saps"].add(sap)
        if hp and hp != "0" and not f["hp"]:
            f["hp"] = hp
        if kind == "net":
            f["names"][key] += 10_000      # שם הרשת תמיד מנצח
        else:
            f["names"][name] += 1

        cat = category_of(detail, rev["categories"], rev["retainer_categories_exclude"])
        if cat == "ret":
            for m in months:
                f["ret"][m] += num(row[layout["month_columns"][m]])
        f["cat"][cat] += num(row[layout["month_columns"][month]])

    clients = []
    for f in fam.values():
        paid = [m for m in months if round(f["ret"][m]) >= 1]
        if not paid:
            continue                       # מעולם לא שילם ריטיינר — לא לקוח

        last_paid = paid[-1]
        gap = len(months) - months.index(last_paid) - 1     # חודשים רצופים בלי ריטיינר
        first_paid = paid[0]

        if gap >= inactive_after:
            status = "לא פעיל"
        elif gap > 0:
            # הפסיק לשלם אבל עוד לא חלף מפתח הסבלנות. זו ההתראה המוקדמת —
            # בלי הסטטוס הזה נטישה גדולה מתחבאת בין הפעילים עד שיהיה מאוחר.
            status = "בסיכון"
        elif first_paid == month and len(paid) == 1:
            status = "חדש"
        else:
            status = "פעיל"

        # Tier נגזר מהחודש האחרון ששולם, כדי שלקוח לא-פעיל ישמור על הסיווג שהיה לו
        tier = tier_of(f["ret"][last_paid], rev["tiers"])
        cur, prev = round(f["ret"][month]), round(f["ret"][months[-2]]) if len(months) > 1 else 0
        c = f["cat"]
        clients.append({
            "name": canonical_name(f["names"]), "hp": f["hp"], "sap": min(f["saps"]),
            "nsap": len(f["saps"]), "tier": tier, "status": status,
            "ret": cur, "prev": prev, "change": cur - prev,
            "gap": gap, "last_paid": month_label(last_paid), "month": month_label(month),
            "ret_last": round(f["ret"][last_paid]),
            "mod": round(c["mod"]), "sms": round(c["sms"] + c["sms1"]),
            "train": round(c["train"]), "dev": round(c["dev"]), "hw": round(c["hw"]),
            "tags": round(c["tags"]),
            "other": round(c["tech"] + c["ship"] + c["other"]),
            "total": round(sum(c.values())),
        })

    # שם שחוזר בשני גופים משפטיים שונים מקבל את המזהה שלו, אחרת השורות נראות ככפילות
    seen = Counter(c["name"] for c in clients)
    for c in clients:
        if seen[c["name"]] > 1:
            c["name"] = f"{c['name']} · {'ח.פ ' + c['hp'] if c['hp'] else 'סאפ ' + c['sap']}"

    order = {"בסיכון": 0, "לא פעיל": 1, "חדש": 2, "פעיל": 3}
    clients.sort(key=lambda c: (order[c["status"]], -max(c["ret"], c["ret_last"])))
    return clients, months


COLUMNS = [
    ("Name", "name"), ("סטטוס", "status"), ("Tier", "tier"),
    ("מספר סאפ", "sap"), ("ח.פ", "hp"), ("חודש הדוח", "month"),
    ("ריטיינר", "ret"), ("ריטיינר קודם", "prev"), ("שינוי", "change"),
    ("חודשים ללא ריטיינר", "gap"), ("שילם לאחרונה", "last_paid"),
    ("ריטיינר בתשלום האחרון", "ret_last"),
    ('סה"כ החודש', "total"), ("מודולים", "mod"), ("SMS", "sms"),
    ("הדרכות והתקנה", "train"), ("פיתוח", "dev"), ("חומרה", "hw"),
    ("תגי קרבה", "tags"), ("אחר", "other"), ("מספר כרטיסי חיוב", "nsap"),
]


def write_outputs(clients, month):
    OUT.mkdir(exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "לקוחות"
    ws.sheet_view.rightToLeft = True
    ws.append([t for t, _ in COLUMNS])       # Name ראשונה — אחרת monday מפילה כל שורה
    for c in clients:
        ws.append([c[k] for _, k in COLUMNS])
    for col, width in zip("ABCDEFGHIJKLMNOPQRSTU",
                          [34, 11, 9, 11, 12, 12, 11, 12, 10, 10, 13, 15, 12, 11, 10, 13, 10, 10, 10, 10, 10]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    xlsx = OUT / f"clients_status_{month}.xlsx"
    wb.save(xlsx)

    js = OUT / f"clients_status_{month}.json"
    js.write_text(json.dumps(clients, ensure_ascii=False, indent=1), encoding="utf-8")
    return xlsx, js


def report(clients, months, month, inactive_after):
    by_status = defaultdict(list)
    for c in clients:
        by_status[c["status"]].append(c)
    print(f"חודשים בדוח: {', '.join(months)}")
    print(f"כלל ההחרגה: {inactive_after} חודשים רצופים ללא ריטיינר → לא פעיל\n")
    print(f"סה\"כ לקוחות בבורד: {len(clients)}")
    for st in ("פעיל", "חדש", "בסיכון", "לא פעיל"):
        g = by_status[st]
        at_risk = sum(c["ret_last"] for c in g) if st in ("בסיכון", "לא פעיל") else sum(c["ret"] for c in g)
        label = "בסכנה" if st in ("בסיכון", "לא פעיל") else "ריטיינר"
        print(f"  {st:<9} {len(g):>4}   {label} ₪{at_risk:>9,}")

    tiers = defaultdict(lambda: [0, 0])
    for c in clients:
        tiers[c["tier"]][1 if c["status"] == "לא פעיל" else 0] += 1
    print("\nלפי Tier (פעיל / לא פעיל):")
    for t in ("VIP", "גדול", "בינוני", "קטן", "מוב"):
        a, i = tiers[t]
        print(f"  {t:<8} {a:>4} / {i:<4}")

    risk = sorted(by_status["בסיכון"], key=lambda c: -c["ret_last"])
    print(f"\n⚠️  בסיכון ({len(risk)}) — הפסיקו החודש, עוד לא חלף מפתח הסבלנות:")
    for c in risk:
        print(f"  {c['tier']:<7} ₪{c['ret_last']:>7,}  שילם עד {c['last_paid']:<12} {c['name'][:36]}")

    inactive = sorted(by_status["לא פעיל"], key=lambda c: -c["ret_last"])
    print(f"\nלא פעילים ({len(inactive)}) — לפי הריטיינר שהיה להם בתשלום האחרון:")
    for c in inactive[:12]:
        print(f"  {c['tier']:<7} ₪{c['ret_last']:>7,}  {c['gap']} חודשים  שילם עד {c['last_paid']:<12} {c['name'][:34]}")
    if len(inactive) > 12:
        print(f"  ... ועוד {len(inactive)-12}")

    movers = sorted((c for c in clients if abs(c["change"]) >= 1000), key=lambda c: -abs(c["change"]))
    print(f"\nתנועות ריטיינר מעל ₪1,000 ({len(movers)}):")
    for c in movers[:10]:
        print(f"  {c['change']:>+8,}  ₪{c['prev']:>7,} → ₪{c['ret']:>7,}  {c['name'][:34]}")


def main():
    p = argparse.ArgumentParser(description="סטטוס פעיל/לא-פעיל של לקוחות Fizikal")
    p.add_argument("--file", required=True, help="נתיב לדוח ההכנסות")
    p.add_argument("--month", required=True, help="חודש היעד, למשל 2026-07")
    p.add_argument("--inactive-after", type=int, default=2,
                   help="חודשים רצופים ללא ריטיינר עד סימון 'לא פעיל' (ברירת מחדל 2)")
    args = p.parse_args()

    rev, nets = load_config()
    clients, months = build_history(Path(args.file), args.month, rev, nets, args.inactive_after)
    report(clients, months, args.month, args.inactive_after)
    xlsx, js = write_outputs(clients, args.month)
    print(f"\nנכתב:\n  {xlsx.relative_to(ROOT)}  ← לייבוא ל-monday\n  {js.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
