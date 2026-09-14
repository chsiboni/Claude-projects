#!/usr/bin/env python3
"""
Fizikal — רשימת מתנות ההשקה: אדם אחד לכל סניף של לקוח פעיל.

    python3 src/gift_list.py --tag 2026-09

החלטות Chen (14/09/2026): מתנה לכל מנהל סניף (לא ללקוח) · כולל לקוחות "בסיכון" ו"חדש" · כל המתנות זהות ·
כל הדרגות כולל VIP/גדול. "לא פעיל" (3+ חודשים בלי תשלום) — לא.

מי מקבל בסניף: ההרשאה הגבוהה ביותר (הרשאת על → ניהול חברה → מנהל → מנהל משמרת); בתיקו — הותיק ביותר.
אדם שמנהל כמה סניפים של אותו לקוח מקבל מתנה אחת (על הסניף הראשון), והסניפים האחרים מקבלים את הבא בתור.
סניף בלי אף איש קשר → שורה עם "אין איש קשר", כדי שמישהו ישלים ידנית.

קלט: out/branch_subitems_<tag>.json · out/sap_to_company_<tag>.json · data/moshe_2026-09/employees.csv
פלט: out/Fizikal_מתנות_השקה_<tag>.xlsx (גיליון אחד, שורה לסניף)
"""

import argparse
import json
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from fresh_export import DATA, RANK, load_employees
from revenue_pipeline import OUT

HEAD = ["לקוח", "דרגה", "סטטוס", "מס חברה", "מס סניף", "סניף", "מקבל המתנה", "תפקיד", "נייד", "מייל", "כתובת למשלוח", "הערה"]


def joined(e):
    try:
        return datetime.strptime(e.get("תאריך תחילת עבודה", ""), "%d/%m/%Y")
    except ValueError:
        return datetime.max


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", default="2026-09")
    a = ap.parse_args()

    branches = json.load(open(OUT / f"branch_subitems_{a.tag}.json", encoding="utf-8"))
    clients = json.load(open(OUT / f"sap_to_company_{a.tag}.json", encoding="utf-8"))
    by_key = defaultdict(list)
    for e in load_employees(DATA / "employees.csv"):
        by_key[(e["code"], e["branch"])].append(e)

    rows, taken = [], set()                       # taken: (sap, code, emp) — אדם אחד = מתנה אחת אצל אותו לקוח
    for b in sorted(branches, key=lambda x: (clients[x["sap"]]["name"], x["code"], int(x["branch_no"] or 0))):
        c = clients[b["sap"]]
        if c["status"] == "לא פעיל":
            continue
        pool = sorted(by_key[(b["code"], b["branch_no"])], key=lambda e: (-RANK.get(e["role"], 0), joined(e)))
        pick = next((e for e in pool if (b["sap"], e["code"], e["emp"]) not in taken), None)
        note = ""
        if pick:
            taken.add((b["sap"], pick["code"], pick["emp"]))
            if not pick["phone"]:
                note = "אין נייד"
        else:
            note = "אין איש קשר בסניף" if not pool else "כל אנשי הקשר כבר מקבלים בסניף אחר"
        if not b["address"]:
            note = (note + " · " if note else "") + "אין כתובת"
        if not b["active"]:
            note = (note + " · " if note else "") + "סניף מסומן לא פעיל בתוכנה"
        rows.append([c["name"], c["tier"], c["status"], b["code"], b["branch_no"], b["name"],
                     pick["name"] if pick else "", pick["role"] if pick else "",
                     pick["phone"] if pick else "", pick["email"] if pick else "", b["address"], note])

    # לקוחות פעילים שלא קושרו למערכת Fizikal (הבירור): אין סניף/מקבל — שורה אחת ללקוח, עם טלפון/מייל מבורד הגביה אם יש
    gviya = {r[1]: r for r in json.load(open(OUT / f"collections_{a.tag}.json", encoding="utf-8"))}   # [name, sap, ?, phone, email]
    all_clients = json.load(open(OUT / f"clients_status_2026-07.json", encoding="utf-8"))
    n_linked = len(rows)
    for c in sorted(all_clients, key=lambda c: c["name"]):
        if c["sap"] in clients or c["status"] == "לא פעיל":
            continue
        g = gviya.get(c["sap"], ["", "", "", "", ""])
        rows.append([c["name"], c["tier"], c["status"], "", "", "", "", "", g[3], g[4], "",
                     "לא מקושר למערכת Fizikal — להשלים מקבל וכתובת" + ("" if g[3] or g[4] else " · אין פרטי קשר גם בגביה")])

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "מתנות"
    ws.sheet_view.rightToLeft = True
    ws.append(HEAD)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append(r)
    for i, w in enumerate([28, 8, 8, 8, 7, 26, 22, 12, 13, 28, 34, 30], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    out = OUT / f"Fizikal_מתנות_השקה_{a.tag}.xlsx"
    wb.save(out)

    n = len(rows); ok = sum(1 for r in rows if r[6] and r[8] and r[10])
    print(f"{out.name}: {n} שורות = {n_linked} סניפים + {n - n_linked} לקוחות לא מקושרים · {ok} מוכנים למשלוח (יש מקבל + נייד + כתובת)")
    print(f"  לא מקושרים עם טלפון/מייל מהגביה: {sum(1 for r in rows[n_linked:] if r[8] or r[9])}")
    print(f"  בלי איש קשר: {sum(1 for r in rows if not r[6])} · בלי נייד: {sum(1 for r in rows if r[6] and not r[8])} · בלי כתובת: {sum(1 for r in rows if not r[10])}")
    print(f"  לקוחות: {len({r[0] for r in rows})} · לפי דרגה: " + ", ".join(f"{t} {sum(1 for r in rows if r[1]==t)}" for t in ("VIP", "גדול", "בינוני", "קטן", "מוב")))


if __name__ == "__main__":
    main()
