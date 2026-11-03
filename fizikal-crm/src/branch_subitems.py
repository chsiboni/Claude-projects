#!/usr/bin/env python3
"""
Fizikal — הסניפים של הלקוחות המשלמים, כסאב-אייטמים לבורד "לקוחות".

    python3 src/branch_subitems.py --tag 2026-09

קלט:  out/sap_to_company_<tag>.json (מ-fresh_export) + data/moshe_2026-09/branches.csv
פלט:  out/branch_subitems_<tag>.json — שורה לכל סניף בתחום של לקוח: מספר סאפ של האב + פרטי הסניף.
       out/branch_leak_<tag>.json    — 🚩 סניפים פעילים בתוכנה שלא שייכים לשום לקוח משלם.

החלטת Chen (09/09/2026): לקוחות → סאב-אייטמים סניפים. בורד "סניפים" השטוח נשאר
כמראה של כל המערכת (1,679), עם עמודת "לקוח משלם".
"""

import argparse
import json
from collections import defaultdict

from fresh_export import DATA, load_branches
from revenue_pipeline import OUT

FLAGS = {"MOVE": "move", "ניהול מלאי": "stock", "קופה רושמת": "pos", "טפסים דיגטליים": "forms",
         "רכישה אינטרנטית": "web", "בקרת שערים": "gates", "יש עמדה ראשונה": "station"}


def row(r, sap=None):
    d = {
        "sap": sap, "code": r["code"], "branch_no": r["branch"],
        "name": r.get("שם הסניף") or r.get("שם חברה") or "", "company": r.get("שם חברה", ""),
        "hp": r["hp"], "active": r["active"], "level": r.get("רמה", ""),
        "business": r.get("סוג עסק", ""), "members": r.get("כמות פעילים", ""),
        "app": r.get("אפליקציה", ""), "credit": r.get("סולק אשראי", ""),
        "joined": r.get("ת.הצטרפות", ""), "last_receipt": r.get("קבלה אחרונה", ""),
        "phone": r.get("טלפון", ""), "email": r.get('דוא"ל', ""), "address": r.get("כתובת", ""),
    }
    for heb, key in FLAGS.items():
        d[key] = r.get(heb, "") == "V"
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", default="2026-09")
    a = ap.parse_args()

    linked = json.load(open(OUT / f"sap_to_company_{a.tag}.json", encoding="utf-8"))
    branches = load_branches(DATA / "branches.csv")
    by_code = defaultdict(list)
    for b in branches:
        by_code[b["code"]].append(b)

    subs, claimed = [], set()
    for sap, rec in linked.items():
        for code in rec["codes"]:
            scope = rec.get("scope", {}).get(code)         # רשימת סניפים, או ריק = כל החברה
            for b in by_code.get(code, []):
                if scope and b["branch"] not in scope:
                    continue
                subs.append(row(b, sap))
                claimed.add((code, b["branch"]))

    leak = [row(b) for b in branches if b["active"] and (b["code"], b["branch"]) not in claimed]

    json.dump(subs, open(OUT / f"branch_subitems_{a.tag}.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump(leak, open(OUT / f"branch_leak_{a.tag}.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    act = sum(1 for s in subs if s["active"])
    print(f"סניפים של לקוחות: {len(subs)} ({act} פעילים, {len(subs)-act} לא) תחת {len({s['sap'] for s in subs})} לקוחות")
    print(f"🚩 פעילים בתוכנה בלי לקוח משלם: {len(leak)} סניפים ב-{len({b['code'] for b in leak})} חברות")


if __name__ == "__main__":
    main()
