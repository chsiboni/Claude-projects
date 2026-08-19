#!/usr/bin/env python3
"""
Fizikal — קובץ אנשי קשר לשיווק.

מייצר אקסל מסודר עם הניידים והמיילים של אנשי הקשר אצל הלקוחות הפעילים,
מדורג לפי Tier, מנוקה מכפילויות ומוכן להזנה למערכת דיוור או SMS.

    python3 src/marketing_contacts.py --month 2026-07

שתי בעיות בנתוני המקור שהקובץ הזה פותר:
  · האפס המוביל של הנייד נאכל בייצוא ('527186196' במקום '0527186196') — 2,324 שורות.
  · אותו אדם מופיע בכל סניף שהוא מנהל, ומספר משרד מופיע תחת עשרות שמות.
"""

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "out"
HANDOFF = DATA / "fizikal_crm_handoff" / "data"

CONFIG = ROOT / "config"
CONTACTS = HANDOFF / "Fizikal_אנשי_קשר_פעילים.xlsx"
BRANCHES = HANDOFF / "Fizikal_סניפים_פעילים.xlsx"

TIER_ORDER = {"VIP": 0, "גדול": 1, "בינוני": 2, "קטן": 3, "מוב": 4, "": 5}
TIER_FILL = {
    "VIP": "FDAB3D", "גדול": "00C875", "בינוני": "579BFC",
    "קטן": "C4C4C4", "מוב": "9D50DD",
}
HEAD_FILL = PatternFill("solid", fgColor="333333")


def s(v):
    return str(v).strip() if v is not None else ""


def digits(v):
    return re.sub(r"\D", "", str(v)) if v is not None else ""


def mobile(v):
    """
    נייד ישראלי מנורמל ל-05xxxxxxxx, או '' אם זה לא נייד.

    הייצוא מאקסל אכל את האפס המוביל כי הוא התייחס למספר כמספר ולא כטקסט,
    ולכן 9 ספרות שמתחילות ב-5 הן נייד תקין שחסר לו אפס.
    """
    p = digits(v)
    if p.startswith("972"):
        p = "0" + p[3:]
    if len(p) == 9 and p.startswith("5"):
        p = "0" + p
    return p if len(p) == 10 and p.startswith("05") else ""


def email(v):
    e = s(v).lower()
    return e if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", e) else ""


def load_exclusions():
    cfg = json.loads((CONFIG / "marketing.json").read_text(encoding="utf-8"))
    return {
        "names": re.compile("|".join(cfg["internal_name_patterns"])),
        "companies": set(cfg["internal_companies"]),
        "fake": re.compile("|".join(cfg["fake_number_patterns"])),
        "shared_threshold": cfg["shared_number_threshold"],
    }


def classify_internal(rec, ex):
    """
    האם הרשומה היא חשבון פנימי או בדיקה ולא לקוח אמיתי.

    ההתאמה על שם האדם ולא על שם החברה, כדי לא להחריג לקוח שבשמו מופיעה
    במקרה מילה כמו 'בדיקה'. מוחזרת הסיבה, כדי שההחרגה תהיה שקופה וניתנת לביטול.
    """
    if ex["names"].search(rec["name"]):
        return "שם מסמן חשבון פנימי או בדיקה"
    if rec["company"] in ex["companies"]:
        return "חברה פנימית"
    if rec["mobile"] and ex["fake"].fullmatch(rec["mobile"]):
        return "מספר דמה"
    return ""


def read_sheet(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    head = [s(c.value) for c in ws[1]]
    idx = {h: n for n, h in enumerate(head)}
    return [r for r in ws.iter_rows(min_row=2, values_only=True)], idx


def build_tier_lookup(month):
    """קוד חברה → (שם לקוח, Tier, ריטיינר), דרך ח.פ."""
    clients_file = OUT / f"clients_{month}.json"
    if not clients_file.exists():
        raise SystemExit(
            f"חסר {clients_file.name} — להריץ קודם:\n"
            f"  python3 src/revenue_pipeline.py --file 'data/הכנסות <חודש>.xlsx' --month {month}"
        )
    by_hp = {}
    for c in json.load(open(clients_file, encoding="utf-8")):
        if c["hp"] and c["hp"] != "0":
            by_hp[c["hp"]] = c

    rows, i = read_sheet(BRANCHES)
    by_code = {}
    for r in rows:
        code, hp = digits(r[i["קוד חברה"]]), digits(r[i["ח.פ"]])
        if code and hp in by_hp:
            by_code[code] = by_hp[hp]
    return by_code


def build_contacts(month):
    tier_by_code = build_tier_lookup(month)
    ex = load_exclusions()
    rows, i = read_sheet(CONTACTS)

    contacts = []
    review = {}   # (שם, סניף, הערך) → רשומה אחת עם כל ההערות, כדי לא לכפול שורות
    for r in rows:
        code = digits(r[i["קוד חברה"]])
        client = tier_by_code.get(code)
        raw = r[i["טלפון"]]
        mob, mail = mobile(raw), email(r[i['דוא"ל']])

        rec = {
            "name": s(r[i["Name"]]),
            "mobile": mob,
            "email": mail,
            "client": client["name"] if client else "",
            "tier": client["tier"] if client else "",
            "company": s(r[i["שם חברה"]]),
            "code": code,
            "branch": s(r[i["שם סניף"]]),
            "branch_no": digits(r[i["מספר סניף"]]),
            "role": s(r[i["הרשאה"]]),
            "primary": s(r[i["האם מטפל"]]),
        }
        rec["internal"] = classify_internal(rec, ex)
        contacts.append(rec)
        why = []
        if not mob and s(raw):
            why.append("טלפון בפורמט לא מזוהה")
        if not mob and not mail:
            why.append("אין נייד ואין מייל")
        if why:
            key = (rec["name"], rec["branch"], s(raw))
            review[key] = {**rec, "raw": s(raw), "why": " · ".join(why)}

    # מספר שמופיע תחת יותר משם אחד הוא כמעט תמיד קו משרד ולא נייד אישי
    names_per_mobile = defaultdict(set)
    for c in contacts:
        if c["mobile"]:
            names_per_mobile[c["mobile"]].add(c["name"])
    for mob, names in names_per_mobile.items():
        if len(names) >= ex["shared_threshold"]:
            review[("shared", mob, "")] = {
                "name": " / ".join(sorted(names)[:4]) + " …", "mobile": mob, "email": "",
                "client": "", "tier": "", "company": "", "code": "", "branch": "",
                "branch_no": "", "role": "", "primary": "", "raw": mob,
                "why": f"מספר משותף ל-{len(names)} שמות שונים — כנראה קו משרד",
            }
    return contacts, sorted(review.values(), key=lambda r: (r["why"], r["name"]))


def dedupe(contacts, field):
    """
    שורה אחת לכל נייד/מייל ייחודי. הנציג הוא הרשומה בעלת ה-Tier הגבוה ביותר,
    ובעדיפות איש קשר מטפל, כדי שהשורה שנשארת תהיה המשמעותית ביותר.
    """
    groups = defaultdict(list)
    for c in contacts:
        if c[field]:
            groups[c[field]].append(c)

    out = []
    for key, group in groups.items():
        best = min(group, key=lambda c: (TIER_ORDER.get(c["tier"], 5),
                                        c["primary"] != "כן", c["client"] or "zz"))
        clients = sorted({c["client"] or c["company"] for c in group})
        out.append({**best, "n_rows": len(group), "n_clients": len(clients),
                    "all_clients": " · ".join(clients[:6]) + (" …" if len(clients) > 6 else "")})
    out.sort(key=lambda c: (TIER_ORDER.get(c["tier"], 5), c["client"] or "zz", c["name"]))
    return out


SHEETS = {
    "דיוור - ניידים": (
        [("שם", "name", 24), ("נייד", "mobile", 14), ("Tier", "tier", 9),
         ("לקוח", "client", 30), ("שם חברה", "company", 28), ("שם סניף", "branch", 26),
         ("תפקיד", "role", 13), ("איש קשר מטפל", "primary", 13),
         ("מייל", "email", 30), ("מס' שורות במקור", "n_rows", 9),
         ("מס' לקוחות", "n_clients", 9), ("כל הלקוחות", "all_clients", 44)],
        "שורה אחת לכל נייד ייחודי — מוכן להזנה למערכת SMS."),
    "דיוור - מיילים": (
        [("שם", "name", 24), ("מייל", "email", 32), ("Tier", "tier", 9),
         ("לקוח", "client", 30), ("שם חברה", "company", 28), ("שם סניף", "branch", 26),
         ("תפקיד", "role", 13), ("איש קשר מטפל", "primary", 13),
         ("נייד", "mobile", 14), ("מס' שורות במקור", "n_rows", 9),
         ("מס' לקוחות", "n_clients", 9), ("כל הלקוחות", "all_clients", 44)],
        "שורה אחת לכל מייל ייחודי — מוכן להזנה למערכת דיוור."),
    "כל אנשי הקשר": (
        [("שם", "name", 24), ("נייד", "mobile", 14), ("מייל", "email", 30),
         ("Tier", "tier", 9), ("לקוח", "client", 30), ("שם חברה", "company", 28),
         ("קוד חברה", "code", 10), ("שם סניף", "branch", 26), ("מספר סניף", "branch_no", 10),
         ("תפקיד", "role", 13), ("איש קשר מטפל", "primary", 13),
         ("הוחרג מדיוור", "internal", 30)],
        "כל השורות כפי שהן במקור, בלי דה-דופליקציה — לעבודה פרטנית ולא לדיוור."),
    "פנימי ובדיקה - לא לדיוור": (
        [("שם", "name", 26), ("סיבת ההחרגה", "internal", 34), ("נייד", "mobile", 14),
         ("מייל", "email", 30), ("שם חברה", "company", 28), ("שם סניף", "branch", 26),
         ("תפקיד", "role", 13)],
        "הוחרג מגיליונות הדיוור. אם רשומה כאן היא בעצם לקוח אמיתי — להעביר אותה ידנית."),
    "לבדיקה": (
        [("שם", "name", 30), ("הערך במקור", "raw", 18), ("הבעיה", "why", 42),
         ("Tier", "tier", 9), ("לקוח", "client", 30), ("שם סניף", "branch", 26),
         ("מייל", "email", 30)],
        "רשומות שלא ניתן להשתמש בהן כמו שהן — לא נוחשו, מובאות לתיקון ידני."),
}


def write_xlsx(sheets_data, month):
    OUT.mkdir(exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, records in sheets_data.items():
        cols, note = SHEETS[title]
        ws = wb.create_sheet(title)
        ws.sheet_view.rightToLeft = True

        ws["A1"] = f"{note}   ({len(records):,} שורות · {month})"
        ws["A1"].font = Font(bold=True, size=11)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))

        for n, (head, _, width) in enumerate(cols, start=1):
            cell = ws.cell(row=2, column=n, value=head)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = HEAD_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(n)].width = width

        for row_n, rec in enumerate(records, start=3):
            for n, (_, key, _) in enumerate(cols, start=1):
                cell = ws.cell(row=row_n, column=n, value=rec.get(key, ""))
                # נייד נשמר כטקסט, אחרת אקסל יאכל שוב את האפס המוביל
                if key in ("mobile", "code", "branch_no", "raw"):
                    cell.number_format = "@"
            tier = rec.get("tier", "")
            if tier in TIER_FILL:
                col = next((n for n, (_, k, _) in enumerate(cols, start=1) if k == "tier"), None)
                if col:
                    ws.cell(row=row_n, column=col).fill = PatternFill("solid", fgColor=TIER_FILL[tier])

        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{max(2, len(records) + 2)}"

    path = OUT / f"Fizikal_אנשי_קשר_לשיווק_{month}.xlsx"
    wb.save(path)
    return path


def main():
    p = argparse.ArgumentParser(description="קובץ אנשי קשר לשיווק")
    p.add_argument("--month", default="2026-07", help="חודש הדוח לשיוך Tier")
    args = p.parse_args()

    contacts, review = build_contacts(args.month)
    sendable = [c for c in contacts if not c["internal"]]
    internal = sorted((c for c in contacts if c["internal"]),
                      key=lambda c: (c["internal"], c["name"]))
    by_mobile = dedupe(sendable, "mobile")
    by_email = dedupe(sendable, "email")

    for c in contacts:
        c.setdefault("n_rows", 1)
    path = write_xlsx({
        "דיוור - ניידים": by_mobile,
        "דיוור - מיילים": by_email,
        "כל אנשי הקשר": contacts,
        "פנימי ובדיקה - לא לדיוור": internal,
        "לבדיקה": review,
    }, args.month)

    tiers = defaultdict(int)
    for c in by_mobile:
        tiers[c["tier"] or "לא זוהה"] += 1
    print(f"אנשי קשר במקור:      {len(contacts):>6,}")
    print(f"הוחרגו כפנימי/בדיקה: {len(internal):>6,}")
    print(f"ניידים ייחודיים:     {len(by_mobile):>6,}   ({len(contacts) - len(by_mobile):,} כפילויות הוסרו)")
    print(f"מיילים ייחודיים:     {len(by_email):>6,}")
    print(f"לבדיקה ידנית:        {len(review):>6,}")
    print("\nניידים לפי Tier של הלקוח:")
    for t in sorted(tiers, key=lambda t: TIER_ORDER.get(t, 9)):
        print(f"  {t:<10} {tiers[t]:>5,}")
    print(f"\nנכתב: {path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
