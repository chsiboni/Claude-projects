#!/usr/bin/env python3
"""
Fizikal — קישור לקוחות (דוח ההכנסות) לקודי החברה של המערכת.

זה הגשר שבלעדיו אי אפשר לדעת איזה איש קשר שייך לאיזה לקוח משלם.

    python3 src/link_clients.py --month 2026-07

ארבעה תיקונים שהעלו את הכיסוי מ-44% ל-80%:
  · מצליבים מול רשימת 971 החברות המלאה ולא מול 910 הסניפים ה"פעילים" — הרשימה
    ההיא סוננה בעבר לפי חישוב פעילות שגוי, ולכן חסרים בה לקוחות אמיתיים.
  · מרפדים ח.פ ל-9 ספרות. אקסל אכל אפסים מובילים, ולכן 27364231 ו-027364231
    נראו כשני מספרים שונים.
  · מנסים את *כל* מספרי הח.פ של משפחה ולא רק את הראשון — רשת מאחדת כמה ח.פ.
  · נופלים לעמודת 'שם בהנה"ח', שמלאה ב-100% ומכילה בדיוק את השם החשבונאי.
"""

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl

from revenue_pipeline import (OUT, ROOT, digits, family_key, load_config,
                              open_sheet, resolve_hp_by_sap, s)

COMPANIES = ROOT / "data" / "fizikal_crm_handoff" / "data" / "Fizikal_CRM_מקור_אמת.xlsx"
TIER_ORDER = {"VIP": 0, "גדול": 1, "בינוני": 2, "קטן": 3, "מוב": 4}


def pad(hp):
    """ח.פ ל-9 ספרות. הייצוא שומר אותו כמספר ולכן אפסים מובילים נאכלים."""
    return hp.zfill(9) if hp and hp != "0" else ""


def norm(t):
    """שם חברה לצורך השוואה מדויקת: בלי צורות 'בע\"מ', גרשיים ופיסוק."""
    t = (t or "").strip().lower()
    t = re.sub(r'בע["\'״׳]?מ|בעמ|ע["\'״׳]ר|בע\.מ', "", t)
    t = re.sub(r"[\"'״׳`,\.\-–—()\[\]]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def read_companies():
    ws = openpyxl.load_workbook(COMPANIES, data_only=True)["חברות"]
    idx = {c.value: n for n, c in enumerate(ws[1])}
    by_hp, by_name = defaultdict(set), defaultdict(set)
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = digits(r[idx["קוד חברה"]])
        if not code:
            continue
        hp = pad(digits(r[idx["ח.פ"]]))
        if hp:
            by_hp[hp].add(code)
        for col in ('שם בהנה"ח', "שם חברה"):
            if norm(r[idx[col]]):
                by_name[norm(r[idx[col]])].add(code)
    return by_hp, by_name


def read_families(path, rev, nets):
    """לכל משפחת לקוח: כל מספרי הח.פ וכל השמות שהיא מופיעה תחתיהם."""
    layout = rev["layout"]
    rows = [r for r in open_sheet(path).iter_rows(min_row=layout["first_data_row"], values_only=True) if any(r)]
    hp_by_sap, _, _ = resolve_hp_by_sap(rows, layout)
    sap_re = re.compile(rev["client_sap_pattern"])
    skip_names, skip_saps = set(rev["skip_names"]), set(rev["skip_saps"])

    hps, names, of_sap = defaultdict(set), defaultdict(set), {}
    for r in rows:
        sap, name = s(r[layout["col_sap"]]), s(r[layout["col_name"]])
        if not sap_re.fullmatch(sap) or sap in skip_saps or not name:
            continue
        if name in skip_names or any(name.startswith(p) for p in rev["skip_name_prefixes"]):
            continue
        hp = hp_by_sap.get(sap, "")
        key = family_key(name, hp, sap, nets)
        of_sap[sap] = key
        names[key].add(name)
        if pad(hp):
            hps[key].add(pad(hp))
    return hps, names, of_sap


def link(clients, hps, names, of_sap, by_hp, by_name):
    out = {}
    for c in clients:
        key = of_sap.get(c["sap"])
        candidates = set(hps.get(key, set()))
        if pad(c["hp"]):
            candidates.add(pad(c["hp"]))

        codes, how = set(), ""
        for hp in candidates:
            codes |= by_hp.get(hp, set())
        if codes:
            how = "ח.פ"
        else:
            for name in names.get(key, set()) | {c["name"]}:
                if norm(name) in by_name:
                    codes |= by_name[norm(name)]
                    how = "שם"
        if codes:
            out[c["sap"]] = {"codes": sorted(codes), "how": how, "tier": c["tier"],
                             "name": c["name"], "ret": c["ret"], "status": c["status"]}
    return out


def main():
    p = argparse.ArgumentParser(description="קישור לקוחות לקודי חברה")
    p.add_argument("--month", default="2026-07")
    p.add_argument("--file", default=None)
    args = p.parse_args()

    rev, nets = load_config()
    path = Path(args.file) if args.file else sorted((ROOT / "data").glob("הכנסות*.xlsx"))[-1]
    clients = json.load(open(OUT / f"clients_status_{args.month}.json", encoding="utf-8"))

    by_hp, by_name = read_companies()
    hps, names, of_sap = read_families(path, rev, nets)
    linked = link(clients, hps, names, of_sap, by_hp, by_name)

    print("לקוחות: %d" % len(clients))
    print("  מקושרים לקוד חברה: %d  (%.0f%%)" % (len(linked), 100 * len(linked) / len(clients)))
    for how in ("ח.פ", "שם"):
        print("    לפי %s: %d" % (how, sum(1 for v in linked.values() if v["how"] == how)))
    miss = [c for c in clients if c["sap"] not in linked]
    print("  לא מקושרים: %d · ריטיינר ₪%s" % (len(miss), format(sum(c["ret"] for c in miss), ",")))

    dest = OUT / f"sap_to_company_{args.month}.json"
    dest.write_text(json.dumps(linked, ensure_ascii=False, indent=1), encoding="utf-8")
    print("\nנכתב: %s" % dest.relative_to(ROOT))


if __name__ == "__main__":
    main()
