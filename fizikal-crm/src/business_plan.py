#!/usr/bin/env python3
"""
Fizikal — מודל תוכנית עסקית: המחזור החודשי, ספטמבר 2026 עד דצמבר 2027, בשלושה תרחישים.

    python3 src/business_plan.py

החלטות Chen (24/09/2026):
  · היעד הוא על כל המחזור (חוזר + חד-פעמי + הכנסות נוספות): 30%+ על ממוצע ינואר-אוגוסט 2026 (₪1.26M) = ₪1.64M בחודש.
  · מנהלים לפי 5 מנופים: מכירות חדשות · צמצום נטישה · אפסייל · תמחור חדש · חד-פעמי (פיתוחים, חומרה וכו').
כל ההנחות בגיליון "הנחות" (כחול = אפשר לשנות), והכל מחושב בנוסחאות.
גיליון "יעדים ומעקב" נותן יעד חודשי לכל מנוף לפי התרחיש שנבחר, ועמודות "בפועל" למילוי כל חודש.

נתוני הבסיס: data/revenue_latest.xlsx (דוח ההכנסות, ינואר-אוגוסט 2026), שורות של סאפים 2xxxxx/3xxxxx בלבד
(בלי שורות סיכום והתאמות חשבונאיות).
פלט: out/Fizikal_תוכנית_עסקית_2027.xlsx
"""

from datetime import date

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from revenue_pipeline import OUT

F = "Arial"
BLUE, GREEN = "0000FF", "008000"
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
SEC_FILL = PatternFill("solid", fgColor="D9E1F2")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
IN_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
NIS = '₪#,##0;(₪#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
MON = 'mmm-yy'
N = 16                                                  # ספטמבר 2026 → דצמבר 2027

SCEN = ["שמרני", "בסיס", "אגרסיבי"]
SCOL = {"שמרני": "B", "בסיס": "C", "אגרסיבי": "D"}
INP = "'הנחות'!"
MOD = "'מודל חודשי'!"

# ---- נקודת פתיחה (ערך אחד) ----
START = [  # key, label, value, fmt, note
    ("base_ret", "ריטיינר חודשי", 820_000, NIS,
     "אוגוסט 2026: ₪822K, כולל קפיצה חד-פעמית של עיריית רעננה (₪18K+). ריטיינר + מודולים מעוגל ל-₪1.0M (החלטת Chen)."),
    ("base_mod", "מודולים חודשי", 180_000, NIS,
     "אוגוסט 2026: ₪190K, כולל ~₪20K חיובים חריגים (טן דאנס, רמת פולג, מגדלי דוד, רעננה). יולי: ₪175K."),
    ("base_one", "חד-פעמי חודשי (חומרה, פיתוחים, SMS חד-פעמי, תגים, טכנאי, התקנות)", 195_000, NIS,
     "ממוצע ינואר-אוגוסט 2026. טווח: ₪53K (אוגוסט) עד ₪315K (מרץ)."),
    ("base_other", "הכנסות נוספות חודשי (SMS ריטיינר, מכבי MOVE, עמלות, שת\"פ)", 77_000, NIS,
     "ממוצע ינואר-אוגוסט 2026. SMS ריטיינר תלוי שימוש ולכן קופץ."),
    ("target", "יעד מחזור חודשי — דצמבר 2027", 1_640_000, NIS,
     "החלטת Chen: 30%+ על ממוצע המחזור בינואר-אוגוסט 2026 (₪1,261K)."),
    ("clients", "לקוחות פעילים", 407, "#,##0", "client_status אוגוסט 2026: 399 פעיל + 8 חדש."),
]

# ---- הנחות לפי תרחיש ----
SCEN_INPUTS = [  # section/None, key, label, (cons, base, aggr), fmt, note, key-assumption?
    ("מנוף 1 — מכירות חדשות", None, None, None, None, None, False),
    (None, "new_before", "ריטיינר מלקוחות חדשים לחודש — היום", (7_200, 7_200, 7_200), NIS,
     "בפועל פברואר-אוגוסט 2026: ₪43K מ-101 לקוחות חדשים ב-6 חודשים (ממוצע ₪430 ללקוח).", False),
    (None, "new_after", "ריטיינר מלקוחות חדשים לחודש — אחרי המיתוג", (9_000, 12_000, 18_000), NIS,
     "הנחה: מיתוג מחדש + מחירון באתר מגדילים את המכירות החדשות.", True),
    (None, "new_start", "המיתוג משפיע החל מ-", (date(2027, 1, 1),) * 3, MON, "", False),
    ("מנוף 2 — צמצום נטישה", None, None, None, None, None, False),
    (None, "churn_base", "נטישה חודשית — היום", (0.0075, 0.0075, 0.0075), PCT2,
     "בפועל: 43 לקוחות עזבו ב-6 חודשים, ₪44K = כ-0.75% מההכנסה החוזרת בחודש.", False),
    (None, "churn_target", "נטישה חודשית — אחרי תוכנית שימור", (0.0070, 0.0050, 0.0040), PCT2,
     "הנחה: תוכנית Customer Success ללקוחות בסיכון.", True),
    (None, "churn_start", "תוכנית השימור משפיעה החל מ-", (date(2027, 1, 1),) * 3, MON, "", False),
    ("מנוף 3 — אפסייל", None, None, None, None, None, False),
    (None, "expansion", "הגדלות נטו חודשיות אצל לקוחות קיימים", (0.0035, 0.0045, 0.0055), PCT2,
     "בפועל: הגדלות ₪82K פחות הקטנות ₪54K ב-6 חודשים = כ-0.47% בחודש.", False),
    (None, "mod_start", "מכירת מודולים יזומה מתחילה ב-", (date(2026, 10, 1),) * 3, MON, "", False),
    (None, "mod_clients", "לקוחות שמוסיפים מודול בחודש", (2, 4, 6), "0",
     "79 לקוחות בינוני/גדול בלי אף מודול (ריטיינר ₪134K).", True),
    (None, "mod_price", "ערך מודול ממוצע לחודש", (500, 500, 500), NIS,
     "היום: ₪175K מודולים על 163 לקוחות = ~₪1,070 ללקוח. מודול ראשון נמוך מזה.", False),
    (None, "addon_start", "פיצ'רים חדשים (₪50-100) — השקה ב-", (date(2027, 3, 1), date(2027, 1, 1), date(2026, 11, 1)), MON, "", False),
    (None, "addon_adopt", "פיצ'רים חדשים — אימוץ בסוף הרמפה (מתוך הלקוחות)", (0.15, 0.25, 0.40), PCT, "", True),
    (None, "addon_price", "פיצ'רים חדשים — מחיר ממוצע לחודש", (75, 75, 90), NIS, "החלטת Chen: ₪50-100 לפיצ'ר.", False),
    (None, "addon_ramp", "פיצ'רים חדשים — חודשים עד אימוץ מלא", (12, 12, 12), "0", "", False),
    ("מנוף 4 — תמחור חדש (בחידוש חוזה)", None, None, None, None, None, False),
    (None, "price_start", "המחירון החדש מתחיל ב-", (date(2027, 1, 1),) * 3, MON, "", False),
    (None, "price_months", "חודשים עד שכל הבסיס מחדש", (12, 12, 12), "0",
     "חוזים שנתיים: כל חודש מתחדשת 1/12 מהבסיס.", False),
    (None, "price_share", "חלק מהריטיינר שעובר למחירון החדש", (0.50, 0.70, 0.85), PCT,
     "השאר: VIP/גדול במחיר מותאם אישית, או לקוחות שמחירם כבר מעל המחירון.", True),
    (None, "price_uplift", "העלאה ממוצעת למי שעובר", (0.10, 0.15, 0.20), PCT,
     "לכייל מול סימולציית המחירון החדש לפי לקוח — לפני שמתחייבים למספר.", True),
    (None, "price_churn", "נטישה בגלל המחיר (מתוך מי שעובר)", (0.08, 0.05, 0.03), PCT,
     "ברוב חברות ה-SaaS, העלאה של 10%-20% מגדילה את הנטישה ב-3%-8% מהלקוחות שהמחיר שלהם עולה.", True),
    ("מנוף 5 — חד-פעמי (פיתוחים, חומרה, SMS, תגים)", None, None, None, None, None, False),
    (None, "one_end", "חד-פעמי חודשי — דצמבר 2027", (195_000, 260_000, 330_000), NIS,
     "עלייה הדרגתית מהבסיס. חומרה ופיתוחים גדלים עם לקוחות חדשים וגדולים. מודדים כממוצע 3 חודשים.", True),
    ("הכנסות נוספות (SMS ריטיינר, מכבי, עמלות, שת\"פ)", None, None, None, None, None, False),
    (None, "other_growth", "צמיחה חודשית", (0.0, 0.005, 0.01), PCT2, "", False),
]

# שורות בלוק תרחיש במודל החודשי
ROWS = [
    ("rec_open", "הכנסה חוזרת — פתיחה"),
    ("new", "מנוף 1: מכירות חדשות"),
    ("churn", "נטישה בקצב של היום"),
    ("saved", "מנוף 2: צמצום נטישה"),
    ("exp", "מנוף 3: אפסייל — הגדלות נטו"),
    ("mod_new", "מנוף 3: אפסייל — מודולים"),
    ("addon", "מנוף 3: אפסייל — פיצ'רים חדשים"),
    ("up", "מנוף 4: תמחור — העלאה בחידוש"),
    ("pchurn", "מנוף 4: תמחור — נטישה בגלל מחיר"),
    ("rec_close", "הכנסה חוזרת — סגירה"),
    ("one", "מנוף 5: חד-פעמי"),
    ("other", "הכנסות נוספות"),
    ("total", "סה\"כ מחזור"),
    ("path", "מסלול ליעד"),
    ("gap", "פער מהמסלול"),
    ("growth", "צמיחה מהבסיס"),
]
BOLD_ROWS = ("rec_close", "total")

# מנופים בגיליון היעדים והסיכום: (תווית, מפתחות בבלוק, איך מודדים בפועל)
LEVERS = [
    ("מכירות חדשות", ["new"],
     "ריטיינר + מודולים של סאפים שלא שילמו בחודש הקודם."),
    ("נטישה נטו", ["churn", "saved"],
     "ריטיינר + מודולים של סאפים ששילמו בחודש הקודם ו-0 החודש (מספר שלילי)."),
    ("אפסייל", ["exp", "mod_new", "addon"],
     "שינוי נטו בריטיינר + מודולים אצל לקוחות שממשיכים, בלי אלה שעברו למחירון החדש."),
    ("תמחור חדש", ["up", "pchurn"],
     "שינוי אצל לקוחות שעברו למחירון החדש החודש (צריך סימון בבורד: 'עבר למחירון')."),
]


def font(color="000000", bold=False, size=10):
    return Font(name=F, color=color, bold=bold, size=size)


def head(ws, r, c, text):
    x = ws.cell(r, c, text); x.font = font("FFFFFF", True); x.fill = HEAD_FILL
    x.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    return x


def build_inputs(wb):
    ws = wb.active; ws.title = "הנחות"; ws.sheet_view.rightToLeft = True
    ws["A1"] = "Fizikal — תוכנית עסקית 2027: הנחות"; ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = "כחול = הנחה שאפשר לשנות · צהוב = הנחות מפתח · שחור = נוסחה. כל הקובץ מתעדכן אוטומטית."
    ws["A2"].font = font("595959")
    ref = {}
    r = 4
    ws.cell(r, 1, "נקודת פתיחה").font = font(bold=True)
    for c in range(1, 7): ws.cell(r, c).fill = SEC_FILL
    r += 1
    for key, label, val, fmt, note in START:
        ws.cell(r, 1, label).font = font()
        c = ws.cell(r, 2, val); c.font = font(BLUE); c.number_format = fmt
        if key == "target": c.fill = KEY_FILL
        ws.cell(r, 6, note).font = font("595959")
        ref[key] = f"{INP}$B${r}"
        r += 1
    ws.cell(r, 1, "סה\"כ מחזור חודשי — בסיס").font = font(bold=True)
    c = ws.cell(r, 2, f"=SUM(B5:B8)"); c.font = font(bold=True); c.number_format = NIS
    ref["base_total"] = f"{INP}$B${r}"
    ref["base_rec"] = f"({INP}$B$5+{INP}$B$6)"
    r += 2
    for i, h in enumerate(["הנחה", *SCEN, "", "מקור / הערה"], 1):
        head(ws, r, i, h)
    ref["scen_hdr"] = f"{INP}$B${r}:$D${r}"
    r += 1
    for sec, key, label, vals, fmt, note, keyflag in SCEN_INPUTS:
        if sec:
            ws.cell(r, 1, sec).font = font(bold=True)
            for c in range(1, 7): ws.cell(r, c).fill = SEC_FILL
            r += 1; continue
        ws.cell(r, 1, label).font = font()
        for j, v in enumerate(vals):
            c = ws.cell(r, 2 + j, v); c.font = font(BLUE); c.number_format = fmt
            if keyflag: c.fill = KEY_FILL
        ws.cell(r, 6, note).font = font("595959")
        ref[key] = r
        r += 1
    ws.column_dimensions["A"].width = 58
    for col in "BCD": ws.column_dimensions[col].width = 13
    ws.column_dimensions["E"].width = 2
    ws.column_dimensions["F"].width = 95
    ws.freeze_panes = "B4"
    return ref


def month_header(ws, row, first_col, with_base):
    """שורת חודשים: עמודת בסיס (אוגוסט 2026) אם with_base, ואחריה 16 חודשים."""
    c0 = first_col
    if with_base:
        ws.cell(row, c0, date(2026, 8, 1)); c0 += 1
        ws.cell(row, c0, f"=DATE(YEAR({L(c0-1)}{row}),MONTH({L(c0-1)}{row})+1,1)")
    else:
        ws.cell(row, c0, date(2026, 9, 1))
    for c in range(c0 + 1, c0 + N):
        ws.cell(row, c, f"=DATE(YEAR({L(c-1)}{row}),MONTH({L(c-1)}{row})+1,1)")
    for c in range(first_col, c0 + N):
        x = ws.cell(row, c); x.number_format = MON; x.font = font("FFFFFF", True); x.fill = HEAD_FILL
        x.alignment = Alignment(horizontal="center")


def build_model(wb, ref):
    ws = wb.create_sheet("מודל חודשי"); ws.sheet_view.rightToLeft = True
    ws["A1"] = "המחזור החודשי לפי מנופים, ₪"; ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = "עמודה B = בסיס (ממוצע 2026). הכנסה חוזרת = ריטיינר + מודולים + פיצ'רים חדשים."
    ws["A2"].font = font("595959")
    first, last = 3, 3 + N - 1                         # C..R
    head(ws, 3, 1, "חודש")
    month_header(ws, 3, 2, with_base=True)
    ws.cell(4, 2, "בסיס").font = font("595959"); ws.cell(4, 2).alignment = Alignment(horizontal="center")

    def P(key, s):
        return f"{INP}${SCOL[s]}${ref[key]}"

    blocks = {}
    r = 6
    for s in SCEN:
        R = {}
        ws.cell(r, 1, f"תרחיש {s}").font = font("FFFFFF", True, 12)
        for c in range(1, last + 1): ws.cell(r, c).fill = HEAD_FILL
        r += 1
        for k, lab in ROWS:
            R[k] = r; ws.cell(r, 1, lab).font = font(bold=k in BOLD_ROWS); r += 1
        # עמודת בסיס
        ws.cell(R["rec_close"], 2, f"={ref['base_rec']}").font = font(GREEN, True)
        ws.cell(R["one"], 2, f"={ref['base_one']}").font = font(GREEN)
        ws.cell(R["other"], 2, f"={ref['base_other']}").font = font(GREEN)
        ws.cell(R["total"], 2, f"=B{R['rec_close']}+B{R['one']}+B{R['other']}")
        ws.cell(R["path"], 2, f"={ref['base_total']}").font = font(GREEN)
        chunk = f"{ref['base_ret']}*{P('price_share', s)}/{P('price_months', s)}"
        for c in range(first, last + 1):
            X, Pv, n = L(c), L(c - 1), c - 2          # n = 1..16
            d = f"{X}$3"
            ps = P('price_start', s)
            in_price = f"AND({d}>={ps},{d}<DATE(YEAR({ps}),MONTH({ps})+{P('price_months', s)},1))"
            churn_rate = f"IF({d}>={P('churn_start', s)},{P('churn_target', s)},{P('churn_base', s)})"
            a0 = P('addon_start', s)
            months_live = f"((YEAR({d})-YEAR({a0}))*12+MONTH({d})-MONTH({a0})+1)"
            addon_step = f"{ref['clients']}*{P('addon_adopt', s)}*{P('addon_price', s)}/{P('addon_ramp', s)}"
            f = {
                "rec_open": f"={Pv}{R['rec_close']}",
                "new": f"=IF({d}>={P('new_start', s)},{P('new_after', s)},{P('new_before', s)})",
                "churn": f"=-{X}{R['rec_open']}*{P('churn_base', s)}",
                "saved": f"={X}{R['rec_open']}*({P('churn_base', s)}-{churn_rate})",
                "exp": f"={ref['base_ret']}/{ref['base_rec']}*{X}{R['rec_open']}*{P('expansion', s)}",
                "mod_new": f"=IF({d}>={P('mod_start', s)},{P('mod_clients', s)}*{P('mod_price', s)},0)",
                "addon": f"=IF(AND({months_live}>=1,{months_live}<={P('addon_ramp', s)}),{addon_step},0)",
                "up": f"=IF({in_price},{chunk}*(1-{P('price_churn', s)})*{P('price_uplift', s)},0)",
                "pchurn": f"=-IF({in_price},{chunk}*{P('price_churn', s)},0)",
                "rec_close": f"=SUM({X}{R['rec_open']}:{X}{R['pchurn']})",
                "one": f"={ref['base_one']}+({P('one_end', s)}-{ref['base_one']})*{n}/{N}",
                "other": f"={Pv}{R['other']}*(1+{P('other_growth', s)})",
                "total": f"={X}{R['rec_close']}+{X}{R['one']}+{X}{R['other']}",
                "path": f"={ref['base_total']}+({ref['target']}-{ref['base_total']})*{n}/{N}",
                "gap": f"={X}{R['total']}-{X}{R['path']}",
                "growth": f"={X}{R['total']}/{ref['base_total']}-1",
            }
            for k, v in f.items():
                ws.cell(R[k], c, v)
        for k, rr in R.items():
            for c in range(2, last + 1):
                cell = ws.cell(rr, c); cell.number_format = PCT if k == "growth" else NIS
                if cell.font.color is None or cell.font.color.rgb in (None, "FF000000"):
                    cell.font = font(bold=k in BOLD_ROWS)
                if k == "total":
                    cell.fill = TOT_FILL; cell.border = Border(top=THIN, bottom=THIN)
            if k == "total": ws.cell(rr, 1).fill = TOT_FILL
        blocks[s] = R
        r += 1
    ws.cell(r + 1, 1, "הכנסה חוזרת: כל תנועה נכנסת לבסיס ונשארת (ולכן גם נחשפת לנטישה בחודשים הבאים). "
                      "חד-פעמי והכנסות נוספות הם רמה חודשית, לא מצטברים.").font = font("595959")
    ws.column_dimensions["A"].width = 32
    for c in range(2, last + 1): ws.column_dimensions[L(c)].width = 11.5
    ws.freeze_panes = "B4"
    return ws, blocks, first, last


def build_summary(wb, model, blocks, first, last, ref):
    ws = wb.create_sheet("סיכום", 0); ws.sheet_view.rightToLeft = True
    ws["A1"] = "Fizikal — תוכנית עסקית: מ-₪1.26M ל-₪1.64M מחזור בחודש עד דצמבר 2027"
    ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = "מחזור = הכנסה חוזרת (ריטיינר + מודולים) + חד-פעמי (פיתוחים, חומרה, SMS, תגים) + הכנסות נוספות (SMS ריטיינר, מכבי, עמלות, שת\"פ)."
    ws["A2"].font = font("595959")
    Dc = L(last)
    r = 4
    for i, h in enumerate(["", *SCEN], 1): head(ws, r, i, h)
    lines = [
        ("מחזור חודשי — בסיס", lambda R: f"={MOD}B{R['total']}", NIS, False),
        ("מחזור חודשי — דצמבר 2027", lambda R: f"={MOD}{Dc}{R['total']}", NIS, True),
        ("צמיחה", lambda R: f"={MOD}{Dc}{R['growth']}", PCT, False),
        ("יעד", lambda R: f"={ref['target']}", NIS, False),
        ("פער ליעד (שלילי = חסר)", lambda R: f"={MOD}{Dc}{R['total']}-{ref['target']}", NIS, True),
        ("מגיע ליעד?", lambda R: f"=IF({MOD}{Dc}{R['total']}>={ref['target']},\"כן\",\"לא\")", "@", False),
        ("מתוכו: הכנסה חוזרת בדצמבר 2027", lambda R: f"={MOD}{Dc}{R['rec_close']}", NIS, False),
    ]
    for lab, fn, fmt, bold in lines:
        r += 1
        ws.cell(r, 1, lab).font = font(bold=bold)
        for j, s in enumerate(SCEN):
            c = ws.cell(r, 2 + j, fn(blocks[s])); c.number_format = fmt; c.font = font(GREEN, bold)
            c.alignment = Alignment(horizontal="center")
    r += 2
    ws.cell(r, 1, "מאיפה מגיע השינוי בדצמבר 2027 מול הבסיס").font = font(bold=True)
    for c in range(1, 5): ws.cell(r, c).fill = SEC_FILL
    b0 = r + 1
    bridge = [
        ("מנוף 1: מכירות חדשות", lambda R: f"=SUM({MOD}{L(first)}{R['new']}:{Dc}{R['new']})"),
        ("נטישה בקצב של היום", lambda R: f"=SUM({MOD}{L(first)}{R['churn']}:{Dc}{R['churn']})"),
        ("מנוף 2: צמצום נטישה", lambda R: f"=SUM({MOD}{L(first)}{R['saved']}:{Dc}{R['saved']})"),
        ("מנוף 3: אפסייל (הגדלות + מודולים + פיצ'רים)",
         lambda R: "=" + "+".join(f"SUM({MOD}{L(first)}{R[k]}:{Dc}{R[k]})" for k in ("exp", "mod_new", "addon"))),
        ("מנוף 4: תמחור חדש (העלאה פחות נטישה בגלל מחיר)",
         lambda R: "=" + "+".join(f"SUM({MOD}{L(first)}{R[k]}:{Dc}{R[k]})" for k in ("up", "pchurn"))),
        ("מנוף 5: חד-פעמי", lambda R: f"={MOD}{Dc}{R['one']}-{MOD}B{R['one']}"),
        ("הכנסות נוספות", lambda R: f"={MOD}{Dc}{R['other']}-{MOD}B{R['other']}"),
    ]
    for lab, fn in bridge:
        r += 1
        ws.cell(r, 1, lab).font = font()
        for j, s in enumerate(SCEN):
            c = ws.cell(r, 2 + j, fn(blocks[s])); c.number_format = NIS; c.font = font(GREEN)
    r += 1
    ws.cell(r, 1, "סה\"כ שינוי").font = font(bold=True); ws.cell(r, 1).fill = TOT_FILL
    for j in range(3):
        col = L(2 + j)
        c = ws.cell(r, 2 + j, f"=SUM({col}{b0}:{col}{r-1})"); c.number_format = NIS; c.font = font(bold=True); c.fill = TOT_FILL
    r += 1
    ws.cell(r, 1, "בדיקה: בסיס + שינוי − דצמבר (צריך להיות 0)").font = font("595959")
    for j in range(3):
        col = L(2 + j)
        c = ws.cell(r, 2 + j, f"=ROUND({col}5+{col}{r-1}-{col}6,0)"); c.number_format = NIS; c.font = font("595959")
    for col, w in zip("ABCD", (52, 15, 15, 15)): ws.column_dimensions[col].width = w

    ch = LineChart(); ch.title = "מחזור חודשי מול מסלול היעד"; ch.height = 9; ch.width = 22
    ch.y_axis.title = "₪"; ch.y_axis.numFmt = '₪#,##0'; ch.x_axis.number_format = MON
    colors = {"שמרני": "EB6834", "בסיס": "2A78D6", "אגרסיבי": "1BAF7A"}
    for s in SCEN:
        ch.add_data(Reference(model, min_col=2, max_col=last, min_row=blocks[s]["total"]), from_rows=True, titles_from_data=False)
        ser = ch.series[-1]; ser.tx = SeriesLabel(v=s)
        ser.graphicalProperties.line.solidFill = colors[s]; ser.graphicalProperties.line.width = 25400; ser.smooth = False
    ch.add_data(Reference(model, min_col=2, max_col=last, min_row=blocks["בסיס"]["path"]), from_rows=True, titles_from_data=False)
    ser = ch.series[-1]; ser.tx = SeriesLabel(v="מסלול ליעד")
    ser.graphicalProperties.line.solidFill = "8C8C8C"; ser.graphicalProperties.line.dashStyle = "dash"
    ser.graphicalProperties.line.width = 19050; ser.smooth = False
    ch.set_categories(Reference(model, min_col=2, max_col=last, min_row=3))
    ch.legend.position = "b"
    ws.add_chart(ch, f"A{r + 3}")


def build_tracking(wb, blocks, first, ref):
    """יעד חודשי לכל מנוף (לפי התרחיש שנבחר) + עמודות בפועל למילוי."""
    ws = wb.create_sheet("יעדים ומעקב", 1); ws.sheet_view.rightToLeft = True
    ws["A1"] = "יעדים ומעקב חודשי לפי מנוף"; ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = "תרחיש היעד:"; ws["A2"].font = font(bold=True)
    sel = ws["B2"]; sel.value = "בסיס"; sel.font = font(BLUE, True); sel.fill = KEY_FILL
    dv = DataValidation(type="list", formula1='"שמרני,בסיס,אגרסיבי"', allow_blank=False); ws.add_data_validation(dv); dv.add("B2")
    ws["C2"] = f"=MATCH(B2,{ref['scen_hdr']},0)"; ws["C2"].font = font("FFFFFF")      # אינדקס התרחיש (מוסתר בלבן)
    ws["D2"] = "הצהובים = למלא כל חודש מדוח ההכנסות. יעד המנופים 1-4 = תוספת להכנסה החוזרת באותו חודש; חד-פעמי ומחזור = הרמה החודשית."
    ws["D2"].font = font("595959")

    groups = [("מחזור", None)] + [(lab, keys) for lab, keys, _ in LEVERS] + [("חד-פעמי", ["one"])]
    r0 = 4
    head(ws, r0, 1, "חודש"); ws.merge_cells(start_row=r0, start_column=1, end_row=r0 + 1, end_column=1)
    col = 2; gcols = {}
    for lab, _ in groups:
        n_sub = 3
        head(ws, r0, col, lab); ws.merge_cells(start_row=r0, start_column=col, end_row=r0, end_column=col + n_sub - 1)
        for i, sub in enumerate(["יעד", "בפועל", "פער"]): head(ws, r0 + 1, col + i, sub)
        gcols[lab] = col; col += n_sub
    last_col = col - 1

    def pick(keys, mc):
        parts = []
        for s in SCEN:
            R = blocks[s]
            parts.append("+".join(f"{MOD}{mc}{R[k]}" for k in keys))
        return f"=CHOOSE($C$2,{','.join(parts)})"

    for i in range(N):
        rr = r0 + 2 + i
        mc = L(first + i)
        x = ws.cell(rr, 1, f"={MOD}{mc}3"); x.number_format = MON; x.font = font(GREEN)
        for lab, keys in groups:
            c = gcols[lab]
            ws.cell(rr, c, pick(keys or ["total"], mc)).font = font(GREEN)
            a = ws.cell(rr, c + 1); a.fill = IN_FILL; a.font = font(BLUE)
            g = ws.cell(rr, c + 2, f'=IF({L(c+1)}{rr}="","",{L(c+1)}{rr}-{L(c)}{rr})')
            for k in range(3): ws.cell(rr, c + k).number_format = NIS
            ws.conditional_formatting.add(f"{L(c+2)}{rr}", FormulaRule(formula=[f'AND({L(c+2)}{rr}<>"",{L(c+2)}{rr}<0)'], font=Font(name=F, color="C00000", bold=True)))
            ws.conditional_formatting.add(f"{L(c+2)}{rr}", FormulaRule(formula=[f'AND({L(c+2)}{rr}<>"",{L(c+2)}{rr}>=0)'], font=Font(name=F, color="008000", bold=True)))
    rr = r0 + 2 + N + 1
    ws.cell(rr, 1, "איך מודדים בפועל").font = font(bold=True)
    for c in range(1, 8): ws.cell(rr, c).fill = SEC_FILL
    notes = [("מחזור", "סה\"כ הכנסות החודש מכל הסעיפים, סאפים 2xxxxx/3xxxxx.")] + \
            [(lab, how) for lab, _, how in LEVERS] + \
            [("חד-פעמי", "חומרה + פיתוח + SMS חד-פעמי + תגי קרבה + טכנאי + התקנות + משלוחים. להסתכל על ממוצע 3 חודשים.")]
    for lab, how in notes:
        rr += 1
        ws.cell(rr, 1, lab).font = font(bold=True); ws.cell(rr, 2, how).font = font("595959")
    ws.cell(rr + 2, 1, "את הבפועל אני יכול לחשב אוטומטית מכל דוח הכנסות חודשי חדש.").font = font("595959")
    ws.column_dimensions["A"].width = 16
    for c in range(2, last_col + 1): ws.column_dimensions[L(c)].width = 11
    ws.row_dimensions[r0].height = 30
    ws.freeze_panes = ws.cell(r0 + 2, 2)


def build_actuals(wb):
    ws = wb.create_sheet("בפועל 2026"); ws.sheet_view.rightToLeft = True
    ws["A1"] = "בפועל ינואר-אוגוסט 2026 (₪)"; ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = ("מקור: דוח ההכנסות המעודכן (data/revenue_latest.xlsx), סאפים 2xxxxx ו-3xxxxx, "
                "בלי שורות סיכום והתאמות חשבונאיות. חד-פעמי והכנסות נוספות מעוגלים לאלפים.")
    ws["A2"].font = font("595959")
    months = [date(2026, m, 1) for m in range(1, 9)]
    ret = [811_800, 801_600, 800_800, 832_800, 807_200, 810_100, 822_500, 822_500]
    mod = [168_800, 173_500, 173_600, 170_000, 173_700, 173_600, 175_300, 190_300]
    one = [268_000, 181_000, 315_000, 145_000, 128_000, 297_000, 173_000, 53_000]
    oth = [14_000, 119_000, 93_000, 60_000, 76_000, 75_000, 108_000, 75_000]
    act = [399, 410, 425, 437, 460, 476, 468, 464]
    for i, h in enumerate(["חודש", "ריטיינר", "מודולים", "חוזר", "חד-פעמי", "נוספות", "סה\"כ מחזור", "לקוחות משלמים"], 1):
        head(ws, 4, i, h)
    for k, m in enumerate(months):
        rr = 5 + k
        ws.cell(rr, 1, m).number_format = MON
        for c, v in ((2, ret[k]), (3, mod[k]), (5, one[k]), (6, oth[k])):
            x = ws.cell(rr, c, v); x.number_format = NIS; x.font = font(BLUE)
        ws.cell(rr, 4, f"=B{rr}+C{rr}").number_format = NIS
        ws.cell(rr, 7, f"=D{rr}+E{rr}+F{rr}").number_format = NIS
        x = ws.cell(rr, 8, act[k]); x.number_format = "#,##0"; x.font = font(BLUE)
    ws.cell(13, 1, "ממוצע").font = font(bold=True)
    for c in range(2, 8):
        x = ws.cell(13, c, f"=AVERAGE({L(c)}5:{L(c)}12)"); x.number_format = NIS; x.font = font(bold=True); x.fill = TOT_FILL
    ws.cell(12, 4).comment = Comment("כולל ~₪35K חיובים חריגים באוגוסט: עיריית רעננה, טן דאנס, רמת פולג, מגדלי דוד. לבדוק עם הנהלת החשבונות.", "Claude")

    ws.cell(15, 1, "גשר הכנסה חוזרת: ינואר-פברואר → יולי-אוגוסט (ממוצע לחודש)").font = font(bold=True)
    for c in range(1, 4): ws.cell(15, c).fill = SEC_FILL
    for i, h in enumerate(["תנועה", "לקוחות", "₪ לחודש"], 1): head(ws, 16, i, h)
    bridge = [("לקוחות חדשים", 101, 43_100), ("לקוחות שעזבו", 43, -44_100),
              ("הגדלות", None, 81_500), ("הקטנות", None, -53_800)]
    for k, (lab, n, v) in enumerate(bridge):
        rr = 17 + k
        ws.cell(rr, 1, lab)
        if n: ws.cell(rr, 2, n).font = font(BLUE)
        c = ws.cell(rr, 3, v); c.number_format = NIS; c.font = font(BLUE)
    ws.cell(21, 1, "נטו").font = font(bold=True)
    ws.cell(21, 3, "=SUM(C17:C20)").number_format = NIS
    ws.cell(23, 1, "לקוח חדש ממוצע: ₪430 בחודש. לקוח שעזב ממוצע: ₪1,025 בחודש. מביאים קטנים, מאבדים גדולים.").font = font("595959")
    for col, w in zip("ABCDEFGH", (22, 13, 13, 13, 13, 13, 14, 14)): ws.column_dimensions[col].width = w


def main():
    wb = openpyxl.Workbook()
    ref = build_inputs(wb)
    model, blocks, first, last = build_model(wb, ref)
    build_summary(wb, model, blocks, first, last, ref)
    build_tracking(wb, blocks, first, ref)
    build_actuals(wb)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and c.font is not None and c.font.name != F:
                    c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, color=c.font.color)
    wb.calculation.fullCalcOnLoad = True
    out = OUT / "Fizikal_תוכנית_עסקית_2027.xlsx"
    wb.save(out)
    print(out)


if __name__ == "__main__":
    main()
