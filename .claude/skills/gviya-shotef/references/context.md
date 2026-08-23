# gviya-shotef · הקשר טכני

## הלוח

- **לוח הגבייה:** `5102318900` · https://leapapp147s-team.monday.com/boards/5102318900
- מפתח הסנכרון: עמודת `קוד לקוח` (`text_mm69gg7g`) — ייחודי. ח.פ. אינו מפתח (חוזר עד 22 פעם).

### עמודות רלוונטיות

| עמודה | id | הערה |
|---|---|---|
| שלב | `color_mm69ws8c` | סטטוס. כתיבה לפי **אינדקס** בלבד — טקסט נדחה |
| קוד לקוח | `text_mm69gg7g` | מפתח. לא לערוך |
| חוב | `numeric_mm69dbtr` | בבעלות הסנכרון. לא לגעת |
| מצב יתרה | `color_mm69xcy2` | בבעלות הסנכרון. לא לגעת |
| תנאי תשלום | `text_mm69ew8e` | עמודת 'שוטף' מהדוח. מקור האמת הוא ה-XLSX |
| נציגה | `multiple_person_mm6933zz` | לא לגעת |

### אינדקסי שלבים (color_mm69ws8c)

`ללא חוב`=0 · `לקוח בחוב`=1 · `בחינת הנהלה`=2 · `פניה ראשונה`=3 · `פניה שניה`=4 ·
`בירור שאלות מלקוח`=5 · `אישור בקשה`=6 · `המתנה לתשלום`=7 · `אישור תשלום`=8 ·
`סירוב`=9 · **`חוב בתחום שוטף`=10** · `קיזוז`=11 · `חסימת גישה`=12 · `לקוח עבר`=13

## מקורות נתונים

- **דוח גיול:** XLSX מהכספים, גיליון לכל שבוע (למשל `16.08.26`). ברירת מחדל:
  הגיליון האחרון. קריאה לפי שם כותרת בלבד — הפריסה נודדת בין שבועות.
  מגיע במייל מ-YasminB@movement-group.com (נושא: "גיול לקוחות חייבים"),
  או שחן מעלה ישירות.
- **היסטוריית סנאפשוטים:** קבצי `YYYY-MM-DD.json` בפורמט
  `{"date","sheet","clients":{code:[יתרה, בפיגור]}}`. מקור: `fizikal-gviya/history/`
  (חן מעלה את התיקייה/זיפ, או שקיימת מסשן קודם ב-`/tmp/gviya/fizikal-gviya/history`).
- **קובץ החזרי הו"ק:** "סטטוס הוק חוזרות" של יסמין. שמות משובשים בו — לעולם לא
  מפתח; הגשר הוא חשבון-בנק→קוד (נבדק: אפס התנגשויות).

## תבנית פרומפט לסוכן-הביצוע (רקע)

החלף `<APPLY>` בנתיב קובץ התוכנית ו-`<N>` במספר השורות בו:

```
You are applying a pre-approved, fully-specified set of changes to a monday.com
board. The plan is decided — do NOT re-derive, change, add, or skip anything.

READ THE PLAN FIRST:  cat <APPLY>
JSON array of <N> objects: item_id, name, cur, target, target_index, update.

BOARD: 5102318900 · STAGE COLUMN: color_mm69ws8c

For each object, in order:
  A. Post the `update` string VERBATIM as an item update via monday MCP
     `create_update`. Do not reword or trim.
  B. If target is not null: move the stage via `change_item_column_values`,
     column_values = {"color_mm69ws8c": {"index": <target_index>}}.
     Send index, never a text label.

RULES:
  - Touch NO other column (never numeric_mm69dbtr, color_mm69xcy2,
    multiple_person_mm6933zz, or any text column).
  - Touch NO item not in the file. No creates/deletes/archives.
  - On a failed call: retry once, then record the failure and continue.
    Never abort the run.

VERIFY AT THE END: re-read the items via get_board_items_page (itemIds =
the ids from the plan, includeColumns true, columnIds
["color_mm69ws8c","text_mm69gg7g","numeric_mm69dbtr"]). Confirm every mover
shows its target stage and numeric_mm69dbtr is unchanged.

FINAL REPORT: updates posted / total, moves applied / total, compact
before→after table, failures with error text, explicit confirmation no חוב
value changed.
```

## עובדות שנמדדו (אל תבדוק שוב)

- הכלל "אפס בשני דוחות רצופים = שולם" שגוי — כפ"ס היה באפס שלושה דוחות וחזר.
- `קוד חברה` בלוח הסניפים ≠ `קוד לקוח` החשבונאי. 0% התאמה.
- סינון/כתיבת סטטוס במאנדיי לפי טקסט נדחה — אינדקס מספרי בלבד.
- אימות מהרצת 23.08.26 (דוח 16.08): 26 מועמדי-גיול גולמיים · הסקריפט סיווג
  15 נקיים / 8 חשודים / 3 כבר-בשלב, וההצלבה מול קובץ ההחזרים פסלה בין השאר
  את שייפ פורום (11 חזרות הו"ק) שהגיול הציג כ"שוטף".
