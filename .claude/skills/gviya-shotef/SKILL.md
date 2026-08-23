---
name: gviya-shotef
description: >-
  Classify Fizikal collections clients whose entire debt is still within their
  payment-terms window (שוטף 30/45/60/90, הו"ק) and move them to the "חוב בתחום
  שוטף" stage on the gviya monday board (5102318900), after guarding against
  recycled/re-issued debt that only looks current. Use this whenever Chen asks
  about "תחום שוטף", "חוב בתחום שוטף", "מי לא באמת בחוב", "מי בשוטף", who is
  within payment terms, cleaning the collections funnel of not-yet-overdue
  clients, or asks to re-run the shotef classification on a new aging report —
  even if he doesn't name the skill. Any request to separate real overdue debt
  from debt that is simply not due yet on the gviya board belongs here.
---

# gviya-shotef · סיווג "חוב בתחום שוטף"

## What this does and why

The finance aging report lists every client with a balance, but a balance is
not the same as overdue debt: a client on שוטף-60 terms whose whole balance
sits in the last two months is simply not due yet. Leaving those clients in
`לקוח בחוב` buries the reps' funnel in ~₪200k+ of non-work. This skill computes
who is genuinely within terms, guards against the one trap that breaks the
math (recycled debt — see below), shows Chen the lists, and only after his
explicit approval moves cards to the `חוב בתחום שוטף` stage with a documented
note on each card.

**The trap (HANDOFF §3, it already cost us once):** debt that bounced (הו"ק
חוזרת) or was credited-and-reinvoiced moves into the current-month bucket of
the aging report, so it *looks* within terms while actually being old unpaid
debt. The script detects this from snapshot history (zero-and-return with the
same amount; overdue dropping while total doesn't) and from Yasmin's
chargeback file, and quarantines those clients into a "suspect" list that
requires Chen's decision. Never silently promote a suspect to the clean list.

## Workflow

Read `references/context.md` first — board ids, column ids, stage indexes,
data locations, and the apply-agent prompt template live there.

1. **Gather inputs.**
   - Latest aging XLSX: Chen usually uploads it; it also arrives by email
     (Outlook MCP, sender YasminB@movement-group.com, subject contains
     "גיול לקוחות חייבים"). Default sheet = the last one in the workbook.
   - Board state: fetch all items from board 5102318900 via monday MCP
     `get_board_items_page` (columns: `text_mm69gg7g` code, `color_mm69ws8c`
     stage, `numeric_mm69dbtr` debt) and dump to JSON keyed by client code:
     `{code: {"id","name","stage","debt"}}`. The result is large — it will be
     saved to a file; parse it with a script, don't read it into context.
   - Snapshot history directory and, if available, Yasmin's chargeback XLSX
     ("סטטוס הוק חוזרות"). Run without them only if truly unavailable — the
     script then prints explicit ⚠ warnings; relay those warnings to Chen.

2. **Classify — read-only.**
   ```bash
   python3 scripts/classify_shotef.py --xlsx <aging.xlsx> --board <board.json> \
       --history <snapshots-dir> --returns <chargebacks.xlsx> --out shotef_apply.json
   ```
   The script never writes to monday. It prints three lists (clean / suspect /
   already-in-stage) and writes `shotef_apply.json` for the clean movers.

3. **Show Chen, then stop.** Present the clean list (with per-client oldest
   invoice and due date), the suspect list with the reason for each flag, and
   the already-in-stage count as validation. Flags cite dates — apply judgment:
   a recycling event from months ago on freshly-accrued debt may be stale, and
   you may say so, but the decision to move a suspect is always Chen's.
   **Do not touch monday until he approves.** He may include or exclude
   specific clients; edit `shotef_apply.json` accordingly (drop or add entries
   — for an added suspect, build its entry with the same fields and an update
   note that records the flag and his decision).

4. **Apply in the background.** Launch a background agent with the prompt
   template in `references/context.md`. It posts each card's `update` verbatim,
   moves the stage by **index** (`{"color_mm69ws8c": {"index": 10}}` — text
   labels are rejected), touches nothing else, and verifies at the end that
   debt values are unchanged. Report its final table back to Chen.

## Hard rules

- The stage move is the ONLY column written. Never touch `חוב`
  (numeric_mm69dbtr), `מצב יתרה` (color_mm69xcy2 — sync-owned), `נציגה`, or any
  text column.
- Terms parsing: "שוטף N"/"ש+N" → due N days after invoice-month end; הו"ק and
  empty terms → due at month end (N=0). Empty terms are flagged ⚠ in output.
- Anything in last-year or prior-years aging columns is overdue by definition.
- Dry-run first is not optional: classification output goes to Chen before any
  write, every time — including re-runs on a new report.
- Cards already in `חוב בתחום שוטף` whose debt has since gone overdue are NOT
  handled here — mention them to Chen if noticed, but reversing the stage is
  the reps' call.
