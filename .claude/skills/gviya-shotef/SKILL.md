---
name: gviya-shotef
description: >-
  Walk the Fizikal gviya monday board (5102318900) and check which debtor
  clients are still within their payment-terms window (שוטף 30/45/60/90, הו"ק)
  — computed from the board's own משך חוב and תנאי תשלום columns — then, after
  Chen's approval, move them to the "חוב בתחום שוטף" stage. Use this whenever
  Chen asks about "תחום שוטף", "חוב בתחום שוטף", "מי בשוטף", "מי לא באמת
  בחוב", who is within payment terms, or cleaning the collections funnel of
  not-yet-overdue clients — even if he doesn't name the skill. Any request to
  separate real overdue debt from debt that is simply not due yet on the gviya
  board belongs here.
---

# gviya-shotef · מי בלוח נמצא בתחום השוטף

## What this does and why

A balance is not overdue debt: a client on שוטף-60 whose oldest open invoice
is from last month is simply not due yet, and leaving them in `לקוח בחוב`
buries the reps' funnel in non-work. The board already carries everything
needed — the sync stamps `משך חוב` (oldest month carrying money) and
`תנאי תשלום` on every card — so this is a pure board pass: no aging file, no
other inputs. The debt is within terms when (end of the oldest invoice's
month + N shotef days) has not passed yet.

## Workflow

`references/context.md` has the column ids, stage indexes, and the
apply-agent prompt template.

1. **Fetch the board.** monday MCP `get_board_items_page`, board 5102318900,
   all items, columns `text_mm69gg7g` (code), `color_mm69ws8c` (stage),
   `numeric_mm69dbtr` (debt), `text_mm69ew8e` (terms), `text_mm6970hn`
   (משך חוב), `date_mm69pvs6` (עודכן מדוח). The result is large and lands in
   a file — convert it with a script to `{code: {"id","name","stage","debt",
   "terms","age","report"}}`; don't read it into context.

2. **Classify — read-only.**
   ```bash
   python3 scripts/classify_shotef.py --board board_full.json --out shotef_apply.json
   ```
   Prints three lists and writes the move plan; never writes to monday:
   - ✅ within terms and not yet in the stage → candidates
   - ✔ already in the stage and still within terms → validation only
   - ⏰ in the stage but now past due → surface to Chen; moving a card *back*
     to the funnel is his or the reps' call, never automatic

3. **Show Chen, then stop — every run, including repeats.** Two cautions to
   apply before he approves:
   - **Recycled debt (HANDOFF §3):** a bounced הו"ק or credit-and-reinvoice
     lands in the current month, so משך חוב looks fresh while the debt is old.
     Verified live: one board pass surfaced 5 "candidates" that were all known
     chronic bouncers. If a candidate is a known bouncer (Yasmin's chargeback
     file, card updates mentioning החזרים, or a suspiciously fresh משך חוב on
     a client who has been in debt for weeks) — point it out and let Chen
     decide. Never silently move such a client.
   - **Negative buckets:** the sync sets משך חוב from any nonzero bucket,
     including negative ones, so a card can show an ancient משך חוב because of
     an old credit. Mainly affects the ⏰ list — mention it when relevant.

4. **Apply in the background** after explicit approval, using the prompt
   template in `references/context.md`. Chen may drop or add clients first —
   edit `shotef_apply.json` accordingly. The agent posts each card's `update`
   verbatim, moves the stage by **index** (`{"color_mm69ws8c": {"index": 10}}`
   — text labels are rejected), touches nothing else, and verifies debt values
   are unchanged. Report its final table to Chen.

## Hard rules

- The stage move is the ONLY write. Never touch `חוב` (numeric_mm69dbtr),
  `מצב יתרה` (color_mm69xcy2 — sync-owned), `נציגה`, or any text column.
- Terms: "שוטף N"/"ש+N" → due N days after invoice-month end; הו"ק and empty
  terms → due at month end (N=0); empty terms are flagged ⚠ in the output.
- משך חוב of "…(שנה קודמת)" or "שנים קודמות" is overdue by definition.
- Classification output goes to Chen before any write, every time.
