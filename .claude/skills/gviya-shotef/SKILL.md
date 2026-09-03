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
other inputs — BUT classify per invoice, never per client from the oldest
month alone. A client is "in shotef" only when **not a single shekel** of
their debt is past its due date (end of the invoice's month + N shotef days).
Chen's ruling, 01.09: "אין צדק ואין כלום — כל מי שבחוב, גם על שקל." No grace
days, no amount floor. Earlier drafts with a 7-day grace and a ₪1,000 floor
were rejected.

## When this runs

**Every weekly sync, without being asked.** Chen's standing instruction (30.08):
the shotef check is part of processing each new aging report, not a separate
request. Order of a weekly run: sync the report diff → run this check in both
directions → present both to Chen → apply after approval. Skipping it lets
not-yet-due debt sit in the reps' funnel and, worse, lets debt-notice campaigns
target clients whose invoices are not due — that happened on 23.08 and had to
be undone.

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
   - ⏰ in the stage with any overdue shekel → move back to `לקוח בחוב` so it
     re-enters the funnel. Standing practice since 30.08.
   - **Precedence rule (01.09):** the check moves cards only between
     `לקוח בחוב` ↔ `חוב בתחום שוטף`. A card a rep or another rule placed
     elsewhere (`המתנה לתשלום`, `פניה ראשונה`, `בחינת הנהלה`…) is never
     touched, even if none of its debt is due — the stage was chosen for a
     reason (e.g. partial payment = active conversation). Example that forced
     the rule: פלאקארד, ₪25k not yet due, sitting in המתנה לתשלום after paying
     half.
   - Set `מצב יתרה` to `תחום שוטף` for every card in the stage — the column
     otherwise says "יש חוב" and reps read it as a contradiction (Yasmin, 30.08).

3. **Show Chen, then stop — every run, including repeats.** Two cautions to
   apply before he approves:
   - **Recycled debt (HANDOFF §3):** a bounced הו"ק or credit-and-reinvoice
     lands in the current month, so משך חוב looks fresh while the debt is old.
     Verified live: one board pass surfaced 5 "candidates" that were all known
     chronic bouncers. If a candidate is a known bouncer (Yasmin's chargeback
     file, card updates mentioning החזרים, or a suspiciously fresh משך חוב on
     a client who has been in debt for weeks) — point it out and let Chen
     decide. Never silently move such a client.
   - **Stale board fields:** until 01.09 the sync refreshed משך חוב / ימי חוב /
     עודכן מדוח only on cards whose amount changed, so the board's oldest-month
     field could lag by weeks (נווה שרת showed 02/26 while the report said
     07/26). Since 01.09 every card is refreshed on every sync — but when the
     aging XLSX is at hand, classify from its buckets, not from the board.
   - **Reissued invoices:** a bucket can vanish between reports (נווה אליעזר's
     "May" invoice was gone a day later). When a rep disputes a date, trust
     the newest report and, if still disputed, ask finance at invoice level.
   - **Negative buckets:** the sync sets משך חוב from any nonzero bucket,
     including negative ones, so a card can show an ancient משך חוב because of
     an old credit.

4. **Apply in the background** after explicit approval, using the prompt
   template in `references/context.md`. Chen may drop or add clients first —
   edit `shotef_apply.json` accordingly. The agent posts each card's `update`
   verbatim, moves the stage by **index** (`{"color_mm69ws8c": {"index": 10}}`
   — text labels are rejected), touches nothing else, and verifies debt values
   are unchanged. Report its final table to Chen.

## Hard rules

- Writes allowed: `שלב` and `מצב יתרה` (label `תחום שוטף`). Never touch `חוב`
  (numeric_mm69dbtr), `נציגה`, or any text column.
- Terms: "שוטף N"/"ש+N" → due N days after invoice-month end; הו"ק and empty
  terms → due at month end (N=0); empty terms are flagged ⚠ in the output.
- משך חוב of "…(שנה קודמת)" or "שנים קודמות" is overdue by definition.
- Classification output goes to Chen before any write, every time.
