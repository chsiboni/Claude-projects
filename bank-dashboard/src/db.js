'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const Database = require('better-sqlite3');
const config = require('./config');

fs.mkdirSync(path.dirname(config.dbPath), { recursive: true });

const db = new Database(config.dbPath);
db.pragma('journal_mode = WAL');

db.exec(`
  CREATE TABLE IF NOT EXISTS transactions (
    id                TEXT PRIMARY KEY,
    source            TEXT NOT NULL,
    source_label      TEXT NOT NULL,
    source_kind       TEXT NOT NULL,
    account           TEXT NOT NULL,
    date              TEXT NOT NULL,
    processed_date    TEXT,
    amount            REAL NOT NULL,
    original_amount   REAL,
    original_currency TEXT,
    description       TEXT NOT NULL,
    memo              TEXT,
    type              TEXT,
    status            TEXT,
    installment_num   INTEGER,
    installment_total INTEGER,
    category          TEXT,
    category_source   TEXT NOT NULL DEFAULT 'rule',
    first_seen        TEXT NOT NULL,
    updated_at        TEXT NOT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_txn_date     ON transactions(date);
  CREATE INDEX IF NOT EXISTS idx_txn_source   ON transactions(source);
  CREATE INDEX IF NOT EXISTS idx_txn_category ON transactions(category);

  CREATE TABLE IF NOT EXISTS balances (
    source       TEXT NOT NULL,
    source_label TEXT NOT NULL,
    account      TEXT NOT NULL,
    balance      REAL,
    currency     TEXT,
    as_of        TEXT NOT NULL,
    PRIMARY KEY (source, account)
  );

  CREATE TABLE IF NOT EXISTS future_debits (
    id           TEXT PRIMARY KEY,
    source       TEXT NOT NULL,
    source_label TEXT NOT NULL,
    account      TEXT,
    charge_date  TEXT,
    amount       REAL NOT NULL,
    currency     TEXT,
    updated_at   TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS runs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at  TEXT NOT NULL,
    finished_at TEXT,
    status      TEXT NOT NULL,
    added       INTEGER NOT NULL DEFAULT 0,
    updated     INTEGER NOT NULL DEFAULT 0,
    detail      TEXT
  );

  -- Manual corrections. A pattern here always beats the automatic rules.
  CREATE TABLE IF NOT EXISTS category_overrides (
    pattern    TEXT PRIMARY KEY,
    category   TEXT NOT NULL,
    created_at TEXT NOT NULL
  );
`);

/**
 * A stable id for a transaction so re-scraping the same window is idempotent.
 * Prefers the bank's own identifier (asmachta) when there is one; otherwise
 * falls back to the fields that together identify a charge.
 */
function transactionId(sourceKey, account, txn) {
  const seed =
    txn.identifier != null && txn.identifier !== ''
      ? `${sourceKey}|${account}|${txn.identifier}|${txn.installments?.number ?? ''}`
      : [
          sourceKey,
          account,
          (txn.date || '').slice(0, 10),
          txn.chargedAmount,
          (txn.description || '').trim(),
          txn.installments?.number ?? '',
        ].join('|');
  return crypto.createHash('sha1').update(seed).digest('hex');
}

const insertTxn = db.prepare(`
  INSERT INTO transactions (
    id, source, source_label, source_kind, account, date, processed_date,
    amount, original_amount, original_currency, description, memo, type, status,
    installment_num, installment_total, category, category_source, first_seen, updated_at
  ) VALUES (
    @id, @source, @source_label, @source_kind, @account, @date, @processed_date,
    @amount, @original_amount, @original_currency, @description, @memo, @type, @status,
    @installment_num, @installment_total, @category, 'rule', @now, @now
  )
  ON CONFLICT(id) DO UPDATE SET
    date           = excluded.date,
    processed_date = excluded.processed_date,
    amount         = excluded.amount,
    status         = excluded.status,
    memo           = excluded.memo,
    -- a manual category is never overwritten by a re-scrape
    category       = CASE WHEN transactions.category_source = 'manual'
                          THEN transactions.category ELSE excluded.category END,
    updated_at     = excluded.updated_at
  WHERE transactions.status  IS NOT excluded.status
     OR transactions.amount  IS NOT excluded.amount
     OR transactions.date    IS NOT excluded.date
`);

const upsertBalance = db.prepare(`
  INSERT INTO balances (source, source_label, account, balance, currency, as_of)
  VALUES (@source, @source_label, @account, @balance, @currency, @as_of)
  ON CONFLICT(source, account) DO UPDATE SET
    balance = excluded.balance, currency = excluded.currency, as_of = excluded.as_of
`);

const upsertFutureDebit = db.prepare(`
  INSERT INTO future_debits (id, source, source_label, account, charge_date, amount, currency, updated_at)
  VALUES (@id, @source, @source_label, @account, @charge_date, @amount, @currency, @updated_at)
  ON CONFLICT(id) DO UPDATE SET
    amount = excluded.amount, currency = excluded.currency, updated_at = excluded.updated_at
`);

/** Returns {added, updated} for the batch. */
function saveTransactions(rows) {
  let added = 0;
  let updated = 0;
  const run = db.transaction((batch) => {
    for (const row of batch) {
      const before = db.prepare('SELECT 1 FROM transactions WHERE id = ?').get(row.id);
      const result = insertTxn.run(row);
      if (!before) added += 1;
      else if (result.changes > 0) updated += 1;
    }
  });
  run(rows);
  return { added, updated };
}

function saveBalances(rows) {
  db.transaction((batch) => batch.forEach((row) => upsertBalance.run(row)))(rows);
}

function saveFutureDebits(sourceKey, rows) {
  db.transaction((batch) => {
    db.prepare('DELETE FROM future_debits WHERE source = ?').run(sourceKey);
    batch.forEach((row) => upsertFutureDebit.run(row));
  })(rows);
}

/**
 * The date to resume from: a little before the newest stored transaction so
 * pending charges that later settle are picked up again, or the initial window
 * on a cold database.
 */
function resumeDate(sourceKey) {
  const row = db.prepare('SELECT MAX(date) AS latest FROM transactions WHERE source = ?').get(sourceKey);
  const start = new Date();
  if (row?.latest) {
    const latest = new Date(row.latest);
    latest.setDate(latest.getDate() - config.overlapDays);
    return latest;
  }
  start.setMonth(start.getMonth() - config.initialMonthsBack);
  return start;
}

function startRun() {
  const info = db
    .prepare(`INSERT INTO runs (started_at, status) VALUES (?, 'running')`)
    .run(new Date().toISOString());
  return info.lastInsertRowid;
}

function finishRun(id, { status, added = 0, updated = 0, detail = null }) {
  db.prepare(
    `UPDATE runs SET finished_at = ?, status = ?, added = ?, updated = ?, detail = ? WHERE id = ?`
  ).run(new Date().toISOString(), status, added, updated, detail, id);
}

function lastRun() {
  return db.prepare('SELECT * FROM runs ORDER BY id DESC LIMIT 1').get() || null;
}

module.exports = {
  db,
  transactionId,
  saveTransactions,
  saveBalances,
  saveFutureDebits,
  resumeDate,
  startRun,
  finishRun,
  lastRun,
};
