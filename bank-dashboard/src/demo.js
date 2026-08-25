'use strict';

/**
 * Fills the database with a year of plausible-looking activity so the dashboard
 * can be explored before any real credentials are wired up.
 * Run with:  npm run demo   (writes to data/demo.db, never to your real data)
 */

process.env.DB_PATH = process.env.DB_PATH || require('path').join(__dirname, '..', 'data', 'demo.db');

const store = require('./db');
const { categorize } = require('./categorize');
const { db } = store;

const FIXED = [
  ['משכנתא בנק הפועלים', 5400, 1],
  ['סלקום', 129, 8],
  ['חברת החשמל', 480, 12],
  ['NETFLIX.COM', 54.9, 14],
  ['SPOTIFY P2C4F', 21.9, 17],
  ['הולמס פלייס', 249, 3],
  ['ביטוח בריאות הראל', 312, 5],
  ['ארנונה עיריית תל אביב', 690, 10],
  ['ועד בית', 220, 2],
  ['OPENAI *CHATGPT', 78.4, 21],
];

const VARIABLE = [
  ['שופרסל דיל', 'סופרמרקט', 180, 520],
  ['רמי לוי שיווק', 'סופרמרקט', 150, 480],
  ['ארומה תל אביב', 'מסעדות ובתי קפה', 22, 62],
  ['WOLT', 'מסעדות ובתי קפה', 55, 140],
  ['פז יעל', 'דלק ותחבורה', 190, 320],
  ['סופר פארם', 'בריאות ותרופות', 45, 220],
  ['פנגו חניה', 'דלק ותחבורה', 8, 35],
  ['TERMINAL X', 'קניות וביגוד', 120, 480],
  ['איקאה נתניה', 'בית וריהוט', 90, 700],
];

/** Deterministic pseudo-random, so a demo run is reproducible. */
let seed = 20260825;
function rand() {
  seed = (seed * 1103515245 + 12345) % 2147483648;
  return seed / 2147483648;
}
function between(min, max) {
  return Math.round((min + rand() * (max - min)) * 100) / 100;
}
function iso(year, month, day) {
  return new Date(Date.UTC(year, month, Math.min(day, 28))).toISOString().slice(0, 10);
}

function build() {
  const now = new Date();
  const rows = [];
  const stamp = now.toISOString();

  const push = (date, description, amount, source, label, kind, account) =>
    rows.push({
      id: store.transactionId(source, account, { date, chargedAmount: amount, description }),
      source,
      source_label: label,
      source_kind: kind,
      account,
      date,
      processed_date: date,
      amount,
      original_amount: amount,
      original_currency: 'ILS',
      description,
      memo: null,
      type: 'normal',
      status: 'completed',
      installment_num: null,
      installment_total: null,
      category: categorize(description, amount),
      now: stamp,
    });

  for (let back = 11; back >= 0; back -= 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - back, 1);
    const year = date.getFullYear();
    const month = date.getMonth();
    const partial = back === 0 ? now.getDate() / 31 : 1;

    push(iso(year, month, 9), 'משכורת פיזיקל בע"מ', between(21000, 23000), 'hapoalim', 'בנק הפועלים', 'bank', '12-345-678901');

    for (const [name, amount, day] of FIXED) {
      if (day / 31 > partial) continue;
      // Netflix quietly gets more expensive halfway through the year.
      const bump = name.startsWith('NETFLIX') && back <= 3 ? 1.24 : 1;
      const jitter = name.includes('חשמל') || name.includes('ארנונה') ? between(0.82, 1.24) : 1;
      push(iso(year, month, day), name, -Math.round(amount * bump * jitter * 100) / 100, 'hapoalim', 'בנק הפועלים', 'bank', '12-345-678901');
    }

    for (const [name, , min, max] of VARIABLE) {
      const times = Math.round(between(2, 7) * partial);
      for (let i = 0; i < times; i += 1) {
        push(iso(year, month, Math.round(between(1, 28))), name, -between(min, max), 'max', 'מקס', 'card', '****4417');
      }
    }

    // One memorable outlier and one accidental double charge.
    if (back === 2) push(iso(year, month, 16), 'איקאה נתניה', -3890, 'max', 'מקס', 'card', '****4417');
    if (back === 1) {
      push(iso(year, month, 6), 'WOLT', -128.5, 'max', 'מקס', 'card', '****4417');
      push(iso(year, month, 7), 'WOLT', -128.5, 'max', 'מקס', 'card', '****4417');
    }

    push(iso(year, month, 2), 'חיוב כרטיס אשראי מקס', -between(3200, 5200), 'hapoalim', 'בנק הפועלים', 'bank', '12-345-678901');
    push(iso(year, month, 20), 'עמלת ערוץ ישיר', -between(8, 14), 'hapoalim', 'בנק הפועלים', 'bank', '12-345-678901');
  }

  return rows;
}

db.exec('DELETE FROM transactions; DELETE FROM balances; DELETE FROM future_debits; DELETE FROM runs;');
const result = store.saveTransactions(build());
store.saveBalances([
  {
    source: 'hapoalim',
    source_label: 'בנק הפועלים',
    account: '12-345-678901',
    balance: 38420.55,
    currency: 'ILS',
    as_of: new Date().toISOString(),
  },
]);
store.saveFutureDebits('max', [
  {
    id: 'max|demo|next',
    source: 'max',
    source_label: 'מקס',
    account: '12-345-678901',
    charge_date: new Date(new Date().getFullYear(), new Date().getMonth() + 1, 2).toISOString().slice(0, 10),
    amount: -4180.3,
    currency: 'ILS',
    updated_at: new Date().toISOString(),
  },
]);
const runId = store.startRun();
store.finishRun(runId, { status: 'ok', added: result.added, updated: 0, detail: 'נתוני הדגמה' });

console.log(`נוצרו ${result.added} תנועות הדגמה ב־${process.env.DB_PATH}`);
console.log('להרצה מול ההדגמה:  DB_PATH=data/demo.db npm start');
