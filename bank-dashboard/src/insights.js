'use strict';

const { db } = require('./db');
const { NON_SPEND, CARD_SETTLEMENT } = require('./categorize');

const NON_SPEND_LIST = [...NON_SPEND];
const NON_SPEND_SQL = NON_SPEND_LIST.map(() => '?').join(',');

const DAY_MS = 24 * 60 * 60 * 1000;

function monthKey(isoDate) {
  return String(isoDate).slice(0, 7);
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function monthsAgo(count) {
  const date = new Date();
  date.setDate(1);
  date.setMonth(date.getMonth() - count);
  return date.toISOString().slice(0, 10);
}

/**
 * Collapses a statement line to the merchant behind it: strips installment
 * markers, card digits, dates and stray punctuation so the same subscription
 * groups together month over month.
 */
function merchantKey(description) {
  return String(description || '')
    .toLowerCase()
    .replace(/תשלום\s*\d+\s*(מתוך|מ)\s*\d+/g, '')
    .replace(/\d+\s*תשלומים?/g, '')
    .replace(/[0-9]{2,}/g, '')
    .replace(/[^\p{L}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

/** Median absolute deviation — robust to the outliers we are hunting for. */
function mad(values, center) {
  if (values.length < 2) return 0;
  return median(values.map((value) => Math.abs(value - center)));
}

function round(value) {
  return Math.round(Number(value) * 100) / 100;
}

/* ------------------------------------------------------------------ queries */

function allSpending(sinceIso) {
  return db
    .prepare(
      `SELECT * FROM transactions
        WHERE amount < 0 AND date >= ?
          AND category NOT IN (${NON_SPEND_SQL})
        ORDER BY date DESC`
    )
    .all(sinceIso, ...NON_SPEND_LIST);
}

function balances() {
  return db.prepare('SELECT * FROM balances ORDER BY source_label').all();
}

function hasCardSources() {
  return (
    db.prepare(`SELECT COUNT(*) AS n FROM transactions WHERE source_kind = 'card'`).get().n > 0
  );
}

/* ------------------------------------------------------------- monthly view */

/** Income and spending per month, newest last, for the trend chart. */
function monthlySeries(months = 12) {
  const rows = db
    .prepare(
      `SELECT substr(date, 1, 7) AS month,
              SUM(CASE WHEN amount > 0 AND category = 'הכנסה' THEN amount ELSE 0 END) AS income,
              SUM(CASE WHEN amount < 0 AND category NOT IN (${NON_SPEND_SQL}) THEN -amount ELSE 0 END) AS spend
         FROM transactions
        WHERE date >= ?
        GROUP BY month
        ORDER BY month`
    )
    .all(...NON_SPEND_LIST, monthsAgo(months));

  return rows.map((row) => ({
    month: row.month,
    income: round(row.income),
    spend: round(row.spend),
    net: round(row.income - row.spend),
  }));
}

function categoryBreakdown(month) {
  const target = month || monthKey(today());
  const rows = db
    .prepare(
      `SELECT category, SUM(-amount) AS total, COUNT(*) AS count
         FROM transactions
        WHERE amount < 0 AND substr(date, 1, 7) = ?
          AND category NOT IN (${NON_SPEND_SQL})
        GROUP BY category
        ORDER BY total DESC`
    )
    .all(target, ...NON_SPEND_LIST);

  const total = rows.reduce((sum, row) => sum + row.total, 0);
  return rows.map((row) => ({
    category: row.category,
    total: round(row.total),
    count: row.count,
    share: total ? round((row.total / total) * 100) : 0,
  }));
}

function topMerchants(month, limit = 12) {
  const target = month || monthKey(today());
  return db
    .prepare(
      `SELECT description, category, SUM(-amount) AS total, COUNT(*) AS count
         FROM transactions
        WHERE amount < 0 AND substr(date, 1, 7) = ?
          AND category NOT IN (${NON_SPEND_SQL})
        GROUP BY description
        ORDER BY total DESC
        LIMIT ?`
    )
    .all(target, ...NON_SPEND_LIST, limit)
    .map((row) => ({ ...row, total: round(row.total) }));
}

/* -------------------------------------------------------- recurring charges */

/**
 * A charge is "recurring" when the same merchant appears at least three times,
 * in at least three different months, at a roughly constant interval.
 */
function recurring({ lookbackMonths = 15 } = {}) {
  const rows = allSpending(monthsAgo(lookbackMonths));
  const groups = new Map();

  for (const row of rows) {
    // Instalment plans are one purchase split up, not a subscription.
    if (row.installment_total > 1) continue;
    const key = merchantKey(row.description);
    if (key.length < 3) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }

  const results = [];
  const now = Date.now();

  for (const [key, items] of groups) {
    if (items.length < 3) continue;
    const months = new Set(items.map((item) => monthKey(item.date)));
    if (months.size < 3) continue;

    const sorted = [...items].sort((a, b) => a.date.localeCompare(b.date));
    const gaps = sorted
      .slice(1)
      .map((item, index) => (new Date(item.date) - new Date(sorted[index].date)) / DAY_MS);
    const interval = median(gaps);

    let cadence = null;
    if (interval >= 24 && interval <= 38) cadence = 'חודשי';
    else if (interval >= 80 && interval <= 100) cadence = 'רבעוני';
    else if (interval >= 170 && interval <= 195) cadence = 'חצי שנתי';
    else if (interval >= 330 && interval <= 400) cadence = 'שנתי';
    if (!cadence) continue;

    const amounts = sorted.map((item) => Math.abs(item.amount));
    const typical = median(amounts);
    const spread = typical ? mad(amounts, typical) / typical : 0;
    const last = sorted[sorted.length - 1];
    const previous = median(amounts.slice(0, -1));
    const daysSinceLast = Math.round((now - new Date(last.date)) / DAY_MS);

    results.push({
      key,
      label: last.description,
      category: last.category,
      source: last.source_label,
      cadence,
      intervalDays: Math.round(interval),
      occurrences: sorted.length,
      typicalAmount: round(typical),
      lastAmount: round(Math.abs(last.amount)),
      lastDate: last.date,
      firstDate: sorted[0].date,
      // A steady amount reads as a subscription; a moving one as a variable bill.
      kind: spread <= 0.08 ? 'קבוע' : 'משתנה',
      // Flag a real price rise, not the wobble of a variable utility bill.
      priceJumpPct:
        previous > 0 && spread <= 0.2 ? round(((Math.abs(last.amount) - previous) / previous) * 100) : 0,
      // Past 1.6 intervals with no charge, it probably ended.
      dormant: daysSinceLast > interval * 1.6,
      daysSinceLast,
      monthlyEquivalent: round((typical * 30.44) / interval),
    });
  }

  return results.sort((a, b) => b.monthlyEquivalent - a.monthlyEquivalent);
}

/* ---------------------------------------------------------------- anomalies */

/**
 * Three kinds of thing worth a look: a charge far above what this category
 * normally costs, the same charge twice within days, and a recurring item
 * whose price went up.
 */
function anomalies({ lookbackMonths = 6 } = {}) {
  const rows = allSpending(monthsAgo(lookbackMonths));
  const found = [];

  const byCategory = new Map();
  for (const row of rows) {
    if (!byCategory.has(row.category)) byCategory.set(row.category, []);
    byCategory.get(row.category).push(row);
  }

  for (const [category, items] of byCategory) {
    if (items.length < 8) continue;
    const amounts = items.map((item) => Math.abs(item.amount));
    const center = median(amounts);
    const deviation = mad(amounts, center);
    if (deviation <= 0) continue;
    const threshold = center + 5 * deviation;

    for (const item of items) {
      const amount = Math.abs(item.amount);
      if (amount <= threshold || amount < 150) continue;
      found.push({
        kind: 'חריג',
        severity: amount > center + 9 * deviation ? 'critical' : 'serious',
        title: item.description,
        detail: `גבוה מהרגיל בקטגוריית ${category} — הטיפוסי הוא כ־${round(center)} ₪`,
        amount: round(amount),
        date: item.date,
        source: item.source_label,
      });
    }
  }

  const seen = new Map();
  for (const row of [...rows].sort((a, b) => a.date.localeCompare(b.date))) {
    const key = `${merchantKey(row.description)}|${Math.abs(row.amount).toFixed(2)}`;
    const previous = seen.get(key);
    const gapDays = previous ? (new Date(row.date) - new Date(previous.date)) / DAY_MS : Infinity;
    if (previous && gapDays <= 3 && Math.abs(row.amount) >= 40) {
      found.push({
        kind: 'כפילות',
        severity: 'warning',
        title: row.description,
        detail: `אותו סכום בדיוק חויב פעמיים תוך ${Math.round(gapDays)} ימים (${previous.date})`,
        amount: round(Math.abs(row.amount)),
        date: row.date,
        source: row.source_label,
      });
    }
    seen.set(key, row);
  }

  for (const item of recurring()) {
    if (item.dormant || item.priceJumpPct < 12) continue;
    found.push({
      kind: 'עליית מחיר',
      severity: item.priceJumpPct >= 30 ? 'serious' : 'warning',
      title: item.label,
      detail: `עלה ב־${item.priceJumpPct}% — היה כ־${item.typicalAmount} ₪, כעת ${item.lastAmount} ₪`,
      amount: item.lastAmount,
      date: item.lastDate,
      source: item.source,
    });
  }

  const order = { critical: 0, serious: 1, warning: 2 };
  return found
    .sort((a, b) => order[a.severity] - order[b.severity] || b.date.localeCompare(a.date))
    .slice(0, 40);
}

/* ----------------------------------------------------------------- forecast */

/**
 * Where the month is heading: what is already committed (recurring items not
 * yet charged, plus card debits the issuer has already scheduled) on top of
 * the run rate of everyday spending.
 */
function forecast() {
  const now = new Date();
  const month = monthKey(today());
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const dayOfMonth = now.getDate();
  const daysLeft = Math.max(0, daysInMonth - dayOfMonth);

  const spentSoFar = round(
    db
      .prepare(
        `SELECT COALESCE(SUM(-amount), 0) AS total FROM transactions
          WHERE amount < 0 AND substr(date, 1, 7) = ? AND category NOT IN (${NON_SPEND_SQL})`
      )
      .get(month, ...NON_SPEND_LIST).total
  );

  const recurringItems = recurring().filter((item) => !item.dormant && item.cadence === 'חודשי');
  const stillDue = recurringItems.filter((item) => monthKey(item.lastDate) !== month);
  const committed = round(stillDue.reduce((sum, item) => sum + item.typicalAmount, 0));

  const scheduled = round(
    db.prepare(`SELECT COALESCE(SUM(amount), 0) AS total FROM future_debits`).get().total
  );

  // Everyday spending only — the recurring items are already counted above.
  const recurringKeys = new Set(recurringItems.map((item) => item.key));
  const history = allSpending(monthsAgo(3)).filter(
    (row) => !recurringKeys.has(merchantKey(row.description))
  );
  const historyDays = Math.max(1, (Date.now() - new Date(monthsAgo(3))) / DAY_MS);
  const dailyVariable = history.reduce((sum, row) => sum + Math.abs(row.amount), 0) / historyDays;
  const variableLeft = round(dailyVariable * daysLeft);

  const income = round(
    db
      .prepare(
        `SELECT COALESCE(SUM(amount), 0) AS total FROM transactions
          WHERE amount > 0 AND category = 'הכנסה' AND substr(date, 1, 7) = ?`
      )
      .get(month).total
  );

  const currentBalance = round(
    balances()
      .filter((row) => row.balance != null)
      .reduce((sum, row) => sum + row.balance, 0)
  );

  const projectedSpend = round(spentSoFar + committed + variableLeft);

  return {
    month,
    dayOfMonth,
    daysInMonth,
    daysLeft,
    spentSoFar,
    committed,
    scheduledCardDebits: Math.abs(scheduled),
    variableLeft,
    projectedSpend,
    incomeSoFar: income,
    currentBalance,
    projectedBalance: round(currentBalance - committed - variableLeft),
    dailyVariable: round(dailyVariable),
    stillDue: stillDue.map((item) => ({ label: item.label, amount: item.typicalAmount })),
  };
}

/* ------------------------------------------------------------------ summary */

function summary() {
  const thisMonth = monthKey(today());
  const previous = new Date();
  previous.setDate(1);
  previous.setMonth(previous.getMonth() - 1);
  const lastMonth = monthKey(previous.toISOString().slice(0, 10));

  const series = monthlySeries(13);
  const current = series.find((row) => row.month === thisMonth) || { income: 0, spend: 0, net: 0 };
  const prior = series.find((row) => row.month === lastMonth) || { income: 0, spend: 0, net: 0 };

  const closedMonths = series.filter((row) => row.month !== thisMonth).slice(-6);
  const avgSpend = closedMonths.length
    ? round(closedMonths.reduce((sum, row) => sum + row.spend, 0) / closedMonths.length)
    : 0;

  const recurringItems = recurring().filter((item) => !item.dormant);
  const fixedMonthly = round(
    recurringItems.reduce((sum, item) => sum + item.monthlyEquivalent, 0)
  );

  const settlements = round(
    db
      .prepare(
        `SELECT COALESCE(SUM(-amount), 0) AS total FROM transactions
          WHERE category = ? AND substr(date, 1, 7) = ?`
      )
      .get(CARD_SETTLEMENT, thisMonth).total
  );

  return {
    thisMonth,
    lastMonth,
    spend: current.spend,
    lastMonthSpend: prior.spend,
    spendChangePct: prior.spend ? round(((current.spend - prior.spend) / prior.spend) * 100) : 0,
    income: current.income,
    net: current.net,
    avgSpend,
    balances: balances(),
    totalBalance: round(balances().reduce((sum, row) => sum + (row.balance || 0), 0)),
    fixedMonthly,
    recurringCount: recurringItems.length,
    cardSettlementsThisMonth: settlements,
    hasCardSources: hasCardSources(),
    transactionCount: db.prepare('SELECT COUNT(*) AS n FROM transactions').get().n,
    coverage: db
      .prepare('SELECT MIN(date) AS first, MAX(date) AS last FROM transactions')
      .get(),
  };
}

/* -------------------------------------------------------------- transaction list */

function transactions({ month, category, source, search, limit = 200, offset = 0 } = {}) {
  const where = ['1 = 1'];
  const params = [];
  if (month) {
    where.push('substr(date, 1, 7) = ?');
    params.push(month);
  }
  if (category) {
    where.push('category = ?');
    params.push(category);
  }
  if (source) {
    where.push('source = ?');
    params.push(source);
  }
  if (search) {
    where.push('description LIKE ?');
    params.push(`%${search}%`);
  }

  const rows = db
    .prepare(
      `SELECT id, source_label, account, date, amount, description, memo, category, status,
              installment_num, installment_total
         FROM transactions
        WHERE ${where.join(' AND ')}
        ORDER BY date DESC, rowid DESC
        LIMIT ? OFFSET ?`
    )
    .all(...params, limit, offset);

  const total = db
    .prepare(`SELECT COUNT(*) AS n FROM transactions WHERE ${where.join(' AND ')}`)
    .get(...params).n;

  return { rows, total };
}

function availableMonths() {
  return db
    .prepare('SELECT DISTINCT substr(date, 1, 7) AS month FROM transactions ORDER BY month DESC')
    .all()
    .map((row) => row.month);
}

module.exports = {
  summary,
  monthlySeries,
  categoryBreakdown,
  topMerchants,
  recurring,
  anomalies,
  forecast,
  transactions,
  availableMonths,
  merchantKey,
};
