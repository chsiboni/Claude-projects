'use strict';

const { createScraper } = require('israeli-bank-scrapers');
const config = require('./config');
const store = require('./db');
const { categorize } = require('./categorize');

/** Never let a credential reach a log line. */
function log(...args) {
  console.log(`[${new Date().toISOString().slice(11, 19)}]`, ...args);
}

function toRow(source, account, txn, now) {
  const description = (txn.description || '').trim();
  return {
    id: store.transactionId(source.key, account.accountNumber, txn),
    source: source.key,
    source_label: source.label,
    source_kind: source.kind,
    account: account.accountNumber,
    date: (txn.date || txn.processedDate || '').slice(0, 10),
    processed_date: (txn.processedDate || '').slice(0, 10) || null,
    amount: Number(txn.chargedAmount),
    original_amount: txn.originalAmount != null ? Number(txn.originalAmount) : null,
    original_currency: txn.originalCurrency || null,
    description,
    memo: txn.memo || null,
    type: txn.type || null,
    status: txn.status || null,
    installment_num: txn.installments?.number ?? null,
    installment_total: txn.installments?.total ?? null,
    category: categorize(description, Number(txn.chargedAmount)),
    now,
  };
}

async function collectSource(source) {
  const startDate = store.resumeDate(source.key);
  log(`${source.label}: מתחבר, מושך מ־${startDate.toISOString().slice(0, 10)}`);

  const scraper = createScraper({
    companyId: source.companyId,
    startDate,
    showBrowser: config.showBrowser,
    timeout: config.timeoutMs,
    defaultTimeout: config.timeoutMs,
    combineInstallments: false,
    verbose: false,
    args: ['--no-sandbox', '--disable-dev-shm-usage'],
  });

  const result = await scraper.scrape(source.creds);
  if (!result.success) {
    throw new Error(`${result.errorType || 'שגיאה'}: ${result.errorMessage || 'ללא פירוט'}`);
  }

  const now = new Date().toISOString();
  const rows = [];
  const balances = [];

  for (const account of result.accounts || []) {
    for (const txn of account.txns || []) rows.push(toRow(source, account, txn, now));
    if (account.balance != null) {
      balances.push({
        source: source.key,
        source_label: source.label,
        account: account.accountNumber,
        balance: Number(account.balance),
        currency: account.currency || 'ILS',
        as_of: account.balanceDate || now,
      });
    }
  }

  const { added, updated } = store.saveTransactions(rows);
  if (balances.length) store.saveBalances(balances);

  store.saveFutureDebits(
    source.key,
    (result.futureDebits || []).map((debit, index) => ({
      id: `${source.key}|${debit.bankAccountNumber || 'na'}|${debit.chargeDate || index}`,
      source: source.key,
      source_label: source.label,
      account: debit.bankAccountNumber || null,
      charge_date: debit.chargeDate ? String(debit.chargeDate).slice(0, 10) : null,
      amount: Number(debit.amount),
      currency: debit.amountCurrency || 'ILS',
      updated_at: now,
    }))
  );

  log(`${source.label}: ${rows.length} תנועות (${added} חדשות, ${updated} עודכנו)`);
  return { added, updated };
}

/** Runs every configured source; one failure never cancels the others. */
async function collectAll() {
  const sources = config.activeSources();
  if (!sources.length) {
    log('לא הוגדר אף מקור ב־.env — אין מה לאסוף.');
    return { added: 0, updated: 0, failures: [] };
  }

  const runId = store.startRun();
  let added = 0;
  let updated = 0;
  const failures = [];

  for (const source of sources) {
    try {
      const result = await collectSource(source);
      added += result.added;
      updated += result.updated;
    } catch (error) {
      log(`${source.label}: נכשל — ${error.message}`);
      failures.push(`${source.label}: ${error.message}`);
    }
  }

  store.finishRun(runId, {
    status: failures.length === sources.length ? 'failed' : failures.length ? 'partial' : 'ok',
    added,
    updated,
    detail: failures.join(' | ') || null,
  });

  log(`סיום: ${added} תנועות חדשות, ${updated} עודכנו${failures.length ? `, ${failures.length} מקורות נכשלו` : ''}`);
  return { added, updated, failures };
}

module.exports = { collectAll };

if (require.main === module) {
  collectAll()
    .then((result) => process.exit(result.failures.length ? 1 : 0))
    .catch((error) => {
      log('שגיאה כללית:', error.message);
      process.exit(1);
    });
}
