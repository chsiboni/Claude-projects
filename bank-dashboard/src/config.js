'use strict';

const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

/**
 * Every source is optional — define only the ones you actually have in .env.
 * Credentials never leave this process: they are read from the local .env and
 * handed straight to the scraper. Nothing is written to the database.
 */
const SOURCES = [
  {
    key: 'hapoalim',
    companyId: 'hapoalim',
    label: 'בנק הפועלים',
    kind: 'bank',
    credentials: () => ({
      userCode: process.env.POALIM_USER_CODE,
      password: process.env.POALIM_PASSWORD,
    }),
  },
  {
    key: 'max',
    companyId: 'max',
    label: 'מקס',
    kind: 'card',
    credentials: () => ({
      username: process.env.MAX_USERNAME,
      password: process.env.MAX_PASSWORD,
    }),
  },
  {
    key: 'visaCal',
    companyId: 'visaCal',
    label: 'כאל',
    kind: 'card',
    credentials: () => ({
      username: process.env.CAL_USERNAME,
      password: process.env.CAL_PASSWORD,
    }),
  },
  {
    key: 'isracard',
    companyId: 'isracard',
    label: 'ישראכרט',
    kind: 'card',
    credentials: () => ({
      id: process.env.ISRACARD_ID,
      card6Digits: process.env.ISRACARD_CARD6,
      password: process.env.ISRACARD_PASSWORD,
    }),
  },
  {
    key: 'amex',
    companyId: 'amex',
    label: 'אמריקן אקספרס',
    kind: 'card',
    credentials: () => ({
      id: process.env.AMEX_ID,
      card6Digits: process.env.AMEX_CARD6,
      password: process.env.AMEX_PASSWORD,
    }),
  },
];

/** A source is active only when every one of its credential fields is filled in. */
function activeSources() {
  return SOURCES.map((source) => ({ ...source, creds: source.credentials() })).filter(
    (source) =>
      Object.values(source.creds).length > 0 &&
      Object.values(source.creds).every((value) => typeof value === 'string' && value.trim() !== '')
  );
}

const config = {
  port: Number(process.env.PORT) || 4000,
  /** Bound to loopback on purpose — this dashboard must never be reachable from the network. */
  host: process.env.HOST || '127.0.0.1',
  dbPath: process.env.DB_PATH || path.join(__dirname, '..', 'data', 'bank.db'),
  /** How far back to reach on the very first collection, in months. */
  initialMonthsBack: Number(process.env.INITIAL_MONTHS_BACK) || 12,
  /** Re-scan this many days before the newest stored transaction, to catch pending→completed flips. */
  overlapDays: Number(process.env.OVERLAP_DAYS) || 10,
  /** Daily auto-collection. Empty string disables the scheduler. */
  cron: process.env.COLLECT_CRON === '' ? null : process.env.COLLECT_CRON || '0 6 * * *',
  showBrowser: process.env.SHOW_BROWSER === '1',
  timeoutMs: Number(process.env.SCRAPE_TIMEOUT_MS) || 180000,
  activeSources,
  SOURCES,
};

module.exports = config;
