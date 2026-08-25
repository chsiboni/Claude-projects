'use strict';

const path = require('path');
const express = require('express');
const cron = require('node-cron');

const config = require('./config');
const store = require('./db');
const insights = require('./insights');
const { allCategories, invalidateOverrides } = require('./categorize');
const { collectAll } = require('./collect');

/** One collection at a time — a second click while it runs is a no-op. */
const collectState = { running: false };

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

/** Wraps a handler so a thrown error becomes a clean 500 instead of a hang. */
const handle = (fn) => (req, res) => {
  try {
    res.json(fn(req));
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
};

app.get('/api/summary', handle(() => insights.summary()));
app.get('/api/monthly', handle((req) => insights.monthlySeries(Number(req.query.months) || 12)));
app.get('/api/categories', handle((req) => insights.categoryBreakdown(req.query.month)));
app.get('/api/merchants', handle((req) => insights.topMerchants(req.query.month)));
app.get('/api/recurring', handle(() => insights.recurring()));
app.get('/api/anomalies', handle(() => insights.anomalies()));
app.get('/api/forecast', handle(() => insights.forecast()));
app.get('/api/months', handle(() => insights.availableMonths()));
app.get('/api/transactions', handle((req) => insights.transactions({
  month: req.query.month,
  category: req.query.category,
  source: req.query.source,
  search: req.query.search,
  limit: Math.min(Number(req.query.limit) || 200, 1000),
  offset: Number(req.query.offset) || 0,
})));

app.get(
  '/api/status',
  handle(() => ({
    lastRun: store.lastRun(),
    sources: config.activeSources().map((source) => ({ key: source.key, label: source.label, kind: source.kind })),
    categories: allCategories(),
    collecting: collectState.running,
    cron: config.cron,
  }))
);

/** Manual re-categorisation: one transaction, or a rule for every match of a phrase. */
app.post('/api/category', (req, res) => {
  const { id, pattern, category } = req.body || {};
  if (!category) return res.status(400).json({ error: 'חסרה קטגוריה' });

  if (pattern) {
    store.db
      .prepare(
        `INSERT INTO category_overrides (pattern, category, created_at) VALUES (?, ?, ?)
         ON CONFLICT(pattern) DO UPDATE SET category = excluded.category`
      )
      .run(pattern, category, new Date().toISOString());
    invalidateOverrides();
    const info = store.db
      .prepare(
        `UPDATE transactions SET category = ?, category_source = 'manual'
          WHERE lower(description) LIKE lower(?)`
      )
      .run(category, `%${pattern}%`);
    return res.json({ updated: info.changes });
  }

  if (!id) return res.status(400).json({ error: 'חסר מזהה תנועה' });
  const info = store.db
    .prepare(`UPDATE transactions SET category = ?, category_source = 'manual' WHERE id = ?`)
    .run(category, id);
  return res.json({ updated: info.changes });
});

app.post('/api/collect', async (req, res) => {
  if (collectState.running) return res.status(409).json({ error: 'איסוף כבר רץ' });
  if (!config.activeSources().length) {
    return res.status(400).json({ error: 'לא הוגדרו פרטי התחברות בקובץ .env' });
  }
  collectState.running = true;
  res.json({ started: true });
  try {
    await collectAll();
  } catch (error) {
    console.error('שגיאת איסוף:', error.message);
  } finally {
    collectState.running = false;
  }
});

if (config.cron) {
  cron.schedule(config.cron, async () => {
    if (collectState.running) return;
    collectState.running = true;
    try {
      await collectAll();
    } catch (error) {
      console.error('שגיאת איסוף מתוזמן:', error.message);
    } finally {
      collectState.running = false;
    }
  });
}

app.listen(config.port, config.host, () => {
  console.log(`\n  ניתוח פיננסי אישי → http://${config.host}:${config.port}`);
  console.log(`  מסד נתונים: ${config.dbPath}`);
  console.log(`  מקורות פעילים: ${config.activeSources().map((s) => s.label).join(', ') || 'אין (ראה .env)'}`);
  console.log(`  איסוף אוטומטי: ${config.cron || 'כבוי'}\n`);
});
