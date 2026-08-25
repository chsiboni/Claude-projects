'use strict';

/**
 * Re-applies the categorisation rules to everything already stored — run this
 * after editing src/categorize.js. Categories you set by hand are left alone.
 */

const { db } = require('./db');
const { categorize } = require('./categorize');

const rows = db
  .prepare(`SELECT id, description, amount, category FROM transactions WHERE category_source != 'manual'`)
  .all();

const update = db.prepare('UPDATE transactions SET category = ? WHERE id = ?');
let changed = 0;

db.transaction(() => {
  for (const row of rows) {
    const category = categorize(row.description, row.amount);
    if (category !== row.category) {
      update.run(category, row.id);
      changed += 1;
    }
  }
})();

console.log(`נבדקו ${rows.length} תנועות, ${changed} סווגו מחדש.`);
