// Drop-in endpoint for the nutrition app on Railway.
//
//   const garmin = require("./garmin-endpoint");
//   app.use(express.json());
//   app.use("/garmin", garmin({ apiKey: process.env.GARMIN_API_KEY }));
//
// The watch then points at  https://<your-app>.up.railway.app/garmin/sets
//
// Storage here is a plain JSON file so it works with zero dependencies. Swap
// `store` for your real database when you wire this into the nutrition app.
const express = require("express");
const fs = require("fs");
const path = require("path");

const DEFAULT_FILE = path.join(process.env.DATA_DIR || __dirname, "garmin-days.json");

function fileStore(file = DEFAULT_FILE) {
  let days = {};
  try {
    days = JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (err) {
    days = {};
  }

  const flush = () => {
    try {
      fs.writeFileSync(file, JSON.stringify(days, null, 2));
    } catch (err) {
      console.error("garmin: could not persist", err.message);
    }
  };

  return {
    put(record) {
      days[record.date] = record;
      flush();
      return record;
    },
    get(date) {
      return days[date] || null;
    },
    all() {
      return Object.values(days).sort((a, b) => a.date.localeCompare(b.date));
    },
  };
}

// The watch posts the whole day every time, so this is idempotent by design:
// the newest payload for a date simply replaces the previous one.
function normalize(body) {
  const num = (value, fallback = 0) =>
    Number.isFinite(Number(value)) ? Number(value) : fallback;

  const date = String(body.date || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return null;
  }

  return {
    date,
    sets: num(body.sets),
    kcalPerSet: num(body.kcalPerSet, 50),
    consumedKcal: num(body.consumedKcal),
    activeKcal: num(body.activeKcal),
    totalKcal: num(body.totalKcal),
    baseBudgetKcal: num(body.baseBudgetKcal),
    allowanceKcal: num(body.allowanceKcal),
    remainingKcal: num(body.remainingKcal),
    steps: num(body.steps),
    source: body.source === "garmin" ? "garmin" : "unknown",
    watchTs: num(body.ts),
    receivedAt: new Date().toISOString(),
  };
}

module.exports = function garminRouter(options = {}) {
  const store = options.store || fileStore(options.file);
  const apiKey = options.apiKey;
  const router = express.Router();

  router.use((req, res, next) => {
    if (!apiKey) {
      return next(); // no key configured - open endpoint
    }
    if (req.get("X-Api-Key") === apiKey) {
      return next();
    }
    return res.status(401).json({ error: "bad api key" });
  });

  router.post("/sets", (req, res) => {
    const record = normalize(req.body || {});
    if (!record) {
      return res.status(400).json({ error: "missing or malformed date" });
    }
    store.put(record);
    console.log(
      `garmin ${record.date}: ${record.sets} sets = ${record.consumedKcal} kcal, ` +
        `${record.remainingKcal} left`
    );
    res.json({ ok: true, day: record });
  });

  router.get("/sets/:date", (req, res) => {
    const day = store.get(req.params.date);
    if (!day) {
      return res.status(404).json({ error: "no data for that date" });
    }
    res.json(day);
  });

  router.get("/sets", (req, res) => {
    res.json({ days: store.all() });
  });

  return router;
};

module.exports.fileStore = fileStore;
module.exports.normalize = normalize;
