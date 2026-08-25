'use strict';

/* ------------------------------------------------------------------ helpers */

const $ = (id) => document.getElementById(id);
const NS = 'http://www.w3.org/2000/svg';

const shekel = new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS', maximumFractionDigits: 0 });
const shekelExact = new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS', minimumFractionDigits: 2 });
const HE_MONTHS = ['ינו', 'פבר', 'מרץ', 'אפר', 'מאי', 'יונ', 'יול', 'אוג', 'ספט', 'אוק', 'נוב', 'דצמ'];

const money = (value) => shekel.format(Math.round(value));
const moneyExact = (value) => shekelExact.format(value);
/** Currency for SVG: charts render LTR, so build the string without bidi marks. */
const plain = new Intl.NumberFormat('he-IL', { maximumFractionDigits: 0 });
const moneyFlat = (value) => `${plain.format(Math.round(value))} \u20aa`;

function monthLabel(key) {
  const [year, month] = key.split('-');
  return `${HE_MONTHS[Number(month) - 1]} ${year.slice(2)}`;
}

function dateLabel(iso) {
  const [year, month, day] = iso.split('-');
  return `${day}.${month}.${year.slice(2)}`;
}

function el(tag, attrs = {}, text) {
  const node = document.createElementNS(NS, tag);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
  if (text != null) node.textContent = text;
  return node;
}

/** A rect with only the data end rounded, so the baseline end stays flush. */
function barPath(x, y, width, height, radius, roundedSide) {
  const r = Math.max(0, Math.min(radius, width, height / 2));
  if (r === 0 || width <= 0) return `M${x},${y}h${Math.max(width, 0)}v${height}h${-Math.max(width, 0)}z`;
  return roundedSide === 'left'
    ? `M${x + r},${y}h${width - r}v${height}h${-(width - r)}a${r},${r} 0 0 1 ${-r},${-r}v${-(height - 2 * r)}a${r},${r} 0 0 1 ${r},${-r}z`
    : `M${x},${y}h${width - r}a${r},${r} 0 0 1 ${r},${r}v${height - 2 * r}a${r},${r} 0 0 1 ${-r},${r}h${-(width - r)}z`;
}

function niceMax(value) {
  if (value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  return Math.ceil(value / (magnitude / 2)) * (magnitude / 2);
}

/* ------------------------------------------------------------------ tooltip */

const tooltip = $('tooltip');

function showTip(event, html) {
  tooltip.innerHTML = html;
  tooltip.classList.add('on');
  const box = tooltip.getBoundingClientRect();
  let left = event.clientX - box.width - 14;
  if (left < 8) left = event.clientX + 14;
  let top = event.clientY - box.height - 12;
  if (top < 8) top = event.clientY + 16;
  tooltip.style.left = `${left}px`;
  tooltip.style.top = `${top}px`;
}

const hideTip = () => tooltip.classList.remove('on');
document.addEventListener('scroll', hideTip, true);

/* --------------------------------------------------------------------- data */

const state = { month: null, months: [], categories: [], sources: [], filters: {} };

async function api(path, options) {
  const response = await fetch(path, options);
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error || response.statusText);
  return response.json();
}

/* ---------------------------------------------------------------- stat tiles */

function renderTiles(summary, forecast) {
  const change = summary.spendChangePct;
  const direction = change > 0 ? 'up' : 'down';
  const arrow = change > 0 ? '▲' : '▼';

  const tiles = [
    {
      label: `הוצאות ${monthLabel(summary.thisMonth)}`,
      value: money(summary.spend),
      sub: summary.lastMonthSpend
        ? `<span class="delta ${direction}">${arrow} ${Math.abs(change)}%</span> מול ${monthLabel(summary.lastMonth)} (${money(summary.lastMonthSpend)})`
        : 'אין עדיין חודש קודם להשוואה',
    },
    {
      label: 'יתרה בחשבון',
      value: summary.balances.length ? money(summary.totalBalance) : '—',
      sub: summary.balances.length
        ? summary.balances.map((b) => `${b.source_label} ${b.account}`).join(' · ')
        : 'הבנק לא החזיר יתרה',
    },
    {
      label: 'בסיס קבוע חודשי',
      value: money(summary.fixedMonthly),
      sub: `${summary.recurringCount} חיובים חוזרים — יוצא לפני שנגעת בכלום`,
    },
    {
      label: 'תחזית לסוף החודש',
      value: money(forecast.projectedSpend),
      sub:
        summary.avgSpend > 0
          ? `הממוצע ב־6 החודשים האחרונים: ${money(summary.avgSpend)}`
          : 'עוד אין מספיק היסטוריה לממוצע',
    },
  ];

  $('tiles').innerHTML = tiles
    .map(
      (tile) => `<div class="tile">
        <div class="label">${tile.label}</div>
        <div class="value">${tile.value}</div>
        <div class="sub">${tile.sub}</div>
      </div>`
    )
    .join('');
}

/* ------------------------------------------------------------- forecast bar */

function renderForecast(forecast, summary) {
  $('forecast-day').textContent = `יום ${forecast.dayOfMonth} מתוך ${forecast.daysInMonth}`;

  const parts = [
    { label: 'כבר יצא', value: forecast.spentSoFar, fill: 'var(--series-1)' },
    { label: 'קבוע שעוד יחויב', value: forecast.committed, fill: 'var(--series-1)', opacity: 0.55 },
    { label: 'צפי הוצאה שוטפת', value: forecast.variableLeft, fill: 'var(--series-1)', opacity: 0.28 },
  ].filter((part) => part.value > 0);

  const total = parts.reduce((sum, part) => sum + part.value, 0);
  const reference = summary.avgSpend;
  const scaleMax = niceMax(Math.max(total, reference) * 1.08) || 1;

  $('forecast-note').innerHTML =
    total > 0
      ? `נשארו ${forecast.daysLeft} ימים. לפי קצב של <b>${money(forecast.dailyVariable)}</b> ליום בהוצאות שוטפות, החודש יסתיים סביב <b>${money(forecast.projectedSpend)}</b>` +
        (reference ? ` — ${total > reference ? 'מעל' : 'מתחת ל'}ממוצע של ${money(reference)}.` : '.') +
        (forecast.scheduledCardDebits
          ? ` בנוסף, חברות האשראי כבר קבעו חיוב עתידי של ${money(forecast.scheduledCardDebits)}.`
          : '')
      : 'עוד אין מספיק נתונים לחודש הזה.';

  const width = Math.min(1100, $('forecast-chart').clientWidth || 900);
  const height = 92;
  const barY = 22;
  const barH = 30;
  const gap = 2;

  const svg = el('svg', { viewBox: `0 0 ${width} ${height}`, width, height, role: 'img' });
  svg.appendChild(el('title', {}, `תחזית הוצאות: ${money(total)} מתוך סקאלה של ${money(scaleMax)}`));

  // Track. Bars grow from the right, matching the page direction.
  svg.appendChild(el('rect', { x: 0, y: barY, width, height: barH, rx: 4, fill: 'var(--grid)', opacity: 0.5 }));

  let cursor = width;
  for (const part of parts) {
    const segment = (part.value / scaleMax) * width;
    const drawn = Math.max(0, segment - gap);
    const isLast = part === parts[parts.length - 1];
    const path = el('path', {
      d: barPath(cursor - drawn, barY, drawn, barH, 4, isLast ? 'left' : 'none'),
      fill: part.fill,
      opacity: part.opacity ?? 1,
    });
    path.addEventListener('mousemove', (event) =>
      showTip(event, `<div class="t-title">${part.label}</div><div class="t-row"><b>${moneyExact(part.value)}</b></div>`)
    );
    path.addEventListener('mouseleave', hideTip);
    svg.appendChild(path);
    cursor -= segment;
  }

  if (reference > 0) {
    const x = width - (reference / scaleMax) * width;
    svg.appendChild(el('line', { x1: x, x2: x, y1: barY - 7, y2: barY + barH + 7, stroke: 'var(--text-secondary)', 'stroke-width': 2, 'stroke-dasharray': '3 3' }));
    // Middle-anchored, so an rtl override reorders the words without moving the label.
    svg.appendChild(el('text', { x, y: barY - 11, 'text-anchor': 'middle', class: 'tick', style: 'direction: rtl' }, `ממוצע ${moneyFlat(reference)}`));
  }

  const legend = document.createElement('div');
  legend.className = 'legend';
  legend.innerHTML = parts
    .map(
      (part) =>
        `<span><i class="swatch" style="background: ${part.fill}; opacity: ${part.opacity ?? 1}"></i>${part.label} ${money(part.value)}</span>`
    )
    .join('');

  $('forecast-chart').replaceChildren(svg, legend);
}

/* ------------------------------------------------------------ monthly chart */

let monthlyData = [];

function renderMonthly(series) {
  monthlyData = series;
  const host = $('monthly-chart');
  if (!series.length) {
    host.innerHTML = '<p class="empty">אין עדיין נתונים.</p>';
    return;
  }

  const width = Math.max(560, Math.min(1100, host.clientWidth || 900));
  const height = 300;
  const pad = { top: 18, right: 62, bottom: 30, left: 62 };
  const plotW = width - pad.left - pad.right;
  const plotH = height - pad.top - pad.bottom;

  const max = niceMax(Math.max(...series.flatMap((row) => [row.income, row.spend])) * 1.1) || 1;
  const x = (index) => pad.left + (series.length === 1 ? plotW / 2 : (index / (series.length - 1)) * plotW);
  const y = (value) => pad.top + plotH - (value / max) * plotH;

  const svg = el('svg', { viewBox: `0 0 ${width} ${height}`, width, height, role: 'img' });
  svg.appendChild(el('title', {}, 'הכנסות מול הוצאות לפי חודש'));

  for (let step = 0; step <= 4; step += 1) {
    const value = (max / 4) * step;
    const gy = y(value);
    svg.appendChild(el('line', { x1: pad.left, x2: pad.left + plotW, y1: gy, y2: gy, class: 'grid-line' }));
    svg.appendChild(el('text', { x: pad.left + plotW + 8, y: gy + 4, class: 'tick num' }, moneyFlat(value)));
  }

  series.forEach((row, index) => {
    if (series.length > 8 && index % 2 === 1 && index !== series.length - 1) return;
    svg.appendChild(
      el('text', { x: x(index), y: height - 10, 'text-anchor': 'middle', class: 'tick' }, monthLabel(row.month))
    );
  });

  const lines = [
    { key: 'spend', color: 'var(--series-1)', label: 'הוצאות' },
    { key: 'income', color: 'var(--series-2)', label: 'הכנסות' },
  ];

  for (const line of lines) {
    const d = series.map((row, index) => `${index ? 'L' : 'M'}${x(index)},${y(row[line.key])}`).join('');
    svg.appendChild(el('path', { d, fill: 'none', stroke: line.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));

    series.forEach((row, index) => {
      // A 2px surface ring keeps the markers legible where the lines cross.
      svg.appendChild(el('circle', { cx: x(index), cy: y(row[line.key]), r: 4.5, fill: line.color, stroke: 'var(--surface-1)', 'stroke-width': 2 }));
    });

    // Direct label at the last point — the relief the contrast check asks for.
    const last = series[series.length - 1];
    svg.appendChild(
      el('text', { x: x(series.length - 1) - 10, y: y(last[line.key]) - 11, 'text-anchor': 'middle', class: 'series-label', fill: line.color }, line.label)
    );
  }

  // One transparent hit band per month drives the crosshair.
  const crosshair = el('line', { y1: pad.top, y2: pad.top + plotH, class: 'axis-line', 'stroke-dasharray': '3 3', opacity: 0 });
  svg.appendChild(crosshair);

  const bandW = plotW / Math.max(1, series.length - 1 || 1);
  series.forEach((row, index) => {
    const band = el('rect', {
      x: x(index) - bandW / 2,
      y: pad.top,
      width: bandW,
      height: plotH,
      fill: 'transparent',
    });
    band.addEventListener('mousemove', (event) => {
      crosshair.setAttribute('x1', x(index));
      crosshair.setAttribute('x2', x(index));
      crosshair.setAttribute('opacity', 1);
      showTip(
        event,
        `<div class="t-title">${monthLabel(row.month)}</div>
         <div class="t-row"><i class="swatch" style="background: var(--series-1)"></i>הוצאות<b>${money(row.spend)}</b></div>
         <div class="t-row"><i class="swatch" style="background: var(--series-2)"></i>הכנסות<b>${money(row.income)}</b></div>
         <div class="t-row">נטו<b>${money(row.net)}</b></div>`
      );
    });
    band.addEventListener('mouseleave', () => {
      crosshair.setAttribute('opacity', 0);
      hideTip();
    });
    svg.appendChild(band);
  });

  host.replaceChildren(svg);

  $('monthly-table').innerHTML = `<table>
    <thead><tr><th>חודש</th><th class="num">הכנסות</th><th class="num">הוצאות</th><th class="num">נטו</th></tr></thead>
    <tbody>${[...series].reverse().map((row) => `<tr>
      <td>${monthLabel(row.month)}</td>
      <td class="num">${money(row.income)}</td>
      <td class="num">${money(row.spend)}</td>
      <td class="num ${row.net >= 0 ? 'pos' : ''}">${money(row.net)}</td>
    </tr>`).join('')}</tbody>
  </table>`;
}

/* ----------------------------------------------------------- category chart */

function renderCategories(rows) {
  const host = $('category-chart');
  if (!rows.length) {
    host.innerHTML = '<p class="empty">אין הוצאות בחודש הזה.</p>';
    return;
  }

  const width = Math.max(520, Math.min(1100, host.clientWidth || 900));
  const rowH = 32;
  const labelW = 150;
  const valueW = 132;
  const trackW = width - labelW - valueW;
  const height = rows.length * rowH + 14;
  const max = Math.max(...rows.map((row) => row.total));

  const svg = el('svg', { viewBox: `0 0 ${width} ${height}`, width, height, role: 'img' });
  svg.appendChild(el('title', {}, 'פילוח הוצאות לפי קטגוריה'));

  rows.forEach((row, index) => {
    const y = index * rowH + 6;
    const barH = 18;
    const barW = Math.max(2, (row.total / max) * trackW);
    // Right-anchored: the label sits on the right, the bar grows leftward.
    const right = width - labelW;

    svg.appendChild(el('text', { x: width, y: y + 13, 'text-anchor': 'end', class: 'bar-name' }, row.category));

    const bar = el('path', {
      d: barPath(right - barW, y, barW, barH, 4, 'left'),
      fill: 'var(--series-1)',
    });
    bar.addEventListener('mousemove', (event) =>
      showTip(
        event,
        `<div class="t-title">${row.category}</div>
         <div class="t-row">סה"כ<b>${moneyExact(row.total)}</b></div>
         <div class="t-row">תנועות<b>${row.count}</b></div>
         <div class="t-row">מסך ההוצאות<b>${row.share}%</b></div>`
      )
    );
    bar.addEventListener('mouseleave', hideTip);
    bar.style.cursor = 'pointer';
    bar.addEventListener('click', () => {
      $('filter-category').value = row.category;
      loadTransactions();
      $('txn-table').scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
    svg.appendChild(bar);

    svg.appendChild(
      el('text', { x: Math.max(valueW - 8, right - barW - 10), y: y + 13, 'text-anchor': 'end', class: 'bar-value' }, `${moneyFlat(row.total)} · ${row.share}%`)
    );
  });

  host.replaceChildren(svg);
}

/* ---------------------------------------------------------------- anomalies */

const ICONS = { critical: '⛔', serious: '⚠️', warning: '❗' };

function renderAlerts(items) {
  const host = $('alerts');
  if (!items.length) {
    host.innerHTML = '<p class="empty">לא נמצא שום דבר חריג. 👌</p>';
    return;
  }
  host.innerHTML = items
    .map(
      (item) => `<div class="alert ${item.severity}">
        <span class="icon" aria-hidden="true">${ICONS[item.severity] || '❗'}</span>
        <div class="body">
          <div class="a-title">${item.kind} · ${escapeHtml(item.title)}</div>
          <div class="a-detail">${escapeHtml(item.detail)}</div>
          <div class="a-meta">${dateLabel(item.date)} · ${escapeHtml(item.source)}</div>
        </div>
        <div class="a-amount">${moneyExact(item.amount)}</div>
      </div>`
    )
    .join('');
}

/* ---------------------------------------------------------------- recurring */

function renderRecurring(items) {
  const active = items.filter((item) => !item.dormant);
  $('fixed-total').textContent = `${money(active.reduce((sum, item) => sum + item.monthlyEquivalent, 0))} לחודש`;

  const host = $('recurring-table');
  if (!items.length) {
    host.innerHTML = '<p class="empty">עוד לא זוהו חיובים חוזרים — צריך לפחות 3 חודשי היסטוריה.</p>';
    return;
  }

  host.innerHTML = `<table>
    <thead><tr>
      <th>חיוב</th><th>קטגוריה</th><th>תדירות</th>
      <th class="num">סכום טיפוסי</th><th class="num">חיוב אחרון</th><th>מתי</th><th></th>
    </tr></thead>
    <tbody>${items
      .map(
        (item) => `<tr${item.dormant ? ' style="opacity:.55"' : ''}>
          <td>${escapeHtml(item.label)}</td>
          <td class="muted-cell">${escapeHtml(item.category)}</td>
          <td><span class="pill">${item.cadence} · ${item.kind}</span></td>
          <td class="num">${moneyExact(item.typicalAmount)}</td>
          <td class="num">${moneyExact(item.lastAmount)}</td>
          <td class="muted-cell">${dateLabel(item.lastDate)}</td>
          <td>${
            item.dormant
              ? '<span class="pill">רדום</span>'
              : item.priceJumpPct >= 8
                ? `<span class="delta up">▲ ${item.priceJumpPct}%</span>`
                : ''
          }</td>
        </tr>`
      )
      .join('')}</tbody>
  </table>`;
}

/* ------------------------------------------------------------- transactions */

function renderTransactions({ rows, total }) {
  $('txn-count').textContent = `${total.toLocaleString('he-IL')} תנועות`;
  const host = $('txn-table');
  if (!rows.length) {
    host.innerHTML = '<p class="empty">אין תנועות שמתאימות לסינון.</p>';
    return;
  }

  const options = state.categories.map((category) => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`).join('');

  host.innerHTML = `<table>
    <thead><tr><th>תאריך</th><th>תיאור</th><th>קטגוריה</th><th>מקור</th><th class="num">סכום</th></tr></thead>
    <tbody>${rows
      .map(
        (row) => `<tr>
          <td class="muted-cell">${dateLabel(row.date)}</td>
          <td>${escapeHtml(row.description)}${
            row.installment_total > 1 ? ` <span class="pill">תשלום ${row.installment_num}/${row.installment_total}</span>` : ''
          }${row.status === 'pending' ? ' <span class="pill">טרם נסלק</span>' : ''}</td>
          <td><select class="cat-pick" data-id="${row.id}" aria-label="שינוי קטגוריה">${options}</select></td>
          <td class="muted-cell">${escapeHtml(row.source_label)}</td>
          <td class="num ${row.amount > 0 ? 'pos' : 'neg'}">${moneyExact(row.amount)}</td>
        </tr>`
      )
      .join('')}</tbody>
  </table>`;

  host.querySelectorAll('.cat-pick').forEach((select, index) => {
    select.value = rows[index].category;
    select.addEventListener('change', async () => {
      await api('/api/category', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: select.dataset.id, category: select.value }),
      });
      refreshAnalytics();
    });
  });
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]
  );
}

/* ---------------------------------------------------------------- app wiring */

async function loadTransactions() {
  const params = new URLSearchParams();
  if (state.month) params.set('month', state.month);
  const category = $('filter-category').value;
  const source = $('filter-source').value;
  const search = $('filter-search').value.trim();
  if (category) params.set('category', category);
  if (source) params.set('source', source);
  if (search) params.set('search', search);
  renderTransactions(await api(`/api/transactions?${params}`));
}

async function refreshAnalytics() {
  const [summary, forecast, series, categories, recurring, anomalies] = await Promise.all([
    api('/api/summary'),
    api('/api/forecast'),
    api('/api/monthly?months=12'),
    api(`/api/categories?month=${state.month || ''}`),
    api('/api/recurring'),
    api('/api/anomalies'),
  ]);

  renderTiles(summary, forecast);
  renderForecast(forecast, summary);
  renderMonthly(series);
  renderCategories(categories);
  renderRecurring(recurring);
  renderAlerts(anomalies);
  $('cat-month').textContent = state.month ? monthLabel(state.month) : '';
  await loadTransactions();
}

function renderStatus(status) {
  const dot = $('status').querySelector('.dot');
  const text = $('status-text');

  if (status.collecting) {
    dot.className = 'dot busy';
    text.textContent = 'אוסף נתונים מהבנק…';
    return true;
  }
  if (!status.sources.length) {
    dot.className = 'dot bad';
    text.textContent = 'לא הוגדרו מקורות';
    $('setup').hidden = false;
    return false;
  }

  $('setup').hidden = true;
  const run = status.lastRun;
  if (!run || !run.finished_at) {
    dot.className = 'dot stale';
    text.textContent = 'עוד לא בוצע איסוף';
    return false;
  }

  const ageHours = (Date.now() - new Date(run.finished_at)) / 3600000;
  dot.className = run.status === 'failed' ? 'dot bad' : ageHours > 36 ? 'dot stale' : 'dot';
  const when = new Date(run.finished_at).toLocaleString('he-IL', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
  text.textContent =
    run.status === 'failed'
      ? `האיסוף האחרון נכשל (${when})`
      : `עודכן ${when}${run.detail ? ' · עם שגיאות חלקיות' : ''}`;
  text.title = run.detail || '';
  return false;
}

async function pollStatus() {
  const status = await api('/api/status');
  const busy = renderStatus(status);
  $('refresh').disabled = busy || !status.sources.length;
  return busy;
}

async function init() {
  const status = await api('/api/status');
  state.categories = status.categories;
  state.sources = status.sources;
  renderStatus(status);
  $('refresh').disabled = !status.sources.length;

  state.months = await api('/api/months');
  state.month = state.months[0] || new Date().toISOString().slice(0, 7);

  $('month-select').innerHTML = state.months.map((month) => `<option value="${month}">${monthLabel(month)}</option>`).join('');
  $('month-select').value = state.month;
  $('month-select').addEventListener('change', (event) => {
    state.month = event.target.value;
    refreshAnalytics();
  });

  $('filter-category').innerHTML =
    '<option value="">כל הקטגוריות</option>' + state.categories.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  $('filter-source').innerHTML =
    '<option value="">כל המקורות</option>' + state.sources.map((s) => `<option value="${s.key}">${escapeHtml(s.label)}</option>`).join('');

  $('filter-category').addEventListener('change', loadTransactions);
  $('filter-source').addEventListener('change', loadTransactions);
  let searchTimer;
  $('filter-search').addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(loadTransactions, 250);
  });

  $('toggle-monthly-table').addEventListener('click', () => {
    const table = $('monthly-table');
    const chart = $('monthly-chart');
    const showTable = table.hidden;
    table.hidden = !showTable;
    chart.hidden = showTable;
    $('toggle-monthly-table').textContent = showTable ? 'הצג כגרף' : 'הצג כטבלה';
  });

  $('refresh').addEventListener('click', async () => {
    $('refresh').disabled = true;
    try {
      await api('/api/collect', { method: 'POST' });
    } catch (error) {
      alert(`לא ניתן להתחיל איסוף: ${error.message}`);
      $('refresh').disabled = false;
      return;
    }
    // Poll until the collection finishes, then reload everything it brought in.
    const timer = setInterval(async () => {
      if (await pollStatus()) return;
      clearInterval(timer);
      state.months = await api('/api/months');
      $('month-select').innerHTML = state.months.map((m) => `<option value="${m}">${monthLabel(m)}</option>`).join('');
      $('month-select').value = state.month = state.months[0] || state.month;
      refreshAnalytics();
    }, 2500);
  });

  const themeButton = $('theme');
  const stored = (() => {
    try { return localStorage.getItem('theme'); } catch { return null; }
  })();
  if (stored) document.documentElement.dataset.theme = stored;
  themeButton.addEventListener('click', () => {
    const isDark = document.documentElement.dataset.theme === 'dark'
      || (!document.documentElement.dataset.theme && matchMedia('(prefers-color-scheme: dark)').matches);
    const next = isDark ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem('theme', next); } catch { /* private mode — theme just won't persist */ }
    refreshAnalytics();
  });

  await refreshAnalytics();

  let resizeTimer;
  addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(refreshAnalytics, 200);
  });
}

init().catch((error) => {
  document.querySelector('main').insertAdjacentHTML(
    'afterbegin',
    `<section class="setup"><h2>שגיאה בטעינה</h2><p class="sec-note">${escapeHtml(error.message)}</p></section>`
  );
});
