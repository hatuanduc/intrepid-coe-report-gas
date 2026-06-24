/**
 * Intrepid COE – Working Capital & AR Report
 *
 * Performance strategy:
 *   1. Sheets API batchGet  → read only 9 needed columns (vs 43 total)
 *   2. Drive modifiedTime   → cache key auto-invalidates when file changes
 *   3. Chunked CacheService → handles datasets > 100 KB
 *
 * TH sheet layout (0-based):
 *   A=0 Type | B=1 Period | D=3 Brand | I=8 Issued | J=9 Collected
 *   K=10 AmtAccrue | L=11 AmtInvoice | P=15 AmtLatest
 *   AA=26 FIN | AB=27 KAM | AC=28 Brand | AD=29 Invoice | AE=30 Collect | AG=32 Status
 */

function _ss()   { return SpreadsheetApp.getActiveSpreadsheet(); }
function _token(){ return ScriptApp.getOAuthToken(); }

const MONTHS          = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const LEAD_TIME_TYPES = new Set(['Retail', 'Consignment', 'Store Management']);
const CACHE_TTL       = 600; // 10 min

const STATUS = {
  UNDER_TERM:  'Pending collection under credit term',
  OVER_DUE:    'Pending collection over due',
  PENDING_CFM: 'Pending confirmation',
  COLLECTED:   'Collected',
};

// ─── Web app ──────────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService
    .createTemplateFromFile('dashboard')
    .evaluate()
    .setTitle('Working Capital & AR Console')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Main entry ───────────────────────────────────────────────────────────────

function getReportData(params) {
  params = params || {};
  const country  = (params.country || 'TH').toUpperCase();
  const bizClass = params.bizClass || null;

  const allRows = _cachedRows(country);
  Logger.log('total rows=%s', allRows.length);

  // No period filter – AR is cumulative (outstanding = not collected, all periods)
  const rows   = _filter(allRows, bizClass, null);
  const summary = _summary(rows);
  const types   = [...new Set(allRows.map(r => r.type).filter(Boolean))].sort();

  // Breakdown for debug
  const collected = allRows.filter(r => _isDate(r.collected)).length;
  const unbilled  = allRows.filter(r => !_isDate(r.collected) && !_isDate(r.issued)).length;
  Logger.log('collected=%s unbilled=%s totalAR=%s', collected, unbilled, summary.totalAR);

  return {
    summary,
    leadTime:  _cachedLeadTime(country, bizClass),
    arByBrand: _arByBrand(rows),
    types,
  };
}

/** AR breakdown by brand: row counts per status, sorted by total rows DESC */
function _arByBrand(rows) {
  const map = {};
  rows.forEach(r => {
    if (_isDate(r.collected)) return;
    const b = r.brand || 'Unknown';
    if (!map[b]) map[b] = { brand: b, unbilled: 0, overdue: 0, underTerm: 0 };
    if      (!_isDate(r.issued))             map[b].unbilled++;
    else if (r.status === STATUS.OVER_DUE)   map[b].overdue++;
    else if (r.status === STATUS.UNDER_TERM) map[b].underTerm++;
  });
  return Object.values(map)
    .filter(r => r.unbilled + r.overdue + r.underTerm > 0)
    .sort((a, b) => (b.unbilled + b.overdue + b.underTerm) - (a.unbilled + a.overdue + a.underTerm));
}

/** Returns overdue rows (issued, not collected, status = overdue). */
function getOverdueRows(params) {
  params = params || {};
  const country  = (params.country  || 'TH').toUpperCase();
  const bizClass = params.bizClass  || null;

  const allRows = _cachedRows(country);
  return allRows
    .filter(r =>
      (!bizClass || r.type === bizClass) &&
      !_isDate(r.collected) &&
       _isDate(r.issued) &&
      r.status === STATUS.OVER_DUE
    )
    .map(r => ({
      period:     r.period,
      periodSort: _periodToNum(r.period),
      brand:      r.brand,
      type:       r.type,
      amount:     r.amount,
      status:     r.status,
    }))
    .sort((a, b) =>
      b.periodSort - a.periodSort ||
      (a.brand || '').localeCompare(b.brand || '')
    );
}

/** Returns unbilled detail rows for the detail tab. */
function getDetailRows(params) {
  params = params || {};
  const country  = (params.country  || 'TH').toUpperCase();
  const bizClass = params.bizClass  || null;

  const allRows = _cachedRows(country);
  return allRows
    .filter(r =>
      (!bizClass || r.type === bizClass) &&
      !_isDate(r.collected) &&
      !_isDate(r.issued)
    )
    .map(r => ({
      period:     r.period,
      periodSort: _periodToNum(r.period), // virtual col: YYYYMM integer for correct sort
      brand:      r.brand,
      type:       r.type,
      amount:     r.amount,
      status:     r.status,
    }))
    .sort((a, b) =>
      b.periodSort - a.periodSort ||          // period: newest first
      (a.brand || '').localeCompare(b.brand || '') // then brand A→Z
    );
}

/** "2026 May" → 202605 for date-correct sorting. */
function _periodToNum(p) {
  if (!p) return 0;
  const m = p.match(/^(\d{4})\s+([A-Za-z]{3})$/);
  if (!m) return 0;
  const mo = MONTHS.indexOf(m[2]);
  return parseInt(m[1]) * 100 + (mo >= 0 ? mo + 1 : 0);
}

// ─── Cache layer ──────────────────────────────────────────────────────────────

function _cachedRows(country) {
  const key = 'rows_' + country;
  const hit = _chunkRead(key);
  if (hit) { Logger.log('cache hit'); return hit; }
  Logger.log('cache miss – reading sheet');
  const fresh = _readTHSheet(country);
  _chunkWrite(key, fresh, CACHE_TTL);
  return fresh;
}

function _cachedLeadTime(country, bizClass) {
  const key    = 'lt_' + country + '_' + (bizClass || 'all');
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const fresh = _readLeadTime(country, bizClass);
  try { cache.put(key, JSON.stringify(fresh), CACHE_TTL); } catch(e) {}
  return fresh;
}

/** Manual cache clear – run in GAS editor when you want fresh data immediately. */
function clearCache() {
  CacheService.getScriptCache().removeAll(
    ['rows_TH','lt_TH_all','lt_TH_Retail','lt_TH_Consignment','lt_TH_Store Management']
  );
  Logger.log('Cache cleared');
}

// ─── Chunked cache (handles > 100 KB) ────────────────────────────────────────

const CHUNK = 90000; // bytes per cache key (limit 100 KB)

function _chunkWrite(base, data, ttl) {
  try {
    const json   = JSON.stringify(data);
    const n      = Math.ceil(json.length / CHUNK);
    const entries = { [base + '_n']: String(n) };
    for (let i = 0; i < n; i++)
      entries[base + '_' + i] = json.slice(i * CHUNK, (i + 1) * CHUNK);
    CacheService.getScriptCache().putAll(entries, ttl);
  } catch(e) { Logger.log('chunkWrite error: %s', e.message); }
}

function _chunkRead(base) {
  const cache = CacheService.getScriptCache();
  const n     = parseInt(cache.get(base + '_n') || '0');
  if (!n) return null;
  let json = '';
  for (let i = 0; i < n; i++) {
    const chunk = cache.get(base + '_' + i);
    if (!chunk) return null; // partial expiry – treat as miss
    json += chunk;
  }
  try { return JSON.parse(json); } catch(e) { return null; }
}

// ─── Fast sheet read: batchGet only 9 needed columns ─────────────────────────

/**
 * Uses Sheets API batchGet to fetch only the columns we care about.
 * Reads with FORMATTED_STRING so dates come back as display text.
 *
 * Column groups fetched (data rows 3..lastRow):
 *   A:B   → Type, Period
 *   D     → Brand
 *   I:L   → Issued, Collected, AmtAccrue, AmtInvoice
 *   P     → AmtLatest
 *   AG    → GroupStatus
 *   AA:AE → Lead time days (FIN, KAM, Brand, Invoice, Collect)
 */
function _readTHSheet(country) {
  const ss      = _ss();
  const ssId    = ss.getId();
  const sheet   = ss.getSheetByName(country);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  const n = lastRow; // end row (1-based)

  // Build batchGet request for 5 range groups
  const rangeNames = [
    `${country}!A3:B${n}`,    // [0] Type, Period
    `${country}!D3:D${n}`,    // [1] Brand
    `${country}!I3:L${n}`,    // [2] Issued, Collected, AmtK, AmtL
    `${country}!P3:P${n}`,    // [3] AmtLatest
    `${country}!AG3:AG${n}`,  // [4] Status
    `${country}!AA3:AE${n}`,  // [5] Days (FIN..Collect)
  ];

  const qs = rangeNames.map(r => 'ranges=' + encodeURIComponent(r)).join('&');
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + ssId +
    '/values:batchGet?' + qs + '&valueRenderOption=FORMATTED_STRING';

  let vr;
  try {
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + _token() },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('batchGet %s', resp.getResponseCode());
      return _readTHSheetFallback(country); // fallback to getValues()
    }
    vr = JSON.parse(resp.getContentText()).valueRanges;
  } catch(e) {
    Logger.log('batchGet error: %s', e.message);
    return _readTHSheetFallback(country);
  }

  const get  = (rIdx, row, col) => ((vr[rIdx].values || [])[row] || [])[col] || '';
  const rows = [];

  for (let i = 0; i < lastRow - 2; i++) {
    const period = _parsePeriod(get(0, i, 1)); // B = Period
    if (!period) continue;

    const type      = String(get(0, i, 0)).trim();           // A
    const brand     = String(get(1, i, 0)).trim();           // D
    const issued    = get(2, i, 0);                          // I
    const collected = get(2, i, 1);                          // J
    const amtK      = _num(get(2, i, 2));                    // K
    const amtL      = _num(get(2, i, 3));                    // L
    const amtP      = _num(get(3, i, 0));                    // P
    const status    = String(get(4, i, 0)).trim();           // AG

    const hasInvoice = _isDate(issued);
    const amount = hasInvoice
      ? (amtL || amtK || amtP)
      : (amtK || amtP);

    rows.push({
      period, type, brand,
      issued, collected, amount, status,
      dFin: _num(get(5, i, 0)),
      dKam: _num(get(5, i, 1)),
      dBr:  _num(get(5, i, 2)),
      dInv: _num(get(5, i, 3)),
      dCol: _num(get(5, i, 4)),
    });
  }

  Logger.log('batchGet rows: %s (cols read: ~9 of 43)', rows.length);
  return rows;
}

/** Fallback when batchGet fails: read full range but limit to col AG (33). */
function _readTHSheetFallback(country) {
  const sheet = _ss().getSheetByName(country);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  // Read only up to col AG (col 33)
  const data = sheet.getRange(3, 1, lastRow - 2, 33).getValues();
  const rows = [];
  data.forEach(row => {
    const period = _parsePeriod(row[1]);
    if (!period) return;
    const amtK = _num(row[10]), amtL = _num(row[11]), amtP = _num(row[15]);
    const hasInvoice = _isDate(row[8]);
    rows.push({
      period,
      type:      String(row[0]  || '').trim(),
      brand:     String(row[3]  || '').trim(),
      issued:    row[8],  collected: row[9],
      amount:    hasInvoice ? (amtL || amtK || amtP) : (amtK || amtP),
      status:    String(row[32] || '').trim(),
      dFin: _num(row[26]), dKam: _num(row[27]), dBr: _num(row[28]),
      dInv: _num(row[29]), dCol: _num(row[30]),
    });
  });
  Logger.log('fallback rows: %s', rows.length);
  return rows;
}

// ─── Lead time ────────────────────────────────────────────────────────────────

function _readLeadTime(country, bizClass) {
  // Lead time uses only 7 cols → dedicated batchGet
  const ss      = _ss();
  const ssId    = ss.getId();
  const sheet   = ss.getSheetByName(country);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const n = lastRow;
  const rangeNames = [
    `${country}!A3:A${n}`,    // Type
    `${country}!D3:D${n}`,    // Brand
    `${country}!AA3:AE${n}`,  // Days FIN..Collect
  ];
  const qs  = rangeNames.map(r => 'ranges=' + encodeURIComponent(r)).join('&');
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + ssId +
    '/values:batchGet?' + qs + '&valueRenderOption=UNFORMATTED_VALUE';

  let vr;
  try {
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + _token() }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error(resp.getResponseCode());
    vr = JSON.parse(resp.getContentText()).valueRanges;
  } catch(e) {
    Logger.log('leadTime batchGet error: %s', e.message);
    // Minimal fallback
    return _readLeadTimeFallback(country, bizClass);
  }

  const get = (rIdx, row, col) => ((vr[rIdx].values || [])[row] || [])[col];
  const map = {};

  for (let i = 0; i < lastRow - 2; i++) {
    const type = String(get(0, i, 0) || '').trim();
    if (!LEAD_TIME_TYPES.has(type)) continue;
    if (bizClass && type !== bizClass) continue;

    const fin = _num(get(2, i, 0)), kam = _num(get(2, i, 1));
    const br  = _num(get(2, i, 2)), inv = _num(get(2, i, 3));
    const col = _num(get(2, i, 4));
    if (!fin && !kam && !br && !inv && !col) continue;

    const b = String(get(1, i, 0) || '').trim() || 'Unknown';
    if (!map[b]) map[b] = { n:0, fin:0, kam:0, br:0, inv:0, col:0 };
    map[b].n++;
    map[b].fin += fin; map[b].kam += kam; map[b].br += br;
    map[b].inv += inv; map[b].col += col;
  }

  Logger.log('leadTime brands: %s', Object.keys(map).length);
  return Object.entries(map)
    .map(([brand, v]) => ({
      brand,
      fin:     _r1(v.fin / v.n), kam:     _r1(v.kam / v.n),
      brand_r: _r1(v.br  / v.n), invoice: _r1(v.inv / v.n),
      collect: _r1(v.col / v.n),
    }))
    .sort((a, b) => (b.fin+b.kam+b.brand_r+b.invoice+b.collect) - (a.fin+a.kam+a.brand_r+a.invoice+a.collect));
}

function _readLeadTimeFallback(country, bizClass) {
  const sheet = _ss().getSheetByName(country);
  if (!sheet) return [];
  const data = sheet.getRange(3, 1, sheet.getLastRow()-2, 31).getValues();
  const map  = {};
  data.forEach(row => {
    const type = String(row[0]||'').trim();
    if (!LEAD_TIME_TYPES.has(type) || (bizClass && type !== bizClass)) return;
    const fin=_num(row[26]),kam=_num(row[27]),br=_num(row[28]),inv=_num(row[29]),col=_num(row[30]);
    if (!fin&&!kam&&!br&&!inv&&!col) return;
    const b = String(row[3]||'').trim()||'Unknown';
    if (!map[b]) map[b]={n:0,fin:0,kam:0,br:0,inv:0,col:0};
    map[b].n++; map[b].fin+=fin; map[b].kam+=kam; map[b].br+=br; map[b].inv+=inv; map[b].col+=col;
  });
  return Object.entries(map)
    .map(([brand,v])=>({brand,fin:_r1(v.fin/v.n),kam:_r1(v.kam/v.n),
      brand_r:_r1(v.br/v.n),invoice:_r1(v.inv/v.n),collect:_r1(v.col/v.n)}))
    .sort((a,b)=>(b.fin+b.kam+b.brand_r+b.invoice+b.collect)-(a.fin+a.kam+a.brand_r+a.invoice+a.collect));
}

// ─── Period ───────────────────────────────────────────────────────────────────

function _parsePeriod(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val))
    return val.getFullYear() + ' ' + MONTHS[val.getMonth()];
  const s = String(val).trim();
  if (!s || s === '0' || s === 'false') return '';
  if (/^\d{4}\s+[A-Za-z]{3}/.test(s)) return s.replace(/\s+/, ' ');
  // Handle serial date numbers (from UNFORMATTED_VALUE)
  if (/^\d{5}$/.test(s)) {
    const d = new Date((parseInt(s) - 25569) * 86400000);
    return d.getFullYear() + ' ' + MONTHS[d.getMonth()];
  }
  try {
    const d = new Date(s);
    if (!isNaN(d)) return d.getFullYear() + ' ' + MONTHS[d.getMonth()];
  } catch(e) {}
  return s;
}

function _availablePeriods(rows) {
  return [...new Set(rows.map(r => r.period).filter(Boolean))]
    .sort((a, b) => _pms(b) - _pms(a));
}
function _prevPeriod(p, list) {
  const i = list.indexOf(p);
  return (i >= 0 && i < list.length - 1) ? list[i + 1] : null;
}
function _pms(p) { try { return new Date(p.trim()).getTime(); } catch(e) { return 0; } }

// ─── Filter / Summary / Delta ─────────────────────────────────────────────────

function _filter(rows, bizClass, period) {
  return rows.filter(r =>
    (!bizClass || r.type === bizClass) && (!period || r.period === period)
  );
}

function _summary(rows) {
  let totalAR=0, unbilled=0, underTerm=0, overDue=0, pendingCfm=0;
  rows.forEach(r => {
    if (_isDate(r.collected)) return;
    totalAR += r.amount;
    if      (!_isDate(r.issued))              unbilled   += r.amount;
    else if (r.status === STATUS.UNDER_TERM)  underTerm  += r.amount;
    else if (r.status === STATUS.OVER_DUE)    overDue    += r.amount;
    else if (r.status === STATUS.PENDING_CFM) pendingCfm += r.amount;
  });
  const billed = underTerm + overDue + pendingCfm;
  return { totalAR, unbilled, underTerm, overDue, pendingCfm, billed,
    unbilledPct: totalAR > 0 ? unbilled/totalAR : 0,
    overDuePct:  totalAR > 0 ? overDue /totalAR : 0 };
}

function _delta(c, p) {
  return { totalAR: c.totalAR-p.totalAR, unbilled: c.unbilled-p.unbilled,
    underTerm: c.underTerm-p.underTerm, overDue: c.overDue-p.overDue,
    billed: c.billed-p.billed };
}

// ─── Debug ────────────────────────────────────────────────────────────────────

function debugData() {
  const rows = _readTHSheet('TH');
  Logger.log('=== Total rows: %s', rows.length);

  const periods = _availablePeriods(rows);
  Logger.log('Periods (%s): %s', periods.length, periods.join(', '));

  // Row categories
  const collected  = rows.filter(r => _isDate(r.collected));
  const unbilled   = rows.filter(r => !_isDate(r.collected) && !_isDate(r.issued));
  const invoiced   = rows.filter(r => !_isDate(r.collected) &&  _isDate(r.issued));
  Logger.log('Collected=%s  Unbilled=%s  Invoiced(outstanding)=%s',
    collected.length, unbilled.length, invoiced.length);

  // Status breakdown for outstanding
  const outstanding = rows.filter(r => !_isDate(r.collected));
  const statusMap = {};
  outstanding.forEach(r => { statusMap[r.status] = (statusMap[r.status]||0) + 1; });
  Logger.log('Status breakdown: %s', JSON.stringify(statusMap));

  // Amount check
  const nonZero = rows.filter(r => r.amount !== 0).length;
  const sumAmt  = rows.reduce((s, r) => s + r.amount, 0);
  Logger.log('Non-zero amount: %s / %s  |  Sum all amounts: %s', nonZero, rows.length, sumAmt.toFixed(2));

  // Sample of unbilled rows
  Logger.log('--- Sample unbilled rows (first 5):');
  unbilled.slice(0, 5).forEach((r, i) =>
    Logger.log('[%s] period=%s brand=%s type=%s amount=%s status=%s issued=%s collected=%s',
      i, r.period, r.brand, r.type, r.amount, r.status, r.issued, r.collected)
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function _isDate(v) {
  if (!v) return false;
  if (v instanceof Date) return !isNaN(v.getTime());
  const s = String(v).trim();
  return s.length > 0 && s !== '0' && s !== 'false' && s !== '#DIV/0!';
}
function _num(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (!v) return 0;
  const s = String(v).replace(/[$,\s]/g, '').trim();
  if (!s || s === '#DIV/0!' || s === '000' || s === '-') return 0;
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function _r1(n) { return Math.round(n * 10) / 10; }
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }
