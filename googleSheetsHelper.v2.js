/**
 * Portfolio Rebalancing Script for Google Sheets
 * - Modular fetch functions with caching & error handling
 * - Health check for external data feeds
 * - Scoring model using macro indicators
 * - Enhanced GET_REBALANCE_SIGNAL and CHECK_DATA_FEEDS functions
 *
 * [2025-06-13] Update: Portfolio allocations updated and logic reviewed for new asset mix.
 * New target allocations (TFSA portfolio):
 *   CASH.TO 25.03%, XEQT.TO 15.81%, ZLB.TO 10.56%, T.TO 9.95%, SLF.TO 8.84%,
 *   BRK.B 8.28%, ESGV 7.28%, GUD.TO 4.57%, KSI.TO 3.96%, VIU.TO 3.89%, NVDA.TO 1.66%.
 * These values should be reflected in the 'Asset Allocations' sheet (column B as % of total).
 * Added entries for new tickers in 'Asset Metadata' sheet with Class, Region, Sensitivity:
 *   - BRK.B, ESGV, NVDA.TO as U.S. equities (Class='Equity', Region='U.S.');
 *   - T.TO, SLF.TO, ZLB.TO, GUD.TO, KSI.TO as Canadian equities (Class='Equity', Region='Canada');
 *   - XEQT.TO, VIU.TO as global ETFs (Class='Equity', Region='Global').
 * For sensitivity: higher volatility tech stocks (NVDA, KSI) may use sens ~1.2; low-vol/defensive stocks (Telus, ZLB)
 * sens ~0.8-0.9; others sens ~1.0. Macro indicator logic unchanged (U.S./Canada economic signals and VIX guide shifts).
 */
 
// -----------------------------------------------------------------------------
// Configuration Constants
// -----------------------------------------------------------------------------
const CACHE_VERSION = 'v3';
const SCRIPT_PROP = PropertiesService.getScriptProperties();
const CONFIG = {
  // API Keys
  FRED_API_KEY: SCRIPT_PROP.getProperty('FRED_API_KEY'),
 
  // Series IDs
  FRED_SERIES: {
    unemployment: 'UNRATE',
    usCPI:        'CPIAUCSL',
    canCPI:       'CANCPIALLMINMEI',
    canGDP:       'NGDPRSAXDCCAQ',
    yieldCurve:   'T10Y2Y',
    creditSpread: 'BAMLC0A0CM',
    vix:          'VIXCLS'
  },
 
  STATCAN_VECTOR: 41690973,  // Canadian CPI vectorId
  YAHOO_TICKERS: { vix: '%5EVIX' },
 
  // Scoring thresholds
  THRESHOLDS: {
    unemployment: { veryLow: 3.5, low: 4.5, high: 5.5, veryHigh: 6.5 },
    cpi:          { veryLow: 1.0, target: 2.0, tolerance: 0.5, high: 4.0 },
    vix:          { low: 15, normal: 20, elevated: 25, high: 30 }
  },
 
  // Indicator weights
  WEIGHTS: {
    unemployment: 1.2,
    usCPI:        1.2,
    canCPI:       0.8,
    canGDP:       1.0,
    vix:          1.5,
    yieldCurve:   1.7,
    creditSpread: 1.4
  },
 
  // Trend lookback
  TREND_PERIODS: 3,
 
  // Rebalancing parameters
  REBALANCING: {
    minThreshold: 0.005,    // 0.5%
    maxSingleMove: 0.15,    // 15%
    volatilityAdjustment: true,
    balanceConstraint:     true
  }
};
 
// -----------------------------------------------------------------------------
// Utility: Fetch from cache or call function
// -----------------------------------------------------------------------------
function withCache(key, ttlSec, fetchFunction) {
  const cacheKey = `${CACHE_VERSION}_${key}`;
  const cache    = CacheService.getScriptCache();
  const cached   = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  const value = fetchFunction();
  cache.put(cacheKey, JSON.stringify(value), ttlSec);
  return value;
}
 
// -----------------------------------------------------------------------------
// Data Fetchers (with error handling & caching)
// -----------------------------------------------------------------------------
function fetchFredLatest(seriesId) {
  return withCache('FRED_' + seriesId, 3600, () => {
    try {
      const url = `https://api.stlouisfed.org/fred/series/observations` +
                  `?series_id=${seriesId}` +
                  `&api_key=${CONFIG.FRED_API_KEY}` +
                  `&file_type=json&limit=1&sort_order=desc`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) throw new Error(resp.getResponseCode());
      const data = JSON.parse(resp.getContentText());
      const obs = data.observations;
      if (!obs || !obs.length) throw new Error('No obs');
      return { value: parseFloat(obs[0].value), date: obs[0].date, ok: true };
    } catch (e) {
      Logger.log(`FRED ${seriesId} fetch error: ${e}`);
      return { value: null, date: null, ok: false };
    }
  });
}
 
function fetchStatCanCPI() {
  return withCache('STATCAN_CPI', 21600, () => {
    try {
      const url = 'https://www150.statcan.gc.ca/t1/wds/rest/getDataFromVectorsAndLatestNPeriods';
      const payload = JSON.stringify([{ vectorId: CONFIG.STATCAN_VECTOR, latestN: 1 }]);
      const opts = { method: 'post', contentType: 'application/json', payload, muteHttpExceptions: true };
      const resp = UrlFetchApp.fetch(url, opts);
      if (resp.getResponseCode() !== 200) throw new Error(resp.getResponseCode());
      const json = JSON.parse(resp.getContentText());
      const pts = json[0].object.vectorDataPoint;
      if (!pts || !pts.length) throw new Error('No data');
      const latest = pts[pts.length - 1];
      return { value: parseFloat(latest.value), date: latest.refPerRaw, ok: true };
    } catch (e) {
      Logger.log(`StatCan CPI fetch error: ${e}`);
      return { value: null, date: null, ok: false };
    }
  });
}
 
function fetchYahooVIX() {
  return withCache('YH_VIX', 3600, () => {
    try {
      const url = `https://query1.finance.yahoo.com/v8/finance/chart/${CONFIG.YAHOO_TICKERS.vix}` +
                  `?range=5d&interval=1d`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) throw new Error(resp.getResponseCode());
      const c = JSON.parse(resp.getContentText());
      const quotes = c.chart.result[0].indicators.quote[0].close;
      const latest = quotes.pop();
      if (latest == null) throw new Error('No data');
      return { value: Number(latest.toFixed(2)), date: null, ok: true };
    } catch (e) {
      Logger.log(`Yahoo VIX fetch error: ${e}`);
      return { value: null, date: null, ok: false };
    }
  });
}
 
// Historical series for trend analysis
function fetchFredSeries(seriesId, count) {
  return withCache(`FRED_SERIES_${seriesId}_${count}`, 21600, () => {
    try {
      const url = `https://api.stlouisfed.org/fred/series/observations` +
                  `?series_id=${seriesId}` +
                  `&api_key=${CONFIG.FRED_API_KEY}` +
                  `&file_type=json&limit=${count}&sort_order=desc`;
      const resp = UrlFetchApp.fetch(url);
      const data = JSON.parse(resp.getContentText());
      const obs = data.observations
        .filter(o => o.value !== '.' && !isNaN(+o.value))
        .map(o => ({ date: o.date, value: +o.value }))
        .reverse(); // old→new
      return obs;
    } catch (e) {
      Logger.log(`FRED series ${seriesId} error: ${e}`);
      return [];
    }
  });
}
 
// -----------------------------------------------------------------------------
// Indicator Gathering and Health Check
// -----------------------------------------------------------------------------
function getIndicatorData() {
  // Fetch current values
  const u = fetchFredLatest(CONFIG.FRED_SERIES.unemployment);
  const usCPI = fetchFredLatest(CONFIG.FRED_SERIES.usCPI);
  const caCPI = fetchStatCanCPI();
  const gdp  = fetchFredLatest(CONFIG.FRED_SERIES.canGDP);
  const yc   = fetchFredLatest(CONFIG.FRED_SERIES.yieldCurve);
  const cs   = fetchFredLatest(CONFIG.FRED_SERIES.creditSpread);
  const vix  = fetchYahooVIX();
 
  return { unemployment: u, usCPI, canCPI: caCPI, canGDP: gdp, yieldCurve: yc, creditSpread: cs, vix };
}
 
/**
 * Health check: returns [[metric, status], ...] for sheet
 */
function CHECK_DATA_FEEDS() {
  const ind = getIndicatorData();
  const rows = [];
  for (const key in ind) {
    const obj = ind[key];
    rows.push([key, obj.ok ? 'OK' : 'ERROR']);
  }
  return rows;
}
 
// -----------------------------------------------------------------------------
// Trend Calculation
// -----------------------------------------------------------------------------
function calculateTrend(historical, periods) {
  if (!historical || historical.length < periods + 1) return 0;
  const older = historical.slice(0, historical.length - periods);
  const recent = historical.slice(-periods);
  if (!older.length) return 0;
  const avg = arr => arr.reduce((a,b) => a+b,0)/arr.length;
  const change = (avg(recent) - avg(older)) / avg(older);
  return Math.max(-1, Math.min(1, change * 10));
}
 
// -----------------------------------------------------------------------------
// Scoring Functions
// -----------------------------------------------------------------------------
function calculateUnemploymentScore(curr, hist) {
  if (curr == null) return { score:0, desc:'No data' };
  const t = CONFIG.THRESHOLDS.unemployment;
  let base=0, d=`${curr}%`;
  if (curr < t.veryLow)      { base=2; d+=' (Very Low)'; }
  else if (curr < t.low)     { base=1; d+=' (Low)'; }
  else if (curr > t.veryHigh){ base=-2; d+=' (Very High)'; }
  else if (curr > t.high)    { base=-1; d+=' (High)'; }
  else                       { base=0; d+=' (Normal)'; }
  const adj = calculateTrend(hist, CONFIG.TREND_PERIODS);
  if (adj) { base -= adj*0.5; d += adj>0?' Rising':' Falling'; }
  return { score:Math.max(-2,Math.min(2,base)), description:d };
}
 
function calculateCPIScore(curr, hist, region) {
  if (curr==null) return { score:0, desc:'No data' };
  const t = CONFIG.THRESHOLDS.cpi;
  let base=0, d=`${curr}%`;
  if (curr < t.veryLow)             { base=-1; d+=' (Deflation Risk)'; }
  else if (Math.abs(curr - t.target) <= t.tolerance) { base=1; d+=' (On Target)'; }
  else if (curr > t.high)           { base=-2; d+=' (High Inflation)'; }
  else if (curr > t.target + t.tolerance){ base=-1; d+=' (Above Target)'; }
  else                               { base=0; d+=' (Below Target)'; }
  const adj = calculateTrend(hist, CONFIG.TREND_PERIODS);
  if (adj) { base -= Math.sign(adj)*0.3; d+= adj>0?' Rising':' Falling'; }
  return { score:Math.max(-2,Math.min(2,base)), description:d };
}
 
function calculateVIXScore(curr, hist) {
  if (curr==null) return { score:0, desc:'No data' };
  const t = CONFIG.THRESHOLDS.vix;
  let base=0, d=`${curr}`;
  if (curr < t.low)        { base=1; d+=' (Low Vol)'; }
  else if (curr < t.normal){ base=0.5; d+=' (Normal)'; }
  else if (curr < t.elevated){ base=-0.5; d+=' (Elevated)'; }
  else if (curr < t.high)   { base=-1; d+=' (High)'; }
  else                       { base=-2; d+=' (Very High)'; }
  const adj = calculateTrend(hist, CONFIG.TREND_PERIODS);
  if (adj) { base -= adj*0.4; d += adj>0?' Rising':' Falling'; }
  return { score:Math.max(-2,Math.min(2,base)), description:d };
}
 
function calculateGDPScore(curr) {
  if (!curr) return { score:0, desc:'No data' };
  const txt = curr.toString().toLowerCase();
  const up = ['rising','positive','increasing','expanding'];
  const down = ['falling','negative','decreasing','contracting'];
  let base=0;
  if (up.some(w=>txt.includes(w)))      base=1;
  else if (down.some(w=>txt.includes(w))) base=-1;
  return { score:base, description:curr };
}
 
function calculateYieldScore(curr, hist) {
  if (curr==null) return { score:0, desc:'No data' };
  let base=0, d=`${curr}%`;
  if      (curr > 0.5)  { base=1; d+=' (Positive)'; }
  else if (curr > -0.5) { base=0; d+=' (Flat)'; }
  else if (curr > -1.0) { base=-1; d+=' (Inverted)'; }
  else                  { base=-2; d+=' (Deep Inversion)'; }
  const adj = calculateTrend(hist, CONFIG.TREND_PERIODS);
  if (adj) { base += adj*0.3; d += adj>0?' Steepening':' Flattening'; }
  return { score:Math.max(-2,Math.min(2,base)), description:d };
}
 
function calculateCreditScore(curr, hist) {
  if (curr==null) return { score:0, desc:'No data' };
  let base=0, d=`${curr}%`;
  if      (curr < 1.0) base=1, d+=' (Low)';
  else if (curr < 2.0) base=0, d+=' (Normal)';
  else if (curr < 3.0) base=-1, d+=' (Elevated)';
  else                 base=-2, d+=' (High)';
  const adj = calculateTrend(hist, CONFIG.TREND_PERIODS);
  if (adj) { base -= adj*0.4; d += adj>0?' Widening':' Tightening'; }
  return { score:Math.max(-2,Math.min(2,base)), description:d };
}
 
// -----------------------------------------------------------------------------
// Aggregation: Enhanced Market Data
// -----------------------------------------------------------------------------
function getEnhancedMarketData() {
  const ind = getIndicatorData();
  const hist = {
    unemployment: fetchFredSeries(CONFIG.FRED_SERIES.unemployment, 6).map(o=>o.value),
    usCPI:        fetchFredSeries(CONFIG.FRED_SERIES.usCPI, 6).map(o=>o.value),
    canCPI:       fetchFredSeries(CONFIG.FRED_SERIES.canCPI, 6).map(o=>o.value),
    canGDP:       [],
    yieldCurve:   fetchFredSeries(CONFIG.FRED_SERIES.yieldCurve, 6).map(o=>o.value),
    creditSpread: fetchFredSeries(CONFIG.FRED_SERIES.creditSpread, 6).map(o=>o.value),
    vix:          fetchFredSeries(CONFIG.FRED_SERIES.vix, 180).map(o=>o.value) // monthly avg not implemented here
  };
  const scores = {
    unemployment: calculateUnemploymentScore(ind.unemployment.value, hist.unemployment).score,
    usCPI:        calculateCPIScore(ind.usCPI.value, hist.usCPI, 'US').score,
    canCPI:       calculateCPIScore(ind.canCPI.value, hist.canCPI, 'CA').score,
    canGDP:       calculateGDPScore(ind.canGDP.value).score,
    vix:          calculateVIXScore(ind.vix.value, hist.vix).score,
    yieldCurve:   calculateYieldScore(ind.yieldCurve.value, hist.yieldCurve).score,
    creditSpread: calculateCreditScore(ind.creditSpread.value, hist.creditSpread).score
  };
  return { indicators: ind, scores, config: CONFIG };
}
 
// -----------------------------------------------------------------------------
// Main: GET_REBALANCE_SIGNAL(assetTicker)
// -----------------------------------------------------------------------------
function GET_REBALANCE_SIGNAL(assetTicker) {
  const ss = SpreadsheetApp.getActive();
  const allocSh = ss.getSheetByName('Asset Allocations');
  const metaSh = ss.getSheetByName('Asset Metadata');
 
  // Load allocations
  const allocData = allocSh.getRange('A2:B' + allocSh.getLastRow()).getValues();
  const allocMap = {}, tickers = [];
  allocData.forEach(r => {
    const t = r[0], raw = r[1];
    if (!t || raw === '') return;
    let a = parseFloat((''+raw).replace('%',''));
    // if the user typed 5.14 (no %), assume it was meant to be a percent
    if (a > 1) a = a/100;
    // now a is in [0..1]
    allocMap[t] = a;
    tickers.push(t);
  });
 
  // Load metadata
  const meta = metaSh.getRange('A2:D' + metaSh.getLastRow()).getValues();
  const classMap={}, regionMap={}, sensMap={};
  meta.forEach(r=>{ 
    classMap[r[0]] = r[1]; 
    regionMap[r[0]] = r[2]; 
    sensMap[r[0]] = parseFloat(r[3]) || 1; 
  });

  // Fetch market scores
  const mkt = getEnhancedMarketData();
  const scores = mkt.scores;
  const cfg = mkt.config;
 
  // Regional score (weighted average of indicators based on region)
  function getRegionalScore(t) {
    const region = regionMap[t] || 'Global';
    let total = 0, wsum = 0;
    if (region === 'U.S.') {
      [['unemployment',1.5], ['usCPI',1.3], ['vix',1.2], ['yieldCurve',1.1], ['creditSpread',1.1]]
        .forEach(([k, m]) => { total += scores[k] * cfg.WEIGHTS[k] * m; wsum += cfg.WEIGHTS[k] * m; });
    } else if (region === 'Canada') {
      [['canCPI',1.5], ['canGDP',1.4], ['unemployment',0.7], ['vix',0.8]]
        .forEach(([k, m]) => { total += scores[k] * cfg.WEIGHTS[k] * m; wsum += cfg.WEIGHTS[k] * m; });
    } else {
      for (const k in scores) {
        if (cfg.WEIGHTS[k] != null) { 
          total += scores[k] * cfg.WEIGHTS[k]; 
          wsum += cfg.WEIGHTS[k]; 
        }
      }
    }
    return wsum ? total / wsum : 0;
  }
 
  // Delta calculation for a given asset
  function calcDelta(t) {
    const regionScore = getRegionalScore(t);
    let shift = 0;
    if      (regionScore >= 1.5)  shift = 0.08;
    else if (regionScore >= 0.75) shift = 0.05;
    else if (regionScore >= 0.25) shift = 0.025;
    else if (regionScore >= -0.25) shift = 0;
    else if (regionScore >= -0.75) shift = -0.025;
    else if (regionScore >= -1.5) shift = -0.05;
    else                          shift = -0.08;
    shift *= sensMap[t];  // adjust for asset sensitivity
    if (cfg.REBALANCING.volatilityAdjustment) {
      const vixScore = scores.vix;
      if (vixScore < -1) shift *= 0.7;  // reduce shifts in very low-volatility regime
    }
    // Invert shift for defensive assets (they move opposite to risk appetite)
    return (classMap[t] === 'Defensive') ? -shift : shift;
  }
 
  // Compute raw recommended shifts for all assets
  const rawDelta = {}, defList = [], eqList = [];
  tickers.forEach(t => {
    rawDelta[t] = calcDelta(t);
    if (classMap[t] === 'Defensive') defList.push(t);
    else                             eqList.push(t);
  });
 
  // Balance offsets: ensure net sum of changes ~ 0 (no new capital added)
  if (cfg.REBALANCING.balanceConstraint) {
    let net = 0, defAlloc = 0, eqAlloc = 0;
    defList.forEach(t => { net += rawDelta[t] * allocMap[t]; defAlloc += allocMap[t]; });
    eqList.forEach(t  => { net += rawDelta[t] * allocMap[t]; eqAlloc  += allocMap[t]; });
    if (Math.abs(net) > 0.001) {
      const adj = -net / 2;
      if (defAlloc) defList.forEach(t => { rawDelta[t] += adj / defAlloc; });
      if (eqAlloc)  eqList.forEach(t => { rawDelta[t] += adj / eqAlloc; });
    }
  }
 
  // Normalize and get final recommended allocation for the target asset
  const sumNew = tickers.reduce((sum, t) => 
                   sum + Math.max(0, allocMap[t] * (1 + rawDelta[t])), 0);
  const oldAlloc = allocMap[assetTicker];
  const newAlloc = Math.max(0, oldAlloc * (1 + rawDelta[assetTicker])) / sumNew;
  const finalDelta = newAlloc / oldAlloc - 1;
 
  // Format output signal
  if (Math.abs(finalDelta) < cfg.REBALANCING.minThreshold) {
    return 'Hold';
  }

  return (finalDelta > 0)
    ? `Increase ${(finalDelta * 100).toFixed(2)}%`
    : `Decrease ${(Math.abs(finalDelta) * 100).toFixed(2)}%`;
}


/** Debug: spit out what Apps Script thinks your FRED key is */
function DEBUG_GET_KEY() {
  return PropertiesService.getScriptProperties().getProperty('FRED_API_KEY');
}

function TEST_FRED_FETCH() {
  const key = PropertiesService.getScriptProperties().getProperty('FRED_API_KEY').trim();
  const url = `https://api.stlouisfed.org/fred/series/observations`
            + `?series_id=UNRATE`
            + `&api_key=${key}`
            + `&file_type=json&limit=1&sort_order=desc`;
  let resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('HTTP ' + resp.getResponseCode());
  Logger.log(resp.getContentText().substring(0,200));  // first 200 chars
  return `HTTP ${resp.getResponseCode()}`;
}

/**
 * Clears any cached metric data so that fresh calls happen immediately.
 */
function CLEAR_ALL_CACHE() {
  const cache = CacheService.getScriptCache();
  // FRED series
  Object.values(CONFIG.FRED_SERIES).forEach(id => cache.remove('FRED_' + id));
  // StatCan CPI
  cache.remove('STATCAN_CPI');
  // Yahoo VIX
  cache.remove('YH_VIX');
}

function TEST_FRED_LATEST() {
  const result = fetchFredLatest(CONFIG.FRED_SERIES.unemployment);
  Logger.log(result);
}

/**
 * Returns the raw FRED JSON response for UNRATE so you can inspect any error message.
 * Usage: run this in the Apps Script editor, then View → Logs
 * Or call =DEBUG_FRED_RESPONSE() in your sheet (it will spill the JSON text).
 */
function DEBUG_FRED_RESPONSE() {
  const key = PropertiesService.getScriptProperties()
                .getProperty('FRED_API_KEY')
                .trim();
  const url =
    'https://api.stlouisfed.org/fred/series/observations' +
    '?series_id=UNRATE' +
    `&api_key=${key}` +
    '&file_type=json&limit=1&sort_order=desc';
  
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  // Log it in the editor
  Logger.log(resp.getContentText());
  // Return it to the sheet (if you want)
  return resp.getContentText();
}
