function getAmericanUnemploymentRate() {
  const url = `https://api.stlouisfed.org/fred/series/observations?series_id=UNRATE&api_key=${FRED_API_KEY}&file_type=json`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  const latest = data.observations.pop();
  return parseFloat(latest.value);
}

function getUsCpi() {
  const url = `https://api.stlouisfed.org/fred/series/observations?series_id=CPIAUCNS&api_key=${FRED_API_KEY}&file_type=json`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  const latest = data.observations.at(-1);
  return parseFloat(latest.value);
}

function getUsCpiYearOverYear() {
  const url = `https://api.stlouisfed.org/fred/series/observations?series_id=CPIAUCNS&api_key=${FRED_API_KEY}&file_type=json`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  const obs = data.observations;
  const latest = parseFloat(obs.at(-1).value);
  const yearAgo = parseFloat(obs.at(-13).value); // 12 months ago
  const yoy = ((latest - yearAgo) / yearAgo) * 100;
  return Math.round(yoy * 10) / 10; // rounded to 1 decimal place
}

/**
 * Returns the latest YoY % change in the Canadian CPI.
 * 
 * Usage in your sheet:
 *   =GET_CPI_YOY()       // e.g. 1.7
 *   =GET_CPI_YOY(TRUE)   // raw HTTP status + JSON for debugging
 */
function GET_CPI_YOY(debug) {
  const url = 'https://www150.statcan.gc.ca/t1/wds/rest/getDataFromVectorsAndLatestNPeriods';
  const payload = JSON.stringify([
    { vectorId: 41690973, latestN: 13 }
  ]);
  const options = {
    method:             'post',
    contentType:        'application/json',
    payload:            payload,
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText();

  // debug mode: spit back raw HTTP + JSON
  if (debug === true || debug === 'TRUE') {
    return `HTTP ${code}\n` + text;
  }

  // parse
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    return `Parse error (HTTP ${code}): ${e.message}`;
  }

  try {
    // vectorDataPoint is an array of {refPer, refPerRaw, value, …}, oldest→newest
    const points = data[0].object.vectorDataPoint;
    if (!Array.isArray(points) || points.length < 2) {
      throw new Error('not enough data points');
    }

    // latest = last element in that array
    const latest = points[points.length - 1];
    // refPerRaw is YYYY-MM-DD
    const [year, month, day] = latest.refPerRaw.split('-');
    const lastYearRef = `${Number(year) - 1}-${month}-${day}`;

    // find the datapoint exactly 12 months ago
    const prior = points.find(dp => dp.refPerRaw === lastYearRef);
    if (!prior) {
      throw new Error(`no data for ${lastYearRef}`);
    }

    // compute ((new/old) − 1) × 100
    const yoy = ((latest.value / prior.value) - 1) * 100;
    return Number(yoy.toFixed(1));
  } catch (e) {
    return `Error extracting YoY% (HTTP ${code}): ${e.message}`;
  }
}

/**
 * Returns "Rising" if latest Canada Real GDP > previous quarter,
 * or "Shrinking" if it’s lower (or "No change" if equal).
 *
 * Usage in your sheet:
 *   =GET_CANADIAN_GDP_TREND()
 */
function GET_CANADIAN_GDP_TREND() {
  const fredKey = FRED_API_KEY;  // reuse your existing FRED key
  const series = 'NGDPRSAXDCCAQ'; // Real GDP, Canada, quarterly, SA  [oai_citation:0‡FRED](https://fred.stlouisfed.org/series/NGDPRSAXDCCAQ?utm_source=chatgpt.com)
  const url = `https://api.stlouisfed.org/fred/series/observations?series_id=${series}&api_key=${fredKey}&file_type=json`;

  try {
    const resp = UrlFetchApp.fetch(url);
    const data = JSON.parse(resp.getContentText());
    const obs = data.observations;
    if (obs.length < 2) throw new Error("Not enough data points");

    const latest = parseFloat(obs.at(-1).value);
    const prev   = parseFloat(obs.at(-2).value);

    if (latest > prev)   return "Rising";
    if (latest < prev)   return "Shrinking";
                         return "No change";
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Fetches the latest CBOE Volatility Index (VIX) closing price
 * by scraping Yahoo Finance’s chart API.
 *
 * =GET_VIX()  // e.g. 18.52
 */
function GET_VIX() {
  try {
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/%5EVIX?range=5d&interval=1d';
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(resp.getContentText());
    const quotes = json.chart.result[0].indicators.quote[0].close;
    const latest = quotes[quotes.length - 1];
    if (latest == null) throw new Error('No data in response');
    return Number(latest.toFixed(2));
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Returns the 10Y−2Y U.S. Treasury yield spread, in percentage points.
 *
 * =GET_YIELD_CURVE()  // e.g. -0.45  (an inverted curve)
 *
 * Source: FRED series GS10 (10-year) and GS2 (2-year)
 */
function GET_YIELD_CURVE() {
  try {
    const key = FRED_API_KEY;
    // Fetch 10-year
    const tenUrl = `https://api.stlouisfed.org/fred/series/observations?series_id=GS10&api_key=${key}&file_type=json`;
    const twoUrl = `https://api.stlouisfed.org/fred/series/observations?series_id=GS2&api_key=${key}&file_type=json`;
    const ten = JSON.parse(UrlFetchApp.fetch(tenUrl).getContentText())
                    .observations.slice(-1)[0].value;
    const two = JSON.parse(UrlFetchApp.fetch(twoUrl).getContentText())
                    .observations.slice(-1)[0].value;
    const spread = parseFloat(ten) - parseFloat(two);
    return Number(spread.toFixed(2));
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Returns the spread (in percentage points) between
 * Moody’s Baa corporate bond yield (series BAA)
 * and the 10-year U.S. Treasury (GS10).
 *
 * =GET_CREDIT_SPREAD()  // e.g. 1.75
 *
 * This is a common proxy for corporate credit risk.
 */
function GET_CREDIT_SPREAD() {
  try {
    const key = FRED_API_KEY;
    const baaUrl = `https://api.stlouisfed.org/fred/series/observations?series_id=BAA&api_key=${key}&file_type=json`;
    const tenUrl = `https://api.stlouisfed.org/fred/series/observations?series_id=GS10&api_key=${key}&file_type=json`;
    const baa = JSON.parse(UrlFetchApp.fetch(baaUrl).getContentText())
                    .observations.slice(-1)[0].value;
    const ten = JSON.parse(UrlFetchApp.fetch(tenUrl).getContentText())
                    .observations.slice(-1)[0].value;
    const spread = parseFloat(baa) - parseFloat(ten);
    return Number(spread.toFixed(2));
  } catch (e) {
    return `Error: ${e.message}`;
  }
}


/**
 * Fetches historical data for all indicators and returns formatted data
 * for trend analysis. Returns last 6 data points for each indicator.
 * 
 * Usage: =GET_HISTORICAL_INDICATORS()
 * This will populate a range with historical data that can be used by the enhanced functions
 */
function GET_HISTORICAL_INDICATORS() {
  const fredKey = FRED_API_KEY;
  
  // FRED series IDs for each indicator
  const seriesMap = {
    'unemployment': 'UNRATE',           // US Unemployment Rate
    'usCPI': 'CPIAUCSL',               // US CPI (annual % change)
    'canCPI': 'CANCPIALLMINMEI',       // Canada CPI 
    'canGDP': 'NGDPRSAXDCCAQ',         // Canada Real GDP
    'vix': 'VIXCLS',                   // VIX (daily, we'll get monthly avg)
    'yieldCurve': 'T10Y2Y',            // 10Y-2Y Treasury Spread
    'creditSpread': 'BAMLC0A0CM'       // Investment Grade Credit Spread
  };
  
  const results = [];
  const headers = ['Date', 'Unemployment', 'US_CPI', 'Can_CPI', 'Can_GDP', 'VIX', 'Yield_Curve', 'Credit_Spread'];
  results.push(headers);
  
  try {
    // Fetch data for each series
    const allData = {};
    
    for (const [indicator, seriesId] of Object.entries(seriesMap)) {
      const historicalData = fetchFredSeries(seriesId, fredKey, 12); // Get 12 months of data
      allData[indicator] = historicalData;
    }
    
    // Find common dates and build rows
    const dates = getCommonDates(allData);
    
    dates.forEach(date => {
      const row = [date];
      
      // Add each indicator's value for this date
      ['unemployment', 'usCPI', 'canCPI', 'canGDP', 'vix', 'yieldCurve', 'creditSpread'].forEach(indicator => {
        const dataPoint = allData[indicator].find(d => d.date === date);
        row.push(dataPoint ? dataPoint.value : null);
      });
      
      results.push(row);
    });
    
    return results;
    
  } catch (e) {
    return [['Error', e.message]];
  }
}

/**
 * Helper function to fetch a FRED series with specified number of observations
 */
function fetchFredSeries(seriesId, apiKey, count = 12) {
  const url = `https://api.stlouisfed.org/fred/series/observations?series_id=${seriesId}&api_key=${apiKey}&file_type=json&limit=${count}&sort_order=desc`;
  
  try {
    const resp = UrlFetchApp.fetch(url);
    const data = JSON.parse(resp.getContentText());
    const observations = data.observations;
    
    return observations
      .filter(obs => obs.value !== '.' && !isNaN(parseFloat(obs.value)))
      .map(obs => ({
        date: obs.date,
        value: parseFloat(obs.value)
      }))
      .reverse(); // Most recent first
      
  } catch (e) {
    console.log(`Error fetching ${seriesId}: ${e.message}`);
    return [];
  }
}

/**
 * Find dates that are common across multiple series (or closest matches)
 */
function getCommonDates(allData) {
  // Get all unique dates
  const allDates = new Set();
  Object.values(allData).forEach(series => {
    series.forEach(point => allDates.add(point.date));
  });
  
  // Sort dates and take the most recent 6
  return Array.from(allDates)
    .sort((a, b) => new Date(b) - new Date(a))
    .slice(0, 6);
}

/**
 * Alternative: Get specific indicator history
 * Usage: =GET_UNEMPLOYMENT_HISTORY() returns last 6 unemployment rates
 */
function GET_UNEMPLOYMENT_HISTORY() {
  const fredKey = FRED_API_KEY;
  const seriesId = 'UNRATE';
  
  try {
    const data = fetchFredSeries(seriesId, fredKey, 6);
    return data.map(d => [d.date, d.value]);
  } catch (e) {
    return [['Error', e.message]];
  }
}

/**
 * Get VIX history (monthly averages to match other monthly data)
 */
function GET_VIX_HISTORY() {
  const fredKey = FRED_API_KEY;
  const seriesId = 'VIXCLS';
  
  try {
    // Get daily VIX data for last 6 months, then calculate monthly averages
    const url = `https://api.stlouisfed.org/fred/series/observations?series_id=${seriesId}&api_key=${fredKey}&file_type=json&limit=180&sort_order=desc`;
    
    const resp = UrlFetchApp.fetch(url);
    const data = JSON.parse(resp.getContentText());
    const observations = data.observations;
    
    // Group by month and calculate averages
    const monthlyData = {};
    
    observations.forEach(obs => {
      if (obs.value !== '.' && !isNaN(parseFloat(obs.value))) {
        const date = new Date(obs.date);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        
        if (!monthlyData[monthKey]) {
          monthlyData[monthKey] = [];
        }
        monthlyData[monthKey].push(parseFloat(obs.value));
      }
    });
    
    // Calculate monthly averages and return last 6 months
    const results = Object.entries(monthlyData)
      .map(([month, values]) => [
        month + '-01', // Use first day of month for date
        values.reduce((a, b) => a + b, 0) / values.length
      ])
      .sort((a, b) => new Date(b[0]) - new Date(a[0]))
      .slice(0, 6);
    
    return results;
    
  } catch (e) {
    return [['Error', e.message]];
  }
}

/**
 * Enhanced version of getHistoricalData that uses FRED API
 * This replaces the sheet-based version in the enhanced signal function
 */
function getHistoricalDataFromFRED(ss) {
  const fredKey = FRED_API_KEY;
  
  if (!fredKey) {
    console.log("No FRED API key available, skipping historical data");
    return null;
  }
  
  try {
    const seriesMap = {
      'unemployment': 'UNRATE',
      'usCPI': 'CPIAUCSL', 
      'canCPI': 'CANCPIALLMINMEI',
      'vix': 'VIXCLS'
    };
    
    const historical = {};
    
    // Fetch each series
    Object.entries(seriesMap).forEach(([key, seriesId]) => {
      try {
        const data = fetchFredSeries(seriesId, fredKey, 6);
        historical[key] = data.map(d => d.value);
      } catch (e) {
        console.log(`Failed to fetch ${key}: ${e.message}`);
        historical[key] = null;
      }
    });
    
    return historical;
    
  } catch (e) {
    console.log("Error fetching historical data from FRED: " + e.message);
    return null;
  }
}

/**
 * Enhanced market signal function with dynamic thresholds, weighting, and trend analysis
 * Reads from 'Indicators' sheet (A2:B8) and 'Config' sheet for parameters
 * Also reads historical data from 'Historical' sheet for trend analysis
 */
function GET_SIGNAL_STRENGTH() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get current indicator values
  var indicatorData = getIndicatorData(ss);
  
  // Get configuration parameters (thresholds, weights)
  var config = getConfigData(ss);
  
  // Get historical data for trend analysis (null if no Historical sheet)
  var historical = getHistoricalDataFromFRED();
  
  // Calculate weighted scores with trend adjustment
  var totalScore = 0;
  var totalWeight = 0;
  var details = [];
  
  // Unemployment analysis
  var unempScore = calculateUnemploymentScore(indicatorData.unemployment, historical ? historical.unemployment : null, config);
  totalScore += unempScore.score * config.weights.unemployment;
  totalWeight += config.weights.unemployment;
  details.push("Unemployment: " + unempScore.description);
  
  // CPI analysis (US)
  var cpiScore = calculateCPIScore(indicatorData.usCPI, historical ? historical.usCPI : null, config, 'US');
  totalScore += cpiScore.score * config.weights.usCPI;
  totalWeight += config.weights.usCPI;
  details.push("US CPI: " + cpiScore.description);
  
  // CPI analysis (Canada)
  var canCPIScore = calculateCPIScore(indicatorData.canCPI, historical ? historical.canCPI : null, config, 'CAN');
  totalScore += canCPIScore.score * config.weights.canCPI;
  totalWeight += config.weights.canCPI;
  details.push("Can CPI: " + canCPIScore.description);
  
  // VIX analysis
  var vixScore = calculateVIXScore(indicatorData.vix, historical ? historical.vix : null, config);
  totalScore += vixScore.score * config.weights.vix;
  totalWeight += config.weights.vix;
  details.push("VIX: " + vixScore.description);
  
  // GDP analysis
  var gdpScore = calculateGDPScore(indicatorData.canGDP, historical ? historical.canGDP : null, config);
  totalScore += gdpScore.score * config.weights.canGDP;
  totalWeight += config.weights.canGDP;
  details.push("GDP: " + gdpScore.description);
  
  // Yield Curve analysis
  var yieldScore = calculateYieldScore(indicatorData.yieldCurve, historical ? historical.yieldCurve : null, config);
  totalScore += yieldScore.score * config.weights.yieldCurve;
  totalWeight += config.weights.yieldCurve;
  details.push("Yield: " + yieldScore.description);
  
  // Credit Spread analysis
  var creditScore = calculateCreditScore(indicatorData.creditSpread, historical ? historical.creditSpread : null, config);
  totalScore += creditScore.score * config.weights.creditSpread;
  totalWeight += config.weights.creditSpread;
  details.push("Credit: " + creditScore.description);
  
  // Calculate final weighted score
  var finalScore = totalScore / totalWeight;
  
  // Determine signal strength
  var signal = determineSignal(finalScore, config);
  
  // Return detailed result or simple result based on config
  if (config.showDetails) {
    return signal + " | Details: " + details.join("; ");
  } else {
    return signal;
  }
}

function getIndicatorData(ss) {
  var sh = ss.getSheetByName("Indicators");
  var data = sh.getRange("A2:B8").getValues();
  
  var map = {};
  data.forEach(function(row) {
    if (row[0] && row[1] !== "") {
      map[row[0]] = row[1];
    }
  });
  
  return {
    unemployment: parseFloat(map["U.S. Unemployment"]) || null,
    usCPI: parseFloat(map["U.S. CPI"]) || null,
    canCPI: parseFloat(map["Can CPI"]) || null,
    canGDP: map["Can GDP"] || null,
    vix: parseFloat(map["VIX"]) || null,
    yieldCurve: parseFloat(map["Yield Curve"]) || null,
    creditSpread: parseFloat(map["Credit Spread"]) || null
  };
}

function getConfigData(ss) {
  // Default configuration - can be overridden by 'Config' sheet
  var defaultConfig = {
    thresholds: {
      unemployment: {
        veryLow: 3.5,    // Below this = strong growth signal
        low: 4.5,        // Below this = mild growth signal
        high: 5.5,       // Above this = mild defensive signal
        veryHigh: 6.5    // Above this = strong defensive signal
      },
      cpi: {
        veryLow: 1.0,    // Below this = strong growth (deflation risk)
        target: 2.0,     // Central bank target
        tolerance: 0.5,  // +/- tolerance around target
        high: 4.0        // Above this = defensive signal
      },
      vix: {
        low: 15,         // Below this = complacency
        normal: 20,      // Normal range
        elevated: 25,    // Elevated concern
        high: 30         // High fear
      }
    },
    weights: {
      unemployment: 1.5,   // Higher weight for employment
      usCPI: 1.2,
      canCPI: 0.8,
      canGDP: 1.0,
      vix: 1.3,           // Higher weight for volatility
      yieldCurve: 1.4,    // Higher weight for yield curve
      creditSpread: 1.1
    },
    trendPeriods: 3,      // Number of periods to look back for trend
    showDetails: false    // Whether to show detailed breakdown
  };
  
  // Try to read config from sheet, fall back to defaults
  try {
    var configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      // Read custom configuration from sheet
      // Implementation would read key-value pairs and override defaults
    }
  } catch (e) {
    // Config sheet doesn't exist, use defaults
  }
  
  return defaultConfig;
}

// function getHistoricalData(ss) {
//   try {
//     var histSheet = ss.getSheetByName("Historical");
//     if (!histSheet) return null;
    
//     // Read last 6 months of data for trend analysis
//     var data = histSheet.getRange("A2:H7").getValues(); // Adjust range as needed
    
//     return {
//       unemployment: data.map(function(row) { return parseFloat(row[1]); }).filter(function(x) { return !isNaN(x); }),
//       usCPI: data.map(function(row) { return parseFloat(row[2]); }).filter(function(x) { return !isNaN(x); }),
//       canCPI: data.map(function(row) { return parseFloat(row[3]); }).filter(function(x) { return !isNaN(x); }),
//       vix: data.map(function(row) { return parseFloat(row[4]); }).filter(function(x) { return !isNaN(x); }),
//       // Add other indicators as needed
//     };
//   } catch (e) {
//     return null;
//   }
// }

function calculateUnemploymentScore(current, historical, config) {
  if (current === null) return { score: 0, description: "No data" };
  
  var baseScore = 0;
  var description = current + "%";
  
  // Base score from thresholds
  if (current < config.thresholds.unemployment.veryLow) {
    baseScore = 2;
    description += " (Very Low)";
  } else if (current < config.thresholds.unemployment.low) {
    baseScore = 1;
    description += " (Low)";
  } else if (current > config.thresholds.unemployment.veryHigh) {
    baseScore = -2;
    description += " (Very High)";
  } else if (current > config.thresholds.unemployment.high) {
    baseScore = -1;
    description += " (High)";
  } else {
    baseScore = 0;
    description += " (Normal)";
  }
  
  // Trend adjustment
  var trendAdjustment = calculateTrend(historical, config.trendPeriods);
  if (trendAdjustment !== 0) {
    // If unemployment is falling, that's good for growth
    // If unemployment is rising, that's bad for growth
    baseScore -= trendAdjustment * 0.5; // Trend adjustment factor
    description += (trendAdjustment > 0 ? " Rising" : " Falling");
  }
  
  return { score: Math.max(-2, Math.min(2, baseScore)), description: description };
}

function calculateCPIScore(current, historical, config, region) {
  if (current === null) return { score: 0, description: "No data" };
  
  var target = config.thresholds.cpi.target;
  var tolerance = config.thresholds.cpi.tolerance;
  var baseScore = 0;
  var description = current + "%";
  
  // Score based on distance from target
  if (current < config.thresholds.cpi.veryLow) {
    baseScore = -1; // Deflation risk
    description += " (Deflation Risk)";
  } else if (Math.abs(current - target) <= tolerance) {
    baseScore = 1; // Near target is good
    description += " (On Target)";
  } else if (current > config.thresholds.cpi.high) {
    baseScore = -2; // High inflation is bad
    description += " (High Inflation)";
  } else if (current > target + tolerance) {
    baseScore = -1; // Above target
    description += " (Above Target)";
  } else {
    baseScore = 0; // Between deflation risk and target
    description += " (Below Target)";
  }
  
  // Trend adjustment
  var trendAdjustment = calculateTrend(historical, config.trendPeriods);
  if (trendAdjustment !== 0) {
    // Rising inflation is generally bad, falling inflation is good (unless near deflation)
    if (current > config.thresholds.cpi.veryLow) {
      baseScore -= trendAdjustment * 0.3;
    }
    description += (trendAdjustment > 0 ? " Rising" : " Falling");
  }
  
  return { score: Math.max(-2, Math.min(2, baseScore)), description: description };
}

function calculateVIXScore(current, historical, config) {
  if (current === null) return { score: 0, description: "No data" };
  
  var baseScore = 0;
  var description = current.toFixed(1);
  
  if (current < config.thresholds.vix.low) {
    baseScore = 1; // Low volatility is good for growth
    description += " (Low Vol)";
  } else if (current < config.thresholds.vix.normal) {
    baseScore = 0.5;
    description += " (Normal)";
  } else if (current < config.thresholds.vix.elevated) {
    baseScore = -0.5;
    description += " (Elevated)";
  } else if (current < config.thresholds.vix.high) {
    baseScore = -1;
    description += " (High)";
  } else {
    baseScore = -2; // Very high volatility
    description += " (Very High)";
  }
  
  // Trend adjustment - rising VIX is bad
  var trendAdjustment = calculateTrend(historical, config.trendPeriods);
  if (trendAdjustment !== 0) {
    baseScore -= trendAdjustment * 0.4;
    description += (trendAdjustment > 0 ? " Rising" : " Falling");
  }
  
  return { score: Math.max(-2, Math.min(2, baseScore)), description: description };
}

function calculateGDPScore(current, historical, config) {
  if (!current) return { score: 0, description: "No data" };
  
  var score = 0;
  var description = current;
  
  // Simple text matching with more flexibility
  var growthTerms = ["rising", "growing", "increasing", "up", "positive", "expanding"];
  var declineTerms = ["falling", "declining", "decreasing", "down", "negative", "contracting"];
  
  var currentLower = current.toLowerCase();
  
  if (growthTerms.some(function(term) { return currentLower.includes(term); })) {
    score = 1;
  } else if (declineTerms.some(function(term) { return currentLower.includes(term); })) {
    score = -1;
  }
  
  return { score: score, description: description };
}

function calculateYieldScore(current, historical, config) {
  if (current === null) return { score: 0, description: "No data" };
  
  var baseScore = 0;
  var description = current.toFixed(2) + "%";
  
  if (current > 0.5) {
    baseScore = 1; // Positive yield curve is good
    description += " (Positive)";
  } else if (current > -0.5) {
    baseScore = 0; // Flat curve
    description += " (Flat)";
  } else if (current > -1.0) {
    baseScore = -1; // Mildly inverted
    description += " (Inverted)";
  } else {
    baseScore = -2; // Deeply inverted
    description += " (Deep Inversion)";
  }
  
  // Trend adjustment
  var trendAdjustment = calculateTrend(historical, config.trendPeriods);
  if (trendAdjustment !== 0) {
    baseScore += trendAdjustment * 0.3; // Steepening curve is generally good
    description += (trendAdjustment > 0 ? " Steepening" : " Flattening");
  }
  
  return { score: Math.max(-2, Math.min(2, baseScore)), description: description };
}

function calculateCreditScore(current, historical, config) {
  if (current === null) return { score: 0, description: "No data" };
  
  var baseScore = 0;
  var description = current.toFixed(2) + "%";
  
  if (current < 1.0) {
    baseScore = 1; // Low credit spreads are good
    description += " (Low)";
  } else if (current < 2.0) {
    baseScore = 0; // Normal spreads
    description += " (Normal)";
  } else if (current < 3.0) {
    baseScore = -1; // Elevated spreads
    description += " (Elevated)";
  } else {
    baseScore = -2; // High spreads indicate stress
    description += " (High)";
  }
  
  // Trend adjustment - widening spreads are bad
  var trendAdjustment = calculateTrend(historical, config.trendPeriods);
  if (trendAdjustment !== 0) {
    baseScore -= trendAdjustment * 0.4;
    description += (trendAdjustment > 0 ? " Widening" : " Tightening");
  }
  
  return { score: Math.max(-2, Math.min(2, baseScore)), description: description };
}

function calculateTrend(historicalData, periods) {
  if (!historicalData || historicalData.length < 2) return 0;
  
  // Simple trend calculation: compare recent average to older average
  var recentPeriods = Math.min(periods, historicalData.length);
  var recent = historicalData.slice(-recentPeriods);
  var older = historicalData.slice(0, -recentPeriods);
  
  if (older.length === 0) return 0;
  
  var recentAvg = recent.reduce(function(a, b) { return a + b; }, 0) / recent.length;
  var olderAvg = older.reduce(function(a, b) { return a + b; }, 0) / older.length;
  
  var percentChange = (recentAvg - olderAvg) / olderAvg;
  
  // Return normalized trend (-1 to 1)
  return Math.max(-1, Math.min(1, percentChange * 10));
}

function determineSignal(score, config) {
  var roundedScore = Math.round(score * 2) / 2; // Round to nearest 0.5
  
  if (score >= 1.5) return "Strong Growth (" + roundedScore.toFixed(1) + ")";
  if (score >= 0.5) return "Weak Growth (" + roundedScore.toFixed(1) + ")";
  if (score >= -0.5) return "Neutral (" + roundedScore.toFixed(1) + ")";
  if (score >= -1.5) return "Weak Defensive (" + roundedScore.toFixed(1) + ")";
  return "Strong Defensive (" + roundedScore.toFixed(1) + ")";
}

/**
 * Enhanced GET_REBALANCE_SIGNAL(assetTicker)
 * Uses the enhanced market scoring system with weights, trends, and dynamic thresholds
 * Provides more nuanced rebalancing signals based on weighted market indicators
 */
function GET_REBALANCE_SIGNAL(assetTicker) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1) Load current allocations from A & B
  var allocSh   = ss.getSheetByName("Asset Allocations");
  var lastRow   = allocSh.getLastRow();
  var allocData = allocSh.getRange("A2:B" + lastRow).getValues();
  var allocMap = {}, tickers = [];
  allocData.forEach(function(r){
    var t = r[0], raw = r[1], a;
    if (typeof raw === "string" && raw.trim().endsWith("%")) {
      a = parseFloat(raw) / 100;
    } else {
      a = parseFloat(raw);
    }
    if (t && !isNaN(a)) {
      allocMap[t] = a;
      tickers.push(t);
    }
  });
  
  // 2) Get enhanced market signal data
  var marketData = getEnhancedMarketData(ss);
  
  // 3) Load Asset Metadata (A=Asset key, B=Class, C=Region, D=Sensitivity)
  var metaSh = ss.getSheetByName("Asset Metadata");
  var meta   = metaSh.getRange("A2:D" + metaSh.getLastRow()).getValues();
  var classMap = {}, regionMap = {}, sensitivityMap = {};
  meta.forEach(function(r){
    var asset = r[0];
    classMap[asset] = r[1];      // "Equity" or "Defensive"
    regionMap[asset] = r[2];     // "U.S.", "Canada", etc.
    sensitivityMap[asset] = parseFloat(r[3]) || 1.0; // Sensitivity multiplier (default 1.0)
  });
  
  // Fail fast on typos
  if (!(assetTicker in classMap)) {
    throw new Error("No metadata for " + assetTicker + " – check Asset Metadata");
  }
  
  // 4) Enhanced region-specific scoring function
  function getRegionalScore(ticker) {
    var region = regionMap[ticker] || "Global";
    var scores = marketData.scores;
    var config = marketData.config;
    
    var totalScore = 0;
    var totalWeight = 0;
    
    if (region === "U.S.") {
      // US-focused weighting
      totalScore += scores.unemployment * config.weights.unemployment * 1.5; // Higher US weight
      totalWeight += config.weights.unemployment * 1.5;
      
      totalScore += scores.usCPI * config.weights.usCPI * 1.3;
      totalWeight += config.weights.usCPI * 1.3;
      
      totalScore += scores.vix * config.weights.vix * 1.2;
      totalWeight += config.weights.vix * 1.2;
      
      totalScore += scores.yieldCurve * config.weights.yieldCurve * 1.1;
      totalWeight += config.weights.yieldCurve * 1.1;
      
      totalScore += scores.creditSpread * config.weights.creditSpread * 1.1;
      totalWeight += config.weights.creditSpread * 1.1;
      
    } else if (region === "Canada") {
      // Canada-focused weighting
      totalScore += scores.canCPI * config.weights.canCPI * 1.5;
      totalWeight += config.weights.canCPI * 1.5;
      
      totalScore += scores.canGDP * config.weights.canGDP * 1.4;
      totalWeight += config.weights.canGDP * 1.4;
      
      // Still include some US indicators but with lower weight
      totalScore += scores.unemployment * config.weights.unemployment * 0.7;
      totalWeight += config.weights.unemployment * 0.7;
      
      totalScore += scores.vix * config.weights.vix * 0.8;
      totalWeight += config.weights.vix * 0.8;
      
    } else {
      // Global composite - use all indicators with standard weights
      Object.keys(scores).forEach(function(key) {
        if (scores[key] !== null && config.weights[key]) {
          totalScore += scores[key] * config.weights[key];
          totalWeight += config.weights[key];
        }
      });
    }
    
    return totalWeight > 0 ? totalScore / totalWeight : 0;
  }
  
  // 5) Enhanced delta calculation with volatility adjustment
  function calculateEnhancedDelta(ticker) {
    var regionalScore = getRegionalScore(ticker);
    var sensitivity = sensitivityMap[ticker] || 1.0;
    var assetClass = classMap[ticker];
    
    // Base shift calculation with more granular levels
    var baseShift = 0;
    if (regionalScore >= 1.5) {
      baseShift = 0.08;      // Strong growth signal
    } else if (regionalScore >= 0.75) {
      baseShift = 0.05;      // Moderate growth signal
    } else if (regionalScore >= 0.25) {
      baseShift = 0.025;     // Weak growth signal
    } else if (regionalScore >= -0.25) {
      baseShift = 0;         // Neutral
    } else if (regionalScore >= -0.75) {
      baseShift = -0.025;    // Weak defensive signal
    } else if (regionalScore >= -1.5) {
      baseShift = -0.05;     // Moderate defensive signal
    } else {
      baseShift = -0.08;     // Strong defensive signal
    }
    
    // Apply sensitivity multiplier
    baseShift *= sensitivity;
    
    // Volatility adjustment - reduce shifts during high volatility periods
    var vixScore = marketData.scores.vix || 0;
    if (vixScore < -1) { // High VIX
      baseShift *= 0.7; // Reduce rebalancing during high volatility
    }
    
    // Class-specific adjustments
    if (assetClass === "Defensive") {
      return regionalScore > 0 ? -baseShift : baseShift;
    } else { // Equity
      return regionalScore > 0 ? baseShift : -baseShift;
    }
  }
  
  // 6) Calculate raw deltas with enhanced logic
  var rawDelta = {};
  var defList = tickers.filter(t => classMap[t] === "Defensive");
  var eqList = tickers.filter(t => classMap[t] === "Equity");
  var nd = defList.length;
  var ne = eqList.length;
  
  // Calculate individual deltas
  tickers.forEach(function(t) {
    rawDelta[t] = calculateEnhancedDelta(t);
  });
  
  // 7) Portfolio balance constraint - ensure defensive and equity moves offset
  var totalDefMove = 0;
  var totalEqMove = 0;
  var totalDefAlloc = 0;
  var totalEqAlloc = 0;
  
  defList.forEach(function(t) {
    totalDefMove += rawDelta[t] * allocMap[t];
    totalDefAlloc += allocMap[t];
  });
  
  eqList.forEach(function(t) {
    totalEqMove += rawDelta[t] * allocMap[t];
    totalEqAlloc += allocMap[t];
  });
  
  // Apply balancing adjustment to ensure net moves offset
  var netMove = totalDefMove + totalEqMove;
  if (Math.abs(netMove) > 0.001) { // If significant imbalance
    var adjustment = -netMove / 2; // Split the adjustment
    
    if (totalDefAlloc > 0) {
      defList.forEach(function(t) {
        rawDelta[t] += adjustment / totalDefAlloc;
      });
    }
    
    if (totalEqAlloc > 0) {
      eqList.forEach(function(t) {
        rawDelta[t] += adjustment / totalEqAlloc;
      });
    }
  }
  
  // 8) Build raw new allocations and normalize
  var sumRaw = 0;
  var rawAlloc = {};
  tickers.forEach(function(t) {
    var oldA = allocMap[t];
    rawAlloc[t] = Math.max(0, oldA * (1 + rawDelta[t])); // Prevent negative allocations
    sumRaw += rawAlloc[t];
  });
  
  // 9) Calculate final signal for the requested ticker
  var oldA = allocMap[assetTicker];
  var normalized = rawAlloc[assetTicker] / sumRaw;
  var finalD = (normalized / oldA) - 1;
  
  // 10) Enhanced signal formatting with confidence indication
  var regionalScore = getRegionalScore(assetTicker);
  var confidence = Math.abs(regionalScore);
  var confidenceLevel = "";
  
  if (confidence >= 1.5) {
    confidenceLevel = " (High Confidence)";
  } else if (confidence >= 0.75) {
    confidenceLevel = " (Medium Confidence)";
  } else if (confidence >= 0.25) {
    confidenceLevel = " (Low Confidence)";
  }
  
  // Add minimum threshold to avoid tiny changes
  var minThreshold = 0.005; // 0.5% minimum change
  
  if (Math.abs(finalD) < minThreshold) {
    return "Hold";
  } else if (finalD > 0) {
    return "Increase " + (finalD * 100).toFixed(2) + "%";
  } else {
    return "Decrease " + (Math.abs(finalD) * 100).toFixed(2) + "%";
  }
}

/**
 * Helper function to get enhanced market data using the new scoring system
 */
function getEnhancedMarketData(ss) {
  // Get indicator data
  var indicatorData = getIndicatorData(ss);
  
  // Get configuration
  var config = getRebalanceConfig(ss);
  
  // Get historical data for trends
  var historical = getHistoricalDataFromFRED();
  
  // Calculate individual indicator scores using the enhanced functions
  var scores = {
    unemployment: calculateUnemploymentScore(indicatorData.unemployment, historical ? historical.unemployment : null, config).score,
    usCPI: calculateCPIScore(indicatorData.usCPI, historical ? historical.usCPI : null, config, 'US').score,
    canCPI: calculateCPIScore(indicatorData.canCPI, historical ? historical.canCPI : null, config, 'CAN').score,
    canGDP: calculateGDPScore(indicatorData.canGDP, historical ? historical.canGDP : null, config).score,
    vix: calculateVIXScore(indicatorData.vix, historical ? historical.vix : null, config).score,
    yieldCurve: calculateYieldScore(indicatorData.yieldCurve, historical ? historical.yieldCurve : null, config).score,
    creditSpread: calculateCreditScore(indicatorData.creditSpread, historical ? historical.creditSpread : null, config).score
  };
  
  return {
    scores: scores,
    config: config,
    indicators: indicatorData
  };
}

/**
 * Get rebalancing-specific configuration
 */
function getRebalanceConfig(ss) {
  // Start with the same config as market signals
  var config = getConfigData(ss);
  
  // Add rebalancing-specific parameters
  config.rebalancing = {
    minThreshold: 0.005,    // Minimum 0.5% change to trigger signal
    maxSingleMove: 0.15,    // Maximum 15% single asset move
    volatilityAdjustment: true, // Reduce moves during high volatility
    balanceConstraint: true     // Ensure defensive/equity moves offset
  };
  
  return config;
}
