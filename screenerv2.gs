// =============================================================================
// SCREENER.IN MULTI-MTF TRACKER — Google Apps Script v4.0
// =============================================================================
// BUGS FIXED FROM v3:
//  [CRITICAL] Stocks with no ticker column were ALWAYS skipped (tracked nothing
//             when Screener had no explicit "Ticker" column — the default case).
//             Fixed: allow name-only stocks; symbol is resolved via Yahoo search
//             at ADD time so it's available for all future runs.
//  [CRITICAL] applyRowFormatting made N individual getRange() calls in a loop
//             (one per stock for signal color, one per stock for highlight clear).
//             Fixed: batch all background + font color writes with setBackgrounds()
//             / setFontColors() on a single range — 2 calls instead of N×2.
//  [CRITICAL] updateDashboard made per-stock individual API calls (300+ calls for
//             100 stocks across 3 scanners). Fixed: batch all cell formatting.
//  [HIGH]     GAS 6-min timeout guard was local to each scanner's update loop.
//             With 3 scanners the guard reset each time; GAS could still kill the
//             script mid-run. Fixed: global RUN_START_MS passed through all calls.
//  [HIGH]     perfectBull fired for stocks with < 200 days of data because
//             ma200BelowOthers defaulted to `true` when MA200 was null.
//             Fixed: perfectBull now REQUIRES hasMa200; new stocks cap at "✅ BUY".
//  [MEDIUM]   Unresolvable symbols (delisted/renamed) triggered a Yahoo search
//             on every 15-min run forever. Fixed: write "~NOFOUND" sentinel so
//             future runs skip the search immediately.
//  [MEDIUM]   loadExistingStocks keyed by sym||name but processScanner lookup
//             only used stock.symbol — name-only stocks never matched. Fixed:
//             lookup now tries symbol first, then name as fallback.
//  [LOW]      Return columns had no number format set in the sheet — values stored
//             as raw floats. Fixed: writeHeaders() now sets "0.00" on return cols.
//  [LOW]      `window` parameter shadowed a common global. Renamed to `lookback`.
//  [LOW]      log() called getActiveSpreadsheet() on every invocation (API roundtrip).
//             Fixed: runAllScanners passes `ss` into all functions; log() accepts it.
// =============================================================================

// ─────────────────────────────────────────────────────────────────────────────
// CONFIG — edit only this section
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG = {

  // One entry per Screener scanner. Each gets its own sheet tab.
  // URL: screener.in/screens/XXXX → append ?format=json
  // TIP: Add a "Ticker" column in your Screener scanner for the most reliable
  //      symbol extraction. Without it the script auto-searches Yahoo by name,
  //      which still works but adds ~1 sec per new stock.
  SCANNERS: [
  { name: "GOAT 1", url: "https://www.screener.in/screens/3525076/goat1/?format=json", color: "#E8F5E9" },              // Soft Green
  { name: "Weekly Burst", url: "https://www.screener.in/screens/3530492/weekly-burst/?format=json", color: "#E3F2FD" },  // Light Blue
  { name: "Month Returns", url: "https://www.screener.in/screens/3530540/month-returns/?format=json", color: "#FFF8E1" },// Soft Yellow
  { name: "Weekly Screener - Maximum Technical Precision", url: "https://www.screener.in/screens/3530554/weekly-screener-maximum-technical-precision/?format=json", color: "#F3E5F5" }, // Lavender
  { name: "Monthly Opus V3", url: "https://www.screener.in/screens/3534290/monthly-opus-v3/?format=json", color: "#E0F7FA" }, // Cyan
  { name: "Opus Tele", url: "https://www.screener.in/screens/3535927/opus-tele/?format=json", color: "#FCE4EC" },        // Pink
  { name: "Weekly", url: "https://www.screener.in/screens/3536956/weekly/?format=json", color: "#E8EAF6" },              // Indigo Light
  { name: "Daily Scanner", url: "https://www.screener.in/screens/3536981/daily-scanner/?format=json", color: "#E1F5FE" },// Sky Blue
  { name: "Daily Must Buy", url: "https://www.screener.in/screens/3537642/daily-must-buy/?format=json", color: "#E8F5E9" }, // Mint
  { name: "Weekly Must Buy", url: "https://www.screener.in/screens/3537648/weekly-must-buy/?format=json", color: "#FFF3E0" }, // Light Orange
  { name: "Monthly Must Buy", url: "https://www.screener.in/screens/3537651/monthly-must-buy/?format=json", color: "#F1F8E9" }, // Lime
  { name: "Daily Scanner V2 - Less Noise", url: "https://www.screener.in/screens/3539839/daily-scanner-v2-less-noise/?format=json", color: "#E0F2F1" }, // Teal
  { name: "Long Term", url: "https://www.screener.in/screens/3540686/long-term/?format=json", color: "#FBE9E7" },        // Peach
],

  // For private/logged-in scanners: paste your Screener sessionid cookie here.
  // How: screener.in → F12 → Application → Cookies → copy "sessionid" value.
  SCREENER_COOKIE: "sessionid=15mbhoovxdh6a91pjd3gxt6cu93u8yr0",

  // ── CLOUDFLARE WORKER PROXY ───────────────────────────────────────────────
  // Screener.in returns HTTP 404 for Google Apps Script requests because GAS
  // runs from Google datacenter IPs that Screener.in blocks.
  // Your existing Cloudflare Worker (github.com/ajithvnr2001/screener-alerts-multi-mtf)
  // runs from edge IPs that Screener.in allows — route through it.
  //
  // Add this tiny route to your existing Worker (or deploy a new one):
  // ─────────────────────────────────────────────────────────────────
  //   export default {
  //     async fetch(request) {
  //       const target = new URL(request.url).searchParams.get("url");
  //       if (!target) return new Response("missing ?url=", { status: 400 });
  //       return fetch(decodeURIComponent(target), {
  //         headers: { "User-Agent": "Mozilla/5.0", "Accept": "application/json",
  //                    "X-Requested-With": "XMLHttpRequest" }
  //       });
  //     }
  //   };
  // ─────────────────────────────────────────────────────────────────
  // Then paste your Worker URL below. Leave "" to attempt direct (will 404).
  PROXY_URL: "https://screener-proxy.ltimindtree.workers.dev",

  // Email alert when new stocks appear in any scanner. Set "" to disable.
  ALERT_EMAIL: "your@email.com",

  // Yahoo Finance exchange suffix for your market:
  //   NSE stocks → ".NS"    BSE stocks → ".BO"
  YF_SUFFIX: ".NS",

  // Market hours in IST — the script schedules only the next valid slot inside
  // the allowed date filter + this intraday time window.
  MARKET_START_HOUR:   9,
  MARKET_START_MINUTE: 0,   // 9:00 AM IST
  MARKET_END_HOUR:    15,
  MARKET_END_MINUTE:  40,   // 3:40 PM IST

  // Optional explicit calendar date windows in IST. Each range is inclusive.
  // Leave [] to use the default weekday mode (Mon-Fri).
  // Add multiple combinations as needed.
  ACTIVE_DATE_RANGES: [
    // { from: "2026-03-16", to: "2026-03-20", label: "March 3rd" },
    // { from: "2026-03-23", to: "2026-03-27", label: "March 4th" },
  ],

  // Trigger cadence for the self-scheduling next-run trigger.
  TRIGGER_INTERVAL_MINUTES: 15,

  // Signal tuning preset: "conservative" | "balanced" | "aggressive"
  SIGNAL_PROFILE: "balanced",

  // System sheet names (do not rename these once created)
  DASHBOARD_SHEET:     "📊 Dashboard",
  PRICE_HISTORY_SHEET: "📈 Price History",
  CHART_DATA_SHEET:    "📉 Chart Data",
  LOG_SHEET:           "🗒️ Log",

  MAX_LOG_ROWS: 400,

  // Global execution time budget (milliseconds). GAS hard-kills at 6 min (360 000 ms).
  // We stop fetching at 5 min to leave time for the final bulk write + formatting.
  MAX_RUN_MS: 5 * 60 * 1000,
};

const SIGNAL_PROFILES = {
  conservative: {
    ADX_STRONG: 25,
    ADX_WEAK: 18,
    VOL_RATIO_HIGH: 1.5,
    DIST_52W_HIGH_MAX: 2,
    BREAKOUT_20D_MIN: 0.75,
    RSI_OVERSOLD: 30,
    RSI_NEUTRAL_MIN: 45,
    RSI_BULLISH_MIN: 55,
    RSI_OVERBOUGHT: 78,
    MACD_ZERO_TOL_PCT: 0.00005,
  },
  balanced: {
    ADX_STRONG: 20,
    ADX_WEAK: 16,
    VOL_RATIO_HIGH: 1.25,
    DIST_52W_HIGH_MAX: 3,
    BREAKOUT_20D_MIN: 0.25,
    RSI_OVERSOLD: 32,
    RSI_NEUTRAL_MIN: 45,
    RSI_BULLISH_MIN: 52,
    RSI_OVERBOUGHT: 78,
    MACD_ZERO_TOL_PCT: 0.00005,
  },
  aggressive: {
    ADX_STRONG: 18,
    ADX_WEAK: 14,
    VOL_RATIO_HIGH: 1.1,
    DIST_52W_HIGH_MAX: 5,
    BREAKOUT_20D_MIN: 0,
    RSI_OVERSOLD: 35,
    RSI_NEUTRAL_MIN: 42,
    RSI_BULLISH_MIN: 50,
    RSI_OVERBOUGHT: 80,
    MACD_ZERO_TOL_PCT: 0.00005,
  },
};

const SIGNAL_PROFILE_KEY = "signalProfile";
const AUTO_BACKFILL_STATUS_KEY = "priceHistoryAutoBackfillStatus";
const AUTO_BACKFILL_PENDING = "pending";
const AUTO_BACKFILL_COMPLETE = "complete";
const AUTO_BACKFILL_TRIGGER_DELAY_MS = 2 * 60 * 1000;

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN INDEX MAP  (0-based; Sheet column A = index 0)
// ─────────────────────────────────────────────────────────────────────────────
const C = {
  SYMBOL:         0,   // NSE ticker (resolved via Yahoo if not in Screener export)
  NAME:           1,   // Company name from Screener
  FIRST_CAPTURED: 2,   // Datetime first seen in the scanner
  LAST_SEEN:      3,   // Datetime last seen in the scanner
  IN_SCREENER:    4,   // "✅" = currently in scan | "⬜" = dropped off
  CAPTURE_PRICE:  5,   // Closing price on first capture date (set once, never overwritten)
  CURRENT_PRICE:  6,   // Latest closing price from Yahoo Finance
  RET_CAPTURE:    7,   // % return since CAPTURE_PRICE
  RET_1D:         8,   // 1-day return
  RET_1W:         9,   // 5-day return
  RET_1M:         10,  // 21-day return
  RET_3M:         11,  // 63-day return
  RET_6M:         12,  // 126-day return
  RET_1Y:         13,  // 252-day return
  RET_2Y:         14,  // 504-day return
  RET_3Y:         15,  // 756-day return
  AVG_WEEKLY:     16,  // Avg of all non-overlapping 5-day returns (last 1yr)
  AVG_MONTHLY:    17,  // Avg of all non-overlapping 21-day returns (last 2yr)
  AVG_3M:         18,  // Avg of all non-overlapping 63-day returns (full 3yr)
  AVG_6M:         19,  // Avg of all non-overlapping 126-day returns (full 3yr)
  AVG_1Y:         20,  // Avg of all non-overlapping 252-day returns (full 3yr)
  RSI14:          21,  // RSI with Wilder smoothing, 14-period
  MA20:           22,  // 20-day simple moving average
  MA50:           23,  // 50-day simple moving average
  MA200:          24,  // 200-day simple moving average
  SIGNAL:         25,  // Buy / Hold / Sell signal string
  LAST_UPDATED:   26,  // Datetime of last Yahoo Finance update
  ADX14:          27,  // ADX with Wilder smoothing, 14-period
  VOL_RATIO20:    28,  // Current volume / previous 20-day average volume
  MACD_LINE:      29,  // MACD line (12,26)
  MACD_HIST:      30,  // MACD histogram (12,26,9)
  DIST_52W_HIGH:  31,  // % distance from 52-week high (0 = at high)
  BREAKOUT_20D:   32,  // % above/below previous 20-day high (positive = breakout)
  _COUNT:         33,
};

const HEADERS = [
  "Symbol", "Name", "First Captured", "Last Seen", "In Screener?",
  "Capture Price ₹", "Current Price ₹", "Since Capture %",
  "1D %", "1W %", "1M %", "3M %", "6M %", "1Y %", "2Y %", "3Y %",
  "Avg Weekly %", "Avg Monthly %", "Avg 3M %", "Avg 6M %", "Avg 1Y %",
  "RSI (14)", "MA 20", "MA 50", "MA 200",
  "Signal", "Last Updated", "ADX (14)", "Vol Ratio (20)", "MACD Line", "MACD Hist", "52W High Dist %", "20D Breakout %",
];

const H = {
  SNAPSHOT_AT:    0,
  RUN_MODE:       1,
  SCANNER:        2,
  STOCK_KEY:      3,
  SYMBOL:         4,
  NAME:           5,
  IN_SCREENER:    6,
  CAPTURE_PRICE:  7,
  CURRENT_PRICE:  8,
  RET_CAPTURE:    9,
  RET_1D:         10,
  RET_1W:         11,
  RET_1M:         12,
  RET_3M:         13,
  RET_6M:         14,
  RET_1Y:         15,
  RSI14:          16,
  ADX14:          17,
  VOL_RATIO20:    18,
  MACD_LINE:      19,
  MACD_HIST:      20,
  DIST_52W_HIGH:  21,
  BREAKOUT_20D:   22,
  SIGNAL:         23,
  _COUNT:         24,
};

const PRICE_HISTORY_HEADERS = [
  "Snapshot At", "Run Mode", "Scanner", "Stock Key", "Symbol", "Name",
  "In Screener?", "Capture Price ₹", "Current Price ₹", "Since Capture %",
  "1D %", "1W %", "1M %", "3M %", "6M %", "1Y %",
  "RSI (14)", "ADX (14)", "Vol Ratio (20)", "MACD Line", "MACD Hist",
  "52W High Dist %", "20D Breakout %", "Signal",
];

const DASH_UI = {
  TOTAL_COLS: 14,
  SCANNER_CELL: "B6",
  SYMBOL_CELL: "B7",
  RUNS_CELL: "B8",
  LATEST_PRICE_CELL: "H6",
  RETURN_CELL: "H7",
  SIGNAL_CELL: "H8",
  FIRST_RUN_CELL: "K6",
  LAST_RUN_CELL: "K7",
  TABLE_START_ROW: 40,
  PRICE_CHART_ROW: 11,
  PRICE_CHART_COL: 1,
  RETURN_CHART_ROW: 25,
  RETURN_CHART_COL: 1,
  MULTI_RETURN_CHART_ROW: 25,
  MULTI_RETURN_CHART_COL: 8,
};

const DASHBOARD_COL_WIDTHS = [110, 280, 95, 95, 170, 140, 100, 120, 90, 110, 80, 80, 80, 80];

// Return column indices (used for number format + conditional formatting)
const RETURN_COLS = [
  C.RET_CAPTURE, C.RET_1D, C.RET_1W, C.RET_1M,
  C.RET_3M, C.RET_6M, C.RET_1Y, C.RET_2Y, C.RET_3Y,
  C.AVG_WEEKLY, C.AVG_MONTHLY, C.AVG_3M, C.AVG_6M, C.AVG_1Y,
  C.BREAKOUT_20D,
];

const PRICE_COLS = [C.CAPTURE_PRICE, C.CURRENT_PRICE, C.MA20, C.MA50, C.MA200];
const SINGLE_DECIMAL_COLS = [C.RSI14, C.ADX14];
const DOUBLE_DECIMAL_COLS = [C.VOL_RATIO20, C.MACD_LINE, C.MACD_HIST, C.DIST_52W_HIGH, C.BREAKOUT_20D];

// Sentinel written to SYMBOL when Yahoo symbol-search fails, to skip re-searching
const NO_SYMBOL_SENTINEL = "~NOFOUND";

// ═════════════════════════════════════════════════════════════════════════════
// MAIN ENTRY — called by installable time trigger
// ═════════════════════════════════════════════════════════════════════════════
function runAllScanners() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSystemSheets(ss);
  try {
    const schedule = getRunScheduleStatus();
    if (!schedule.isLiveWindow) {
      log(ss, "⏭ Skipped — " + schedule.reason);
      return;
    }

    // Global run timer — shared across ALL scanners and update loops
    const RUN_START = Date.now();
    const runContext = buildRunContext("AUTO");
    const allNewStocks = [];

    CONFIG.SCANNERS.forEach(scanner => {
      if (isTimedOut(RUN_START)) {
        log(ss, `⏱ Global timeout reached — skipping remaining scanners`);
        return;
      }
      try {
        log(ss, `🔍 Running: ${scanner.name}`);
        const newOnes = processScanner(ss, scanner, RUN_START, runContext);
        if (newOnes.length > 0) allNewStocks.push({ scanner: scanner.name, stocks: newOnes });
      } catch (e) {
        log(ss, `❌ [${scanner.name}] ${e.message}\n${e.stack || ""}`);
      }
    });

    updateDashboard(ss);
    maybeScheduleAutomaticBackfill(ss);

    if (CONFIG.ALERT_EMAIL && allNewStocks.length > 0) {
      sendEmailAlert(ss, allNewStocks);
    }

    const total = allNewStocks.reduce((a, b) => a + b.stocks.length, 0);
    log(ss, `✅ Run complete — ${total} new stock(s) captured — ${msToSec(Date.now() - RUN_START)}s elapsed`);
  } finally {
    ensureNextRunTrigger(ss);
  }
}

// ═════════════════════════════════════════════════════════════════════════════
// PROCESS ONE SCANNER
// ═════════════════════════════════════════════════════════════════════════════
function processScanner(ss, scanner, RUN_START, runContext) {
  const sheet = getOrCreateScannerSheet(ss, scanner);

  // ── 1. Fetch current screener results ──────────────────────────────────────
  const screenerStocks = fetchScreenerScan(scanner.url);
  if (screenerStocks.length === 0) {
    log(ss, `  ↳ 0 stocks — screen is empty or returned no results today`);
    return [];
  }
  log(ss, `  ↳ ${screenerStocks.length} stocks currently in screener`);

  // ── 2. Load existing tracked stocks (1 bulk read) ─────────────────────────
  const lastRow    = sheet.getLastRow();
  const existingMap = loadExistingStocks(sheet, lastRow);
  // Map: key (symbol OR name) → { rowIdx (1-based), symbol, name }

  // ── 3. Mark ALL existing stocks as "not in screener" (1 bulk write) ───────
  batchMarkNotInScreener(sheet, lastRow);

  // ── 4. Add new stocks / mark existing ones present ────────────────────────
  const nowStr        = formatDate(new Date());
  const newStocks     = [];
  const lastSeenUpd   = [];    // { rowIdx, value }
  const inScreenerUpd = [];    // { rowIdx, value }

  for (const stock of screenerStocks) {
    // FIX: do NOT skip stocks with empty symbol — we allow name-only tracking.
    // At least one of symbol or name must be present.
    if (!stock.symbol && !stock.name) continue;

    // Try to find existing record: first by symbol, then by name
    const lookupKey = findExistingKey(existingMap, stock);

    if (lookupKey !== null) {
      // Existing stock — queue updates
      const { rowIdx } = existingMap.get(lookupKey);
      lastSeenUpd.push({ rowIdx, value: nowStr });
      inScreenerUpd.push({ rowIdx, value: "✅" });

      const existing = existingMap.get(lookupKey);
      const storedSym = String(existing.symbol || "").trim();
      const isUnresolved = !storedSym
        || storedSym === NO_SYMBOL_SENTINEL
        || storedSym.startsWith(NO_SYMBOL_SENTINEL + "|");

      if (isUnresolved) {
        if (stock.symbol) {
          // Screener now gave us a real NSE ticker — use it directly
          existing.symbolUpdate = stock.symbol;
          log(ss, `  🔄 Upgraded "${stock.name}" symbol: ${storedSym} → ${stock.symbol}`);
        } else if (stock.bseCode) {
          // Try to resolve now that we have the BSE code
          const resolved = searchYahooSymbol(stock.name, stock.bseCode);
          if (resolved) {
            existing.symbolUpdate = resolved;
            log(ss, `  🔄 BSE resolved "${stock.name}" (BSE:${stock.bseCode}) → ${resolved}`);
          } else {
            // Upgrade bare ~NOFOUND to ~NOFOUND|BSE:CODE so retry logic can use it
            const newSentinel = NO_SYMBOL_SENTINEL + "|BSE:" + stock.bseCode;
            if (storedSym !== newSentinel) {
              existing.symbolUpdate = newSentinel;
              log(ss, `  📌 Stored BSE code for "${stock.name}" → will retry next run`);
            }
          }
        }
      } else if (stock.symbol && storedSym !== stock.symbol) {
        // Already resolved but screener gave a different ticker — keep existing
      }
    } else {
      // Brand-new stock — resolve symbol NOW so all future runs have it
      // FIX: symbol resolution happens at add-time, not at update-time
      let resolvedSymbol = stock.symbol;
      if (!resolvedSymbol) {
        // Pass bseCode (if available from HTML scrape) for better Yahoo lookup
        resolvedSymbol = searchYahooSymbol(stock.name, stock.bseCode || "");
        if (resolvedSymbol) {
          log(ss, `  🔎 Resolved "${stock.name}" → ${resolvedSymbol}`);
        } else {
          // Store BSE code in sentinel so next run can retry with it
          resolvedSymbol = stock.bseCode
            ? NO_SYMBOL_SENTINEL + "|BSE:" + stock.bseCode
            : NO_SYMBOL_SENTINEL;
          log(ss, `  ⚠️ No symbol found for "${stock.name}"${stock.bseCode ? " (BSE:" + stock.bseCode + ")" : ""} — will retry`);
        }
      }

      const newRowIdx = sheet.getLastRow() + 1;
      const emptyRow  = buildEmptyRow(stock, resolvedSymbol, nowStr);
      sheet.getRange(newRowIdx, 1, 1, C._COUNT).setValues([emptyRow]);
      sheet.getRange(newRowIdx, 1, 1, C._COUNT).setBackground("#FFEB3B"); // yellow = new
      existingMap.set(resolvedSymbol || stock.name, {
        rowIdx: newRowIdx,
        symbol: resolvedSymbol,
        name:   stock.name,
      });
      newStocks.push(stock);
      log(ss, `  ➕ New: ${resolvedSymbol || stock.name}`);
    }
  }

  // ── 5. Batch-write last-seen + in-screener + symbol upgrades ───────────────
  batchWriteColumn(sheet, lastSeenUpd,   C.LAST_SEEN   + 1, lastRow);
  batchWriteColumn(sheet, inScreenerUpd, C.IN_SCREENER + 1, lastRow);

  // Flush any symbol upgrades (BSE sentinel → resolved, or bare ~NOFOUND → ~NOFOUND|BSE:CODE)
  const symbolUpd = [];
  for (const entry of existingMap.values()) {
    if (entry.symbolUpdate !== undefined && entry.rowIdx <= lastRow) {
      symbolUpd.push({ rowIdx: entry.rowIdx, value: entry.symbolUpdate });
    }
  }
  if (symbolUpd.length > 0) {
    batchWriteColumn(sheet, symbolUpd, C.SYMBOL + 1, lastRow);
    log(ss, `  💾 Updated ${symbolUpd.length} symbol(s) in sheet`);
  }

  // ── 6. Update performance data for all stocks ─────────────────────────────
  updateSheetPerformance(ss, sheet, RUN_START, scanner.name, runContext);

  return newStocks;
}

// ═════════════════════════════════════════════════════════════════════════════
// FETCH SCREENER.IN SCANNER  (JSON API)
// ═════════════════════════════════════════════════════════════════════════════
function fetchScreenerScan(url) {
  // ── HOW THIS WORKS ────────────────────────────────────────────────────────
  // 1. Try ?format=json via proxy (works for some screens when cookie is set)
  // 2. Fall back to HTML scraping via proxy (works for ALL public screens,
  //    no cookie required — parses the stock table directly from the page HTML)
  //
  // GAS runs from Google datacenter IPs which Screener.in blocks → proxy needed.
  // Set CONFIG.PROXY_URL to route through your Cloudflare Worker.

  const idMatch = url.match(/\/screens\/(\d+)(\/[^?]*)?/i);
  if (!idMatch) throw new Error("Cannot parse screen ID from URL: " + url);
  const screenId = idMatch[1];
  const slug     = (idMatch[2] || "").replace(/\/$/, ""); // e.g. "/below-book-value-stocks"

  const proxyBase = CONFIG.PROXY_URL ? CONFIG.PROXY_URL.replace(/\/$/, "") : "";

  function proxyFetch(target, withCookie) {
    let fetchUrl = target;
    const headers = {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    };
    if (proxyBase) {
      fetchUrl = proxyBase + "?url=" + encodeURIComponent(target);
      headers["User-Agent"] = "Mozilla/5.0 (compatible; GAS/1.0)";
      if (withCookie && CONFIG.SCREENER_COOKIE) {
        headers["X-Screener-Cookie"] = CONFIG.SCREENER_COOKIE;
      }
    } else if (withCookie && CONFIG.SCREENER_COOKIE) {
      headers["Cookie"] = CONFIG.SCREENER_COOKIE;
    }
    return UrlFetchApp.fetch(fetchUrl, {
      method: "GET", headers, muteHttpExceptions: true, followRedirects: true,
    });
  }

  // ── ATTEMPT 1: JSON API with cookie ───────────────────────────────────────
  const jsonCandidates = [
    "https://www.screener.in/screens/" + screenId + slug + "/?format=json",
    "https://www.screener.in/screens/" + screenId + "/?format=json",
    "https://www.screener.in/api/screens/" + screenId + "/?format=json",
  ].filter((v, i, a) => a.indexOf(v) === i);

  for (const target of jsonCandidates) {
    const resp = proxyFetch(target, true);
    if (resp.getResponseCode() !== 200) continue;
    const body = resp.getContentText().trim();
    if (body.startsWith("{")) {
      // Got valid JSON ✅
      const json    = JSON.parse(body);
      const columns = json.columns || [];
      const results = json.results  || [];
      return results.map(function(row) {
        const stock = { symbol: "", name: "" };
        columns.forEach(function(col, i) {
          const colKey = String(col && col.name ? col.name : col || "").toLowerCase().trim();
          if (colKey === "name" || colKey === "company name" || colKey === "company") {
            stock.name = String(row[i] || "").trim();
          }
          const TICKER_COLS = ["ticker","symbol","nse code","nse symbol","bse code","scrip","script","isin code"];
          if (TICKER_COLS.indexOf(colKey) !== -1 && row[i]) {
            const val = String(row[i]).trim().toUpperCase();
            if (!/^\d+$/.test(val) && !/^[A-Z]{2}\d{10}$/.test(val)) stock.symbol = val;
          }
        });
        if (!stock.name && row.length > 1) stock.name = String(row[1] || "").trim();
        return stock;
      }).filter(function(s) { return s.name || s.symbol; });
    }
  }

  // ── ATTEMPT 2: HTML scraping (works for ALL public screens, no cookie needed)
  // The HTML table has: <a href="/company/TICKER/...">Company Name</a>
  // We extract both the NSE ticker (from href) and the name (from link text).
  const htmlUrl = "https://www.screener.in/screens/" + screenId + slug + "/";
  const resp    = proxyFetch(htmlUrl, true);
  const code    = resp.getResponseCode();

  if (code !== 200) {
    throw new Error(
      "Screen " + screenId + ": HTTP " + code + " fetching HTML page.\n" +
      "  → Confirm the screen URL is correct: " + htmlUrl + "\n" +
      (proxyBase ? "" : "  → Set CONFIG.PROXY_URL — GAS IPs are blocked by Screener.in.\n")
    );
  }

  const html = resp.getContentText();

  // Detect login redirect (page returned without being logged in on a private screen)
  // Detect login redirect — login page has no company rows AND no query form
  const hasRows  = html.indexOf('data-row-company-id') !== -1;
  const hasQuery = html.indexOf('name="query"') !== -1 || html.indexOf('query-builder') !== -1;
  const hasLogin = html.indexOf('action="/login/"') !== -1 || html.indexOf('/accounts/login/') !== -1;

  if (hasLogin && !hasRows) {
    throw new Error(
      "Screen " + screenId + ": Login required.\n" +
      "  → Set CONFIG.SCREENER_COOKIE with your sessionid."
    );
  }

  // Valid page with zero results — two layouts:
  //  (a) Query-builder screen:    has name="query" or query-builder
  //  (b) Paginated-results screen: has data-page-results or data-page-info (0 results)
  const hasPageResults = html.indexOf('data-page-results') !== -1
                      || html.indexOf('data-page-info') !== -1;

  if (!hasRows && (hasQuery || hasPageResults)) {
    return []; // empty screen — not an error
  }

  if (!hasRows && !hasQuery) {
    // Log first 500 chars of the page so we can diagnose what came back
    const preview = html.trim().substring(0, 500).replace(/\s+/g, " ");
    // Could be a redirect page, Cloudflare challenge, or empty screen with different HTML
    // Treat as empty rather than crashing — log a warning so it's visible
    Logger.log("⚠️ Screen " + screenId + " — unexpected page, treating as empty.\nPreview: " + preview);
    return [];
  }

  // Extract all company rows: href gives NSE ticker OR BSE code, link text = name
  // Pattern: <a href="/company/TICKER_OR_BSECODE/[consolidated/]" ...>Name</a>
  const rowPattern = /href="\/company\/([^/"]+)\/?(?:consolidated\/)?"\s[^>]*>([^<]+)<\/a>/g;
  const stocks     = [];
  let   match;

  while ((match = rowPattern.exec(html)) !== null) {
    const rawCode = match[1].trim();
    const name    = match[2].trim();
    if (!name) continue;

    // NSE ticker → use directly; BSE numeric code → store in bseCode for Yahoo fallback
    const isNumeric = /^\d+$/.test(rawCode);
    const symbol    = isNumeric ? "" : rawCode.toUpperCase();
    const bseCode   = isNumeric ? rawCode : "";

    // Dedupe by name
    if (!stocks.some(function(s) { return s.name === name; })) {
      stocks.push({ symbol: symbol, name: name, bseCode: bseCode });
    }
  }

  if (stocks.length === 0 && hasRows) {
    throw new Error("Screen " + screenId + ": Rows found in HTML but regex matched nothing — structure may have changed.");
  }

  return stocks;
}

// ═════════════════════════════════════════════════════════════════════════════
// YAHOO FINANCE — SYMBOL SEARCH BY COMPANY NAME
// ═════════════════════════════════════════════════════════════════════════════
function searchYahooSymbol(companyName, bseCode) {
  if (!companyName && !bseCode) return null;
  const nsSuffix = CONFIG.YF_SUFFIX;       // e.g. ".NS"
  const boSuffix = ".BO";                  // BSE suffix

  function doSearch(q) {
    const url = "https://query2.finance.yahoo.com/v1/finance/search?q=" +
                encodeURIComponent(q) + "&region=IN&quotesCount=8&lang=en-IN";
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: { "User-Agent": "Mozilla/5.0 (compatible; GAS/1.0)" },
        muteHttpExceptions: true,
      });
      if (resp.getResponseCode() !== 200) return null;
      return JSON.parse(resp.getContentText()).quotes || [];
    } catch (_) { return null; }
  }

  // Helper: verify a ticker actually exists on Yahoo chart API
  function tickerExists(ticker) {
    try {
      const url = "https://query2.finance.yahoo.com/v8/finance/chart/" +
                  encodeURIComponent(ticker) + "?range=5d&interval=1d";
      const resp = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: { "User-Agent": "Mozilla/5.0 (compatible; GAS/1.0)" },
        muteHttpExceptions: true,
      });
      if (resp.getResponseCode() !== 200) return false;
      const result = JSON.parse(resp.getContentText())?.chart?.result?.[0];
      return !!(result && result.meta && result.meta.symbol);
    } catch (_) { return false; }
  }

  // 1. For numeric BSE codes: try CODE.BO directly (most reliable for BSE-only stocks)
  if (bseCode && /^\d+$/.test(bseCode)) {
    if (tickerExists(bseCode + boSuffix)) {
      // Also check if an NSE listing exists via Yahoo search
      const quotes = doSearch(bseCode);
      if (quotes) {
        const ns = quotes.find(function(q) {
          return q.quoteType === "EQUITY" && q.symbol && q.symbol.endsWith(nsSuffix);
        });
        if (ns) return ns.symbol.replace(nsSuffix, "");
      }
      return "BSE:" + bseCode;   // confirmed BSE-only
    }
  }

  // 2. Try BSE code as text search → prefer NSE listing
  if (bseCode) {
    const quotes = doSearch(bseCode);
    if (quotes) {
      const ns = quotes.find(function(q) {
        return q.quoteType === "EQUITY" && q.symbol && q.symbol.endsWith(nsSuffix);
      });
      if (ns) return ns.symbol.replace(nsSuffix, "");
      const bo = quotes.find(function(q) {
        return q.quoteType === "EQUITY" && q.symbol && q.symbol.endsWith(boSuffix);
      });
      if (bo) return "BSE:" + bo.symbol.replace(boSuffix, "");
    }
  }

  // 3. Search by company name + exchange suffix
  if (companyName) {
    const suffix = nsSuffix.replace(".", "");
    const namesToTry = [companyName];
    if (companyName.length <= 20 || companyName.endsWith(".")) {
      const shorter = companyName.replace(/[\.\s]+[^\s]+$/, "").trim();
      if (shorter && shorter !== companyName) namesToTry.push(shorter);
    }

    for (const name of namesToTry) {
      const quotes = doSearch(name + " " + suffix);
      if (!quotes) continue;
      const ns = quotes.find(function(q) {
        return q.quoteType === "EQUITY" && q.symbol && q.symbol.endsWith(nsSuffix);
      });
      if (ns) return ns.symbol.replace(nsSuffix, "");
    }
  }

  return null;
}

// ═════════════════════════════════════════════════════════════════════════════
// YAHOO FINANCE — 3-YEAR DAILY CLOSE HISTORY
// ═════════════════════════════════════════════════════════════════════════════
function fetchYahooHistory(symbol) {
  if (!symbol || symbol === NO_SYMBOL_SENTINEL) return null;

  // BSE-only stocks are stored as "BSE:TICKER" — use .BO suffix for those
  const isBse    = symbol.startsWith("BSE:");
  const rawTick  = isBse ? symbol.slice(4) : symbol;          // strip "BSE:" prefix

  // For NSE stocks, also try .BO as a fallback (some are listed on both)
  const suffixes = isBse ? [".BO"] : [CONFIG.YF_SUFFIX, ".BO"];

  const opts = {
    method: "GET",
    headers: { "User-Agent": "Mozilla/5.0 (compatible; GAS/1.0)" },
    muteHttpExceptions: true,
  };

  for (const sfx of suffixes) {
    const ticker = encodeURIComponent(rawTick + sfx);
    const path   = "/v8/finance/chart/" + ticker + "?range=3y&interval=1d";

    for (const host of ["query2", "query1"]) {
      try {
        const resp = UrlFetchApp.fetch("https://" + host + ".finance.yahoo.com" + path, opts);
        if (resp.getResponseCode() !== 200) continue;

        const result = JSON.parse(resp.getContentText())?.chart?.result?.[0];
        if (!result) continue;

        const quote = result.indicators?.quote?.[0] || {};
        return {
          closes: toNumberSeries(quote.close || []),
          highs: toNumberSeries(quote.high || []),
          lows: toNumberSeries(quote.low || []),
          volumes: toNumberSeries(quote.volume || []),
          timestamps: result.timestamp || [],
          usedSuffix: sfx,
        };
      } catch (_) { /* try next */ }
    }
  }
  return null;
}

// ═════════════════════════════════════════════════════════════════════════════
// UPDATE PERFORMANCE FOR ALL STOCKS IN A SHEET
// ═════════════════════════════════════════════════════════════════════════════
function updateSheetPerformance(ss, sheet, RUN_START, scannerName, runContext) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Bulk-read ALL data once
  const allData    = sheet.getRange(2, 1, lastRow - 1, C._COUNT).getValues();
  const updated    = allData.map(r => r.slice()); // mutable copy
  const firstFetch = [];  // row indices (0-based in updated[]) that had no capture price before
  const historyRows = [];

  for (let idx = 0; idx < allData.length; idx++) {

    // Global timeout check — protects against GAS 6-min hard kill
    if (isTimedOut(RUN_START)) {
      log(ss, `⏱ Performance update stopped at row ${idx + 2} (global timeout)`);
      break;
    }

    const row    = allData[idx];
    const symbol = String(row[C.SYMBOL] || "").trim();
    const name   = String(row[C.NAME]   || "").trim();

    // Skip rows with no usable data
    if (!symbol && !name) continue;

    // Sentinel: "~NOFOUND" or "~NOFOUND|BSE:542669" (with BSE code to retry)
    if (symbol === NO_SYMBOL_SENTINEL || symbol.startsWith(NO_SYMBOL_SENTINEL + "|")) {
      // If there is a stored BSE code, attempt one retry
      const parts   = symbol.split("|");
      const bseHint = parts[1] || "";   // e.g. "BSE:542669"
      const bseCode = bseHint.startsWith("BSE:") ? bseHint.slice(4) : "";
      if (bseCode) {
        const retried = searchYahooSymbol(name, bseCode);
        if (retried) {
          updated[idx][C.SYMBOL] = retried;
          log(ss, "  🔎 Retry resolved (BSE:" + bseCode + ") '" + name + "' → " + retried);
          // fall through to performance update below using retried symbol
        } else {
          updated[idx][C.SIGNAL] = "⚠️ Symbol Not Found";
          continue;
        }
      } else {
        updated[idx][C.SIGNAL] = "⚠️ Symbol Not Found";
        continue;
      }
    }

    try {
      // ── Resolve symbol if still blank (shouldn't happen after add-time fix,
      //    but handles rows that existed before v4 upgrade) ──────────────────
      let sym = symbol;
      if (!sym && name) {
        sym = searchYahooSymbol(name) || NO_SYMBOL_SENTINEL;
        updated[idx][C.SYMBOL] = sym;
        if (sym === NO_SYMBOL_SENTINEL) {
          updated[idx][C.SIGNAL] = "⚠️ Symbol Not Found";
          continue;
        }
        log(ss, `  🔎 Resolved (legacy row) "${name}" → ${sym}`);
      }

      // Use updated[idx] symbol in case retry above changed it
      sym = String(updated[idx][C.SYMBOL] || sym).trim();

      // ── Fetch 3yr daily OHLCV history ─────────────────────────────────────
      const rawHist = fetchYahooHistory(sym);
      const hist    = normalizeHistory(rawHist);
      if (!hist || hist.closes.length < 5) {
        updated[idx][C.SIGNAL] = "⚠️ No Data";
        continue;
      }

      const closes  = hist.closes;
      const n       = closes.length;
      const current = closes[n - 1];

      if (!current || isNaN(current) || current <= 0) {
        updated[idx][C.SIGNAL] = "⚠️ Bad Price";
        continue;
      }

      // ── Capture price (set ONCE on first successful fetch) ─────────────────
      const hadCapturePrice = row[C.CAPTURE_PRICE] !== "" && row[C.CAPTURE_PRICE] > 0;
      if (!hadCapturePrice) {
        updated[idx][C.CAPTURE_PRICE] = roundN(current, 2);
        firstFetch.push(idx);  // remember for highlight-clear later
      }
      const capturePrice = updated[idx][C.CAPTURE_PRICE];

      // ── Period returns ────────────────────────────────────────────────────
      updated[idx][C.CURRENT_PRICE] = roundN(current, 2);
      updated[idx][C.RET_CAPTURE]   = capturePrice > 0 ? pct(((current - capturePrice) / capturePrice) * 100) : "";
      updated[idx][C.RET_1D]        = pct(periodReturn(closes, n, 1));
      updated[idx][C.RET_1W]        = pct(periodReturn(closes, n, 5));
      updated[idx][C.RET_1M]        = pct(periodReturn(closes, n, 21));
      updated[idx][C.RET_3M]        = pct(periodReturn(closes, n, 63));
      updated[idx][C.RET_6M]        = pct(periodReturn(closes, n, 126));
      updated[idx][C.RET_1Y]        = pct(periodReturn(closes, n, 252));
      updated[idx][C.RET_2Y]        = pct(periodReturn(closes, n, 504));
      updated[idx][C.RET_3Y]        = pct(periodReturn(closes, n, 756));

      // ── Average period returns ────────────────────────────────────────────
      updated[idx][C.AVG_WEEKLY]  = pct(avgPeriodReturn(closes, 5,   252));  // avg weekly over 1yr
      updated[idx][C.AVG_MONTHLY] = pct(avgPeriodReturn(closes, 21,  504));  // avg monthly over 2yr
      updated[idx][C.AVG_3M]      = pct(avgPeriodReturn(closes, 63,  n));    // avg 3M over all data
      updated[idx][C.AVG_6M]      = pct(avgPeriodReturn(closes, 126, n));    // avg 6M over all data
      updated[idx][C.AVG_1Y]      = pct(avgPeriodReturn(closes, 252, n));    // avg 1Y over all data

      // ── Technical indicators ──────────────────────────────────────────────
      const ma20  = calcMA(closes, 20);
      const ma50  = calcMA(closes, 50);
      const ma200 = calcMA(closes, 200);
      const rsi   = calcRSI(closes, 14);
      const adx   = calcADX(hist.highs, hist.lows, closes, 14);
      const volRatio = calcVolumeRatio(hist.volumes, 20, hist.timestamps);
      const macd  = calcMACD(closes, 12, 26, 9);
      const dist52wHigh = calcDistanceFromHigh(hist.highs, closes, 252);
      const breakout20dPct = calcBreakoutPct(hist.highs, closes, 20);

      updated[idx][C.MA20]         = roundN(ma20, 2);
      updated[idx][C.MA50]         = roundN(ma50, 2);
      updated[idx][C.MA200]        = roundN(ma200, 2);
      updated[idx][C.RSI14]        = roundN(rsi, 1);
      updated[idx][C.ADX14]        = roundN(adx, 1);
      updated[idx][C.VOL_RATIO20]  = roundN(volRatio, 2);
      updated[idx][C.MACD_LINE]    = roundN(macd.macdLine, 2);
      updated[idx][C.MACD_HIST]    = roundN(macd.histogram, 2);
      updated[idx][C.DIST_52W_HIGH]= pct(dist52wHigh);
      updated[idx][C.BREAKOUT_20D] = pct(breakout20dPct);
      updated[idx][C.SIGNAL]       = calcSignal({
        price: current,
        ma20: ma20,
        ma50: ma50,
        ma200: ma200,
        rsi: rsi,
        adx: adx,
        volRatio: volRatio,
        macdLine: macd.macdLine,
        macdHist: macd.histogram,
        dist52wHigh: dist52wHigh,
        breakout20dPct: breakout20dPct,
      });
      updated[idx][C.LAST_UPDATED] = formatDate(new Date());
      historyRows.push(buildPriceHistoryRow(
        runContext || buildRunContext("AUTO"),
        scannerName || sheet.getName(),
        updated[idx]
      ));

      Utilities.sleep(350); // stay well under Yahoo Finance rate limits

    } catch (e) {
      log(ss, `  ⚠️ ${symbol || name}: ${e.message}`);
      updated[idx][C.SIGNAL] = "⚠️ Error";
    }
  }

  // ── Bulk-write ALL rows (1 Sheets API call) ───────────────────────────────
  sheet.getRange(2, 1, updated.length, C._COUNT).setValues(updated);

  // ── Batch all formatting — signal colors + first-fetch highlight clear ────
  // FIX: was N individual getRange() calls; now 2 range calls total
  applyBatchFormatting(sheet, updated, firstFetch);
  try {
    appendPriceHistoryRows(ss, historyRows);
  } catch (e) {
    log(ss, `⚠️ Price history append failed [${scannerName || sheet.getName()}]: ${e.message}`);
  }
}

// ═════════════════════════════════════════════════════════════════════════════
// CALCULATIONS
// ═════════════════════════════════════════════════════════════════════════════

/** Trim NaN from the end of the price series (handles incomplete current bar) */
function trimTrailingNaN(closes) {
  let end = closes.length;
  while (end > 0 && isNaN(closes[end - 1])) end--;
  return closes.slice(0, end);
}

/** Convert an API array to Numbers, using NaN for missing values. */
function toNumberSeries(values) {
  return (values || []).map(function(v) {
    return v == null ? NaN : Number(v);
  });
}

/**
 * Align Yahoo arrays to the same usable range by trimming any incomplete
 * trailing bar where the close is still missing.
 */
function normalizeHistory(hist) {
  if (!hist || !hist.closes || hist.closes.length === 0) return null;
  const end = trimTrailingNaN(hist.closes).length;
  if (end === 0) return null;
  return {
    closes: hist.closes.slice(0, end),
    highs: alignSeries(hist.highs, end),
    lows: alignSeries(hist.lows, end),
    volumes: alignSeries(hist.volumes, end),
    timestamps: (hist.timestamps || []).slice(0, end),
    usedSuffix: hist.usedSuffix,
  };
}

/** Return a fixed-length numeric series padded with NaN for missing values. */
function alignSeries(values, length) {
  const src = values || [];
  const out = [];
  for (let i = 0; i < length; i++) {
    const raw = src[i];
    const num = raw == null ? NaN : Number(raw);
    out.push(isNaN(num) ? NaN : num);
  }
  return out;
}

/** Point-to-point % return: closes[n-1] vs closes[n-1-periods] */
function periodReturn(closes, n, periods) {
  if (n <= periods) return null;
  const past = closes[n - 1 - periods];
  const cur  = closes[n - 1];
  if (isNaN(past) || isNaN(cur) || past === 0) return null;
  return ((cur - past) / past) * 100;
}

/**
 * Average of all non-overlapping period returns within a lookback window.
 * @param {number[]} closes   Full close series (NaN gaps are skipped)
 * @param {number}   period   Bars per one period (e.g. 5 = 1 week)
 * @param {number}   lookback How many bars from the end to include
 */
function avgPeriodReturn(closes, period, lookback) {
  // FIX: param renamed from `window` to `lookback` (avoided shadowing)
  const start = Math.max(0, closes.length - lookback);
  const slice = closes.slice(start);
  const n     = slice.length;
  if (n < period * 2) return null;  // need at least 2 complete periods

  const rets = [];
  for (let i = period; i < n; i += period) {
    const s = slice[i - period];
    const e = slice[i];
    // FIX: use isNaN explicitly (not falsy check) — catches valid prices near 0
    if (isNaN(s) || isNaN(e) || s === 0) continue;
    rets.push(((e - s) / s) * 100);
  }
  if (rets.length === 0) return null;
  return rets.reduce((a, b) => a + b, 0) / rets.length;
}

/** SMA of the last `period` valid (non-NaN) closes. Requires 80% valid bars. */
function calcMA(closes, period) {
  const n = closes.length;
  if (n < period) return null;
  const slice = closes.slice(n - period).filter(v => !isNaN(v));
  if (slice.length < Math.ceil(period * 0.8)) return null;
  return slice.reduce((a, b) => a + b, 0) / slice.length;
}

/**
 * RSI (Wilder's smoothing, 14-period standard).
 * NaN closes are skipped before computing changes to prevent NaN propagation.
 */
function calcRSI(closes, period) {
  // Build changes from consecutive non-NaN pairs only
  const changes = [];
  for (let i = 1; i < closes.length; i++) {
    if (!isNaN(closes[i - 1]) && !isNaN(closes[i])) {
      changes.push(closes[i] - closes[i - 1]);
    }
  }
  if (changes.length < period + 1) return null;

  // Seed with SMA of first `period` changes
  let avgGain = 0, avgLoss = 0;
  for (let i = 0; i < period; i++) {
    if (changes[i] > 0) avgGain += changes[i];
    else                 avgLoss -= changes[i];
  }
  avgGain /= period;
  avgLoss /= period;

  // Wilder exponential smoothing for the rest
  for (let i = period; i < changes.length; i++) {
    const g = changes[i] > 0 ?  changes[i] : 0;
    const l = changes[i] < 0 ? -changes[i] : 0;
    avgGain = (avgGain * (period - 1) + g) / period;
    avgLoss = (avgLoss * (period - 1) + l) / period;
  }

  if (avgLoss === 0) return 100;
  return 100 - 100 / (1 + avgGain / avgLoss);
}

/** Exponential moving average series using SMA seed. */
function calcEMA(values, period) {
  if (!values || values.length < period) return [];
  const ema = new Array(values.length).fill(null);
  let sum = 0;
  for (let i = 0; i < period; i++) {
    const v = Number(values[i]);
    if (isNaN(v)) return [];
    sum += v;
  }
  ema[period - 1] = sum / period;
  const k = 2 / (period + 1);
  for (let i = period; i < values.length; i++) {
    const v = Number(values[i]);
    if (isNaN(v)) return [];
    ema[i] = ((v - ema[i - 1]) * k) + ema[i - 1];
  }
  return ema;
}

/** MACD (12,26,9) on valid close prices. */
function calcMACD(closes, fastPeriod, slowPeriod, signalPeriod) {
  const values = closes.filter(v => !isNaN(v));
  if (values.length < slowPeriod + signalPeriod) {
    return { macdLine: null, signalLine: null, histogram: null };
  }

  const fast = calcEMA(values, fastPeriod);
  const slow = calcEMA(values, slowPeriod);
  const macdSeries = [];

  for (let i = 0; i < values.length; i++) {
    if (fast[i] != null && slow[i] != null) {
      macdSeries.push(fast[i] - slow[i]);
    }
  }

  if (macdSeries.length === 0) {
    return { macdLine: null, signalLine: null, histogram: null };
  }

  const signalSeries = calcEMA(macdSeries, signalPeriod);
  const macdLine = macdSeries[macdSeries.length - 1];
  const signalLine = lastValidNumber(signalSeries);

  return {
    macdLine: macdLine,
    signalLine: signalLine,
    histogram: signalLine == null ? null : macdLine - signalLine,
  };
}

/**
 * Latest completed-session volume divided by the average of the previous
 * `period` completed sessions. During live market hours, the current partial
 * day is ignored to avoid underestimating participation.
 */
function calcVolumeRatio(volumes, period, timestamps) {
  const rows = [];
  for (let i = 0; i < volumes.length; i++) {
    const volume = Number(volumes[i]);
    const ts = timestamps && timestamps[i] ? Number(timestamps[i]) : null;
    if (!isNaN(volume) && volume > 0) {
      rows.push({ volume: volume, timestamp: ts });
    }
  }
  if (rows.length < period + 1) return null;

  let currentIdx = rows.length - 1;
  const latest = rows[currentIdx];
  if (
    latest &&
    latest.timestamp &&
    isLiveExchangeSessionNow() &&
    todayISTString(toISTDate(new Date(latest.timestamp * 1000))) === todayISTString()
  ) {
    currentIdx--;
  }

  if (currentIdx < period) return null;

  const current = rows[currentIdx].volume;
  const baseline = rows.slice(currentIdx - period, currentIdx).map(function(row) {
    return row.volume;
  });
  const avg = baseline.reduce((a, b) => a + b, 0) / baseline.length;
  if (!avg || isNaN(avg)) return null;
  return current / avg;
}

/** % distance from the highest traded high in the last `lookback` sessions. */
function calcDistanceFromHigh(highs, closes, lookback) {
  const rows = [];
  for (let i = 0; i < closes.length; i++) {
    const high = highs[i];
    const close = closes[i];
    if (!isNaN(high) && !isNaN(close)) {
      rows.push({ high: high, close: close });
    }
  }
  if (rows.length === 0) return null;

  const slice = rows.slice(-Math.min(rows.length, lookback));
  if (slice.length === 0) return null;
  const current = slice[slice.length - 1].close;
  const high = Math.max.apply(null, slice.map(function(row) { return row.high; }));
  if (isNaN(current) || isNaN(high) || high <= 0) return null;
  return ((high - current) / high) * 100;
}

/** % above/below the previous `lookback` sessions' high (positive = breakout). */
function calcBreakoutPct(highs, closes, lookback) {
  const rows = [];
  for (let i = 0; i < closes.length; i++) {
    const high = highs[i];
    const close = closes[i];
    if (!isNaN(high) && !isNaN(close)) {
      rows.push({ high: high, close: close });
    }
  }
  if (rows.length < lookback + 1) return null;

  const currentClose = rows[rows.length - 1].close;
  const priorHighs = rows.slice(rows.length - lookback - 1, rows.length - 1).map(r => r.high);
  const breakoutLevel = Math.max.apply(null, priorHighs);

  if (isNaN(currentClose) || isNaN(breakoutLevel) || breakoutLevel <= 0) return null;
  return ((currentClose - breakoutLevel) / breakoutLevel) * 100;
}

/** ADX (Wilder, 14) using high / low / close history. */
function calcADX(highs, lows, closes, period) {
  const rows = [];
  for (let i = 0; i < closes.length; i++) {
    const high = highs[i];
    const low = lows[i];
    const close = closes[i];
    if (!isNaN(high) && !isNaN(low) && !isNaN(close)) {
      rows.push({ high: high, low: low, close: close });
    }
  }
  if (rows.length < period * 2) return null;

  const trs = [];
  const plusDMs = [];
  const minusDMs = [];

  for (let i = 1; i < rows.length; i++) {
    const prev = rows[i - 1];
    const cur = rows[i];
    const upMove = cur.high - prev.high;
    const downMove = prev.low - cur.low;

    plusDMs.push(upMove > downMove && upMove > 0 ? upMove : 0);
    minusDMs.push(downMove > upMove && downMove > 0 ? downMove : 0);
    trs.push(Math.max(
      cur.high - cur.low,
      Math.abs(cur.high - prev.close),
      Math.abs(cur.low - prev.close)
    ));
  }

  if (trs.length < (period * 2 - 1)) return null;

  let trSmooth = sumNumbers(trs.slice(0, period));
  let plusSmooth = sumNumbers(plusDMs.slice(0, period));
  let minusSmooth = sumNumbers(minusDMs.slice(0, period));
  const dxs = [];

  for (let i = period - 1; i < trs.length; i++) {
    if (i > period - 1) {
      trSmooth = trSmooth - (trSmooth / period) + trs[i];
      plusSmooth = plusSmooth - (plusSmooth / period) + plusDMs[i];
      minusSmooth = minusSmooth - (minusSmooth / period) + minusDMs[i];
    }

    const plusDI = trSmooth === 0 ? 0 : 100 * (plusSmooth / trSmooth);
    const minusDI = trSmooth === 0 ? 0 : 100 * (minusSmooth / trSmooth);
    const denom = plusDI + minusDI;
    dxs.push(denom === 0 ? 0 : 100 * Math.abs(plusDI - minusDI) / denom);
  }

  if (dxs.length < period) return null;

  let adx = sumNumbers(dxs.slice(0, period)) / period;
  for (let i = period; i < dxs.length; i++) {
    adx = ((adx * (period - 1)) + dxs[i]) / period;
  }
  return adx;
}

/** Last non-null, numeric value in a series. */
function lastValidNumber(values) {
  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i];
    if (v != null && !isNaN(v)) return v;
  }
  return null;
}

/** Sum a numeric array. */
function sumNumbers(values) {
  return values.reduce((sum, val) => sum + val, 0);
}

/** Active signal profile name with fallback to CONFIG default. */
function getSignalProfileName() {
  const stored = PropertiesService.getScriptProperties().getProperty(SIGNAL_PROFILE_KEY);
  if (stored && SIGNAL_PROFILES[stored]) return stored;
  return SIGNAL_PROFILES[CONFIG.SIGNAL_PROFILE] ? CONFIG.SIGNAL_PROFILE : "balanced";
}

/** Selected signal thresholds with safe fallback to the balanced preset. */
function getSignalSettings() {
  return SIGNAL_PROFILES[getSignalProfileName()] || SIGNAL_PROFILES.balanced;
}

/** Persist the active signal profile for future runs. */
function setSignalProfile(profileName) {
  if (!SIGNAL_PROFILES[profileName]) {
    throw new Error("Unknown signal profile: " + profileName);
  }
  PropertiesService.getScriptProperties().setProperty(SIGNAL_PROFILE_KEY, profileName);
  Logger.log("✅ Signal profile set to: " + profileName);
}

function setConservativeProfile() { setSignalProfile("conservative"); }
function setBalancedProfile()     { setSignalProfile("balanced"); }
function setAggressiveProfile()   { setSignalProfile("aggressive"); }

/** Clear any stored profile override and return to CONFIG default. */
function resetSignalProfile() {
  PropertiesService.getScriptProperties().deleteProperty(SIGNAL_PROFILE_KEY);
  Logger.log("✅ Signal profile reset to CONFIG default: " + CONFIG.SIGNAL_PROFILE);
}

/**
 * Signal engine — MA stack + RSI + trend/volume/momentum confirmation.
 *
 * Signal tiers (strongest → weakest):
 *   🚀 BREAKOUT BUY          — perfect stack + strong trend + volume + near 52W high
 *   🚀 STRONG BUY (Pullback) — perfect stack + RSI oversold + confirmation
 *   🚀 STRONG BUY            — perfect stack + RSI 50–75 + confirmation
 *   ✅ BUY                   — bullish stack with MACD / volume confirmation
 *   🟡 BUY (Overbought)      — above both MAs but RSI > 75 (trail stop)
 *   🟡 HOLD (Overbought)     — above MAs but extended
 *   🟡 HOLD (Weak Trend)     — bullish stack but weak confirmation
 *   ⏸️ HOLD (Pullback)       — above MA50, below MA20, MACD still positive
 *   ⏸️ HOLD (Oversold)       — below MA20 but above MA50, oversold
 *   ⏸️ HOLD                  — mixed / transitional
 *   ⚠️ WEAK (Oversold)       — below MA50, oversold (watch for bounce)
 *   ⚠️ SELL (Oversold)       — bear stack but oversold (watch reversal)
 *   🔴 SELL                  — price < MA20 < MA50 (< MA200 if available)
 *
 * FIX: perfectBull now REQUIRES hasMa200. Previously ma200BelowOthers defaulted
 * to `true` when MA200 was null, allowing STRONG BUY on stocks with only 20 days
 * of data. Now: stocks without MA200 can be at most "✅ BUY".
 */
function calcSignal(metrics) {
  const settings = getSignalSettings();
  const price = metrics.price;
  const ma20 = metrics.ma20;
  const ma50 = metrics.ma50;
  const ma200 = metrics.ma200;
  const rsi = metrics.rsi;
  const adx = metrics.adx;
  const volRatio = metrics.volRatio;
  const macdLine = metrics.macdLine;
  const macdHist = metrics.macdHist;
  const dist52wHigh = metrics.dist52wHigh;
  const breakout20dPct = metrics.breakout20dPct;

  if (!price || !ma20 || !ma50 || isNaN(price) || isNaN(ma20) || isNaN(ma50)) {
    return "⚪ Insufficient Data";
  }

  const above20 = price > ma20;
  const above50 = price > ma50;

  // MA200: only factor in when we actually have 200 bars of data
  const hasMa200  = ma200 != null && !isNaN(ma200) && ma200 > 0;
  const above200  = hasMa200 ? price > ma200 : false;   // FIX: false, not true

  // RSI zones
  const hasRsi     = rsi != null && !isNaN(rsi);
  const overbought = hasRsi && rsi >= settings.RSI_OVERBOUGHT;
  const oversold   = hasRsi && rsi <= settings.RSI_OVERSOLD;
  const bullish    = hasRsi && rsi >= settings.RSI_BULLISH_MIN && rsi < settings.RSI_OVERBOUGHT;
  const neutral    = hasRsi && rsi >= settings.RSI_NEUTRAL_MIN && rsi < settings.RSI_BULLISH_MIN;

  // Trend / participation confirmation
  const hasAdx      = adx != null && !isNaN(adx);
  const trendStrong = hasAdx && adx >= settings.ADX_STRONG;
  const trendWeak   = hasAdx && adx < settings.ADX_WEAK;

  const hasVolume   = volRatio != null && !isNaN(volRatio);
  const highVolume  = hasVolume && volRatio >= settings.VOL_RATIO_HIGH;

  const macdLineTol = Math.max(1e-6, Math.abs(price) * settings.MACD_ZERO_TOL_PCT);
  const macdHistTol = Math.max(1e-6, Math.abs(price) * settings.MACD_ZERO_TOL_PCT);
  const hasMacdLine = macdLine != null && !isNaN(macdLine);
  const hasMacdHist = macdHist != null && !isNaN(macdHist);
  const macdTrendBull = hasMacdLine && macdLine > macdLineTol;
  const macdTrendBear = hasMacdLine && macdLine < -macdLineTol;
  const macdAccelBull = hasMacdHist && macdHist > macdHistTol;

  const near52wHigh = dist52wHigh != null && !isNaN(dist52wHigh) && dist52wHigh <= settings.DIST_52W_HIGH_MAX;
  const breakout20d = breakout20dPct != null && !isNaN(breakout20dPct) && breakout20dPct >= settings.BREAKOUT_20D_MIN;

  // MA stack quality
  const stackBull = above20 && above50 && (ma20 > ma50);
  const stackBear = !above20 && !above50 && (ma20 < ma50);

  // FIX: perfectBull REQUIRES MA200 (price > MA20 > MA50 > MA200)
  // Stocks without enough history for MA200 cannot be "perfect bull".
  const perfectBull = hasMa200 && stackBull && above200 && (ma200 < ma50);
  const perfectBear = stackBear && (!hasMa200 || !above200);
  const bullConfirmed = macdTrendBull || highVolume || breakout20d || (!hasMacdLine && !hasVolume);
  const strongBullConfirmed = (trendStrong || !hasAdx) && bullConfirmed;

  // ── Signal tree ───────────────────────────────────────────────────────────
  if (perfectBull && (near52wHigh || breakout20d) && highVolume && (macdTrendBull || macdAccelBull || breakout20d || !hasMacdLine) && (trendStrong || !hasAdx)) {
    return "🚀 BREAKOUT BUY";
  }
  if (perfectBull && oversold && strongBullConfirmed) return "🚀 STRONG BUY (Pullback)";
  if (perfectBull && bullish && strongBullConfirmed)  return "🚀 STRONG BUY";
  if (perfectBull && (neutral || !hasRsi) && bullConfirmed) return "✅ BUY";
  if (perfectBull && overbought)            return "🟡 BUY (Overbought — trail SL)";

  // Above both MAs but not "perfect" stack (e.g. MA200 missing or mis-ordered)
  if (stackBull && overbought)              return "🟡 HOLD (Overbought — watch)";
  if (stackBull && bullConfirmed && !trendWeak) return "✅ BUY";
  if (stackBull)                            return "🟡 HOLD (Weak Trend)";

  // Below MA20 but above MA50 — transitional
  if (!above20 && above50 && (macdTrendBull || macdAccelBull)) return "⏸️ HOLD (Pullback)";
  if (!above20 && above50 && oversold)      return "⏸️ HOLD (Oversold — possible bounce)";
  if (!above20 && above50)                  return "⏸️ HOLD";

  // Bear territory
  if (perfectBear && oversold && (macdAccelBull || !macdTrendBear)) return "⚠️ SELL (Oversold — watch reversal)";
  if (perfectBear)                          return "🔴 SELL";
  if (!above50 && oversold)                 return "⚠️ WEAK (Oversold — possible bounce)";
  if (!above20 && !above50)                 return "🔴 SELL";

  return "⏸️ HOLD";
}

// ═════════════════════════════════════════════════════════════════════════════
// BATCH FORMATTING  (replaces N individual getRange() calls)
// ═════════════════════════════════════════════════════════════════════════════

/**
 * Apply signal cell colors + clear first-fetch highlights, all in batch.
 * FIX: was N individual getRange().setBackground() calls per row.
 *      Now: 2 range-level calls (setBackgrounds + setFontColors) regardless of N.
 */
function applyBatchFormatting(sheet, updatedData, firstFetchIndices) {
  const n = updatedData.length;
  if (n === 0) return;

  // Build parallel BG + font color arrays for the SIGNAL column
  const signalBgs    = [];
  const signalFonts  = [];

  for (let idx = 0; idx < n; idx++) {
    const sig = String(updatedData[idx][C.SIGNAL] || "");
    const [bg, font] = signalColor(sig);
    signalBgs.push([bg]);
    signalFonts.push([font]);
  }

  // Write both arrays in 2 calls
  const sigRange = sheet.getRange(2, C.SIGNAL + 1, n, 1);
  sigRange.setBackgrounds(signalBgs);
  sigRange.setFontColors(signalFonts);

  // Clear the yellow "new stock" highlight for rows that just got their first price.
  // We must do this AFTER the bulk setValues() write (already done by caller).
  // Batch: collect indices, build a single background range per contiguous block
  // (simpler: just iterate and call setBackground on each affected row — the
  //  number of first-fetch rows is typically tiny: 0–5 per run)
  firstFetchIndices.forEach(idx => {
    sheet.getRange(idx + 2, 1, 1, C._COUNT).setBackground(null);
  });
}

/** Returns [backgroundColor, fontColor] for a given signal string */
function signalColor(sig) {
  if      (sig.includes("BREAKOUT BUY"))        return ["#6A1B9A", "#FFFFFF"];
  else if (sig.includes("STRONG BUY (Pull"))    return ["#00695C", "#FFFFFF"];
  else if (sig.includes("STRONG BUY"))           return ["#1B5E20", "#FFFFFF"];
  else if (sig.includes("✅ BUY"))               return ["#4CAF50", "#FFFFFF"];
  else if (sig.includes("BUY (Overbought"))      return ["#A5D6A7", "#1B5E20"];
  else if (sig.includes("HOLD (Overbought"))     return ["#FFE082", "#BF360C"];
  else if (sig.includes("HOLD (Weak Trend"))     return ["#FFF3E0", "#E65100"];
  else if (sig.includes("HOLD (Pullback"))       return ["#B2DFDB", "#004D40"];
  else if (sig.includes("HOLD (Oversold"))       return ["#B2DFDB", "#004D40"];
  else if (sig.includes("⏸️ HOLD"))              return ["#FFF9C4", "#F57F17"];
  else if (sig.includes("SELL (Oversold"))       return ["#FF8A65", "#FFFFFF"];
  else if (sig.includes("🔴 SELL"))              return ["#F44336", "#FFFFFF"];
  else if (sig.includes("WEAK"))                 return ["#FFCCBC", "#BF360C"];
  else                                           return ["#ECEFF1", "#546E7A"];
}

/**
 * Apply green/red conditional formatting to return columns.
 * Clears all existing rules first to prevent accumulation on repeated calls.
 */
function applyConditionalFormatting(sheet) {
  const maxRow = Math.max(sheet.getLastRow(), 2);
  sheet.clearConditionalFormatRules();

  const rules = [];
  RETURN_COLS.forEach(col => {
    const range = sheet.getRange(2, col + 1, maxRow - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground("#E8F5E9").setFontColor("#1B5E20")
        .setRanges([range]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0).setBackground("#FFEBEE").setFontColor("#C62828")
        .setRanges([range]).build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}

// ═════════════════════════════════════════════════════════════════════════════
// BATCH SHEET HELPERS
// ═════════════════════════════════════════════════════════════════════════════

/** Set entire IN_SCREENER column to "⬜" in 1 write call */
function batchMarkNotInScreener(sheet, lastRow) {
  if (lastRow < 2) return;
  const numRows = lastRow - 1;
  sheet.getRange(2, C.IN_SCREENER + 1, numRows, 1)
       .setValues(Array.from({ length: numRows }, () => ["⬜"]));
}

/**
 * Apply scattered row updates to a single column in bulk (2 API calls total).
 * 1. Read the current column values (1 call)
 * 2. Patch in-memory
 * 3. Write back (1 call)
 */
function batchWriteColumn(sheet, updates, colNum, lastRow) {
  if (updates.length === 0 || lastRow < 2) return;
  const numRows = lastRow - 1;
  const range   = sheet.getRange(2, colNum, numRows, 1);
  const vals    = range.getValues();
  updates.forEach(({ rowIdx, value }) => {
    const i = rowIdx - 2;  // 0-based index into data array
    if (i >= 0 && i < vals.length) vals[i][0] = value;
  });
  range.setValues(vals);
}

// ═════════════════════════════════════════════════════════════════════════════
// SHEET MANAGEMENT
// ═════════════════════════════════════════════════════════════════════════════

/** Normalize sheet names so invisible unicode / spacing differences do not break lookups. */
function normalizeSheetName(name) {
  let value = String(name || "");
  try { value = value.normalize("NFKC"); } catch (_) { /* normalize may be unavailable */ }
  return value
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFE0E\uFE0F]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/** Exact lookup first, then normalized fallback across all sheets. */
function getSheetByNameSafe(ss, targetName) {
  const exact = ss.getSheetByName(targetName);
  if (exact) return exact;

  const targetNorm = normalizeSheetName(targetName);
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (normalizeSheetName(sheets[i].getName()) === targetNorm) {
      return sheets[i];
    }
  }
  return null;
}

function getOrCreateScannerSheet(ss, scanner) {
  let sheet = getSheetByNameSafe(ss, scanner.name);
  if (!sheet) {
    sheet = ss.insertSheet(scanner.name);
    log(ss, `📋 Created sheet: ${scanner.name}`);
  }
  if (scanner.color) sheet.setTabColor(scanner.color); // setTabColor requires "#" prefix
  ensureScannerSheetSchema(sheet);
  return sheet;
}

function ensureScannerSheetSchema(sheet) {
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }
  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsRewrite = HEADERS.some(function(header, idx) {
    return currentHeaders[idx] !== header;
  });
  if (needsRewrite) writeHeaders(sheet);
}

function writeHeaders(sheet) {
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }
  const r = sheet.getRange(1, 1, 1, HEADERS.length);
  r.setValues([HEADERS]);
  r.setBackground("#1A237E").setFontColor("#FFFFFF")
   .setFontWeight("bold").setFontSize(10).setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // Symbol + Name always visible

  // Column widths
  [80,160,110,110,85, 95,95,90, 60,60,60,60,60,60,60,60, 90,90,80,80,80, 70,80,80,80, 160,125, 70,85,85,85,95,95]
    .forEach((w, i) => { if (i < HEADERS.length) sheet.setColumnWidth(i + 1, w); });

  const formatRows = Math.max(sheet.getMaxRows() - 1, 1000);

  // FIX: set numeric "0.00" format on all return columns so they display cleanly
  RETURN_COLS.forEach(col => {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("0.00");
  });
  // MA columns as currency
  PRICE_COLS.forEach(col => {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("₹#,##0.00");
  });
  SINGLE_DECIMAL_COLS.forEach(col => {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("0.0");
  });
  DOUBLE_DECIMAL_COLS.forEach(col => {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("0.00");
  });

  applyConditionalFormatting(sheet);
}

function ensurePriceHistorySheetSchema(sheet) {
  if (sheet.getMaxColumns() < PRICE_HISTORY_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), PRICE_HISTORY_HEADERS.length - sheet.getMaxColumns());
  }
  const currentHeaders = sheet.getRange(1, 1, 1, PRICE_HISTORY_HEADERS.length).getValues()[0];
  const needsRewrite = PRICE_HISTORY_HEADERS.some(function(header, idx) {
    return currentHeaders[idx] !== header;
  });
  if (needsRewrite) writePriceHistoryHeaders(sheet);
}

function writePriceHistoryHeaders(sheet) {
  if (sheet.getMaxColumns() < PRICE_HISTORY_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), PRICE_HISTORY_HEADERS.length - sheet.getMaxColumns());
  }
  const r = sheet.getRange(1, 1, 1, PRICE_HISTORY_HEADERS.length);
  r.setValues([PRICE_HISTORY_HEADERS]);
  r.setBackground("#004D40").setFontColor("#FFFFFF")
   .setFontWeight("bold").setWrap(true);
  sheet.setFrozenRows(1);

  [140, 95, 160, 170, 90, 180, 85, 95, 95, 90, 70, 70, 70, 70, 70, 70, 70, 70, 85, 85, 85, 95, 95, 170]
    .forEach(function(w, i) { if (i < PRICE_HISTORY_HEADERS.length) sheet.setColumnWidth(i + 1, w); });

  const formatRows = Math.max(sheet.getMaxRows() - 1, 1000);
  sheet.getRange(2, H.SNAPSHOT_AT + 1, formatRows, 1).setNumberFormat("dd mmm yyyy hh:mm");
  [H.CAPTURE_PRICE, H.CURRENT_PRICE].forEach(function(col) {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("₹#,##0.00");
  });
  [H.RET_CAPTURE, H.RET_1D, H.RET_1W, H.RET_1M, H.RET_3M, H.RET_6M, H.RET_1Y, H.DIST_52W_HIGH, H.BREAKOUT_20D]
    .forEach(function(col) {
      sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("0.00");
    });
  [H.RSI14, H.ADX14].forEach(function(col) {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("0.0");
  });
  [H.VOL_RATIO20, H.MACD_LINE, H.MACD_HIST].forEach(function(col) {
    sheet.getRange(2, col + 1, formatRows, 1).setNumberFormat("0.00");
  });
}

function ensureChartDataSheetSchema(sheet) {
  const headers = [
    "Snapshot At", "Price ₹",
    "Snapshot At", "Since Capture %",
    "Snapshot At", "1D %", "1W %", "1M %",
    "Signal"
  ];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsRewrite = headers.some(function(header, idx) {
    return currentHeaders[idx] !== header;
  });
  if (needsRewrite) {
    const r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setBackground("#263238").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
    const formatRows = Math.max(sheet.getMaxRows() - 1, 1000);
    sheet.getRange(2, 1, formatRows, 1).setNumberFormat("dd mmm yyyy hh:mm");
    sheet.getRange(2, 2, formatRows, 1).setNumberFormat("₹#,##0.00");
    sheet.getRange(2, 3, formatRows, 1).setNumberFormat("dd mmm yyyy hh:mm");
    sheet.getRange(2, 4, formatRows, 1).setNumberFormat("0.00");
    sheet.getRange(2, 5, formatRows, 1).setNumberFormat("dd mmm yyyy hh:mm");
    sheet.getRange(2, 6, formatRows, 3).setNumberFormat("0.00");
  }
}

function buildStockKey(symbol, name) {
  const sym = String(symbol || "").trim();
  if (sym && sym !== NO_SYMBOL_SENTINEL && !sym.startsWith(NO_SYMBOL_SENTINEL + "|")) {
    return "SYM:" + sym;
  }
  const normalizedName = normalizeSheetName(name);
  return normalizedName ? "NAME:" + normalizedName : "";
}

function buildPriceHistoryRow(runContext, scannerName, row) {
  const ctx = runContext || buildRunContext("AUTO");
  return [
    ctx.snapshotTime,
    ctx.mode,
    scannerName || "",
    buildStockKey(row[C.SYMBOL], row[C.NAME]),
    row[C.SYMBOL],
    row[C.NAME],
    row[C.IN_SCREENER],
    numOrBlank(row[C.CAPTURE_PRICE]),
    numOrBlank(row[C.CURRENT_PRICE]),
    numOrBlank(row[C.RET_CAPTURE]),
    numOrBlank(row[C.RET_1D]),
    numOrBlank(row[C.RET_1W]),
    numOrBlank(row[C.RET_1M]),
    numOrBlank(row[C.RET_3M]),
    numOrBlank(row[C.RET_6M]),
    numOrBlank(row[C.RET_1Y]),
    numOrBlank(row[C.RSI14]),
    numOrBlank(row[C.ADX14]),
    numOrBlank(row[C.VOL_RATIO20]),
    numOrBlank(row[C.MACD_LINE]),
    numOrBlank(row[C.MACD_HIST]),
    numOrBlank(row[C.DIST_52W_HIGH]),
    numOrBlank(row[C.BREAKOUT_20D]),
    row[C.SIGNAL],
  ];
}

function appendPriceHistoryRows(ss, rows) {
  if (!rows || rows.length === 0) return;
  const sheet = getSheetByNameSafe(ss, CONFIG.PRICE_HISTORY_SHEET);
  if (!sheet) return;
  const startRow = sheet.getLastRow() + 1;
  const neededRows = startRow + rows.length - 1;
  if (sheet.getMaxRows() < neededRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
  }
  sheet.getRange(startRow, 1, rows.length, H._COUNT).setValues(rows);
}

function buildBackfilledStockKey(scannerName, stockKey) {
  return String(scannerName || "").trim() + "||" + String(stockKey || "").trim();
}

function buildHistoryRowKey(scannerName, stockKey, snapshotTime) {
  return buildBackfilledStockKey(scannerName, stockKey) + "||" + historyTimeMs(snapshotTime);
}

function historyTimeMs(value) {
  if (value instanceof Date) return value.getTime();
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function loadPriceHistoryBackfillState(sheet) {
  const state = {
    backfilledStocks: new Set(),
    existingRowKeys: new Set(),
  };
  if (!sheet || sheet.getLastRow() < 2) return state;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  rows.forEach(function(row) {
    const snapshotTime = row[H.SNAPSHOT_AT];
    const runMode = String(row[H.RUN_MODE] || "").trim();
    const scanner = String(row[H.SCANNER] || "").trim();
    const stockKey = String(row[H.STOCK_KEY] || "").trim();
    if (!scanner || !stockKey) return;

    const rowKey = buildHistoryRowKey(scanner, stockKey, snapshotTime);
    if (historyTimeMs(snapshotTime) > 0) state.existingRowKeys.add(rowKey);
    if (runMode.indexOf("BACKFILL") === 0) {
      state.backfilledStocks.add(buildBackfilledStockKey(scanner, stockKey));
    }
  });
  return state;
}

function parseStoredISTDate(value) {
  if (!value) return null;
  if (value instanceof Date) return toISTDate(value);

  const raw = String(value).trim();
  if (!raw) return null;

  const direct = new Date(raw);
  if (!isNaN(direct.getTime())) return toISTDate(direct);

  const match = raw.match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})(?:,\s*(\d{1,2}):(\d{2})(?:\s*([APap][Mm]))?)?/);
  if (!match) return null;

  const months = {
    jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
    jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12",
  };
  const dd = pad2(Number(match[1]));
  const mm = months[String(match[2]).toLowerCase()];
  const yyyy = match[3];
  if (!mm) return null;

  let hh = match[4] ? Number(match[4]) : 0;
  const min = match[5] ? Number(match[5]) : 0;
  const meridiem = match[6] ? String(match[6]).toLowerCase() : "";
  if (meridiem === "pm" && hh < 12) hh += 12;
  if (meridiem === "am" && hh === 12) hh = 0;

  return buildISTDate(yyyy + "-" + mm + "-" + dd, hh, min);
}

function buildBackfillSnapshotTime(timestampSec) {
  const ts = Number(timestampSec);
  if (!ts) return null;
  const ist = toISTDate(new Date(ts * 1000));
  return buildISTDate(todayISTString(ist), CONFIG.MARKET_END_HOUR, CONFIG.MARKET_END_MINUTE);
}

function periodReturnAtIndex(closes, endIdx, periods) {
  if (endIdx < periods) return null;
  const past = closes[endIdx - periods];
  const cur  = closes[endIdx];
  if (isNaN(past) || isNaN(cur) || past === 0) return null;
  return ((cur - past) / past) * 100;
}

function computeBackfillLatestMetrics(hist, endIdx) {
  const closes = hist.closes.slice(0, endIdx + 1);
  const highs = hist.highs.slice(0, endIdx + 1);
  const lows = hist.lows.slice(0, endIdx + 1);
  const volumes = hist.volumes.slice(0, endIdx + 1);
  const timestamps = hist.timestamps.slice(0, endIdx + 1);

  const ma20  = calcMA(closes, 20);
  const ma50  = calcMA(closes, 50);
  const ma200 = calcMA(closes, 200);
  const rsi   = calcRSI(closes, 14);
  const adx   = calcADX(highs, lows, closes, 14);
  const volRatio = calcVolumeRatio(volumes, 20, timestamps);
  const macd  = calcMACD(closes, 12, 26, 9);
  const dist52wHigh = calcDistanceFromHigh(highs, closes, 252);
  const breakout20dPct = calcBreakoutPct(highs, closes, 20);
  const current = closes[closes.length - 1];

  return {
    rsi: rsi,
    adx: adx,
    volRatio: volRatio,
    macdLine: macd.macdLine,
    macdHist: macd.histogram,
    dist52wHigh: dist52wHigh,
    breakout20dPct: breakout20dPct,
    signal: calcSignal({
      price: current,
      ma20: ma20,
      ma50: ma50,
      ma200: ma200,
      rsi: rsi,
      adx: adx,
      volRatio: volRatio,
      macdLine: macd.macdLine,
      macdHist: macd.histogram,
      dist52wHigh: dist52wHigh,
      breakout20dPct: breakout20dPct,
    }),
  };
}

function buildBackfillRowsForStock(scannerName, row, hist, existingRowKeys) {
  const symbol = String(row[C.SYMBOL] || "").trim();
  const name = String(row[C.NAME] || "").trim();
  const stockKey = buildStockKey(symbol, name);
  if (!stockKey || !hist || !hist.closes || hist.closes.length === 0) {
    return { stockKey: stockKey, rows: [] };
  }

  const firstCaptured = parseStoredISTDate(row[C.FIRST_CAPTURED]);
  const firstCapturedDay = firstCaptured ? todayISTString(firstCaptured) : "";

  let lastIncludedIdx = -1;
  for (let i = hist.closes.length - 1; i >= 0; i--) {
    const current = Number(hist.closes[i]);
    if (!(current > 0)) continue;
    const snapshotTime = buildBackfillSnapshotTime(hist.timestamps[i]);
    if (!snapshotTime) continue;
    const snapshotDay = todayISTString(snapshotTime);
    if (firstCapturedDay && snapshotDay < firstCapturedDay) continue;
    lastIncludedIdx = i;
    break;
  }
  if (lastIncludedIdx === -1) {
    return { stockKey: stockKey, rows: [] };
  }

  let capturePrice = Number(row[C.CAPTURE_PRICE]);
  if (!(capturePrice > 0)) capturePrice = null;

  const latestMetrics = computeBackfillLatestMetrics(hist, lastIncludedIdx);
  const rows = [];

  for (let i = 0; i <= lastIncludedIdx; i++) {
    const current = Number(hist.closes[i]);
    if (!(current > 0)) continue;

    const snapshotTime = buildBackfillSnapshotTime(hist.timestamps[i]);
    if (!snapshotTime) continue;

    const snapshotDay = todayISTString(snapshotTime);
    if (firstCapturedDay && snapshotDay < firstCapturedDay) continue;

    if (!(capturePrice > 0)) capturePrice = current;

    const rowKey = buildHistoryRowKey(scannerName, stockKey, snapshotTime);
    if (existingRowKeys && existingRowKeys.has(rowKey)) continue;

    const isLatestPoint = i === lastIncludedIdx;
    rows.push([
      snapshotTime,
      "BACKFILL",
      scannerName || "",
      stockKey,
      symbol,
      name,
      row[C.IN_SCREENER],
      roundN(capturePrice, 2),
      roundN(current, 2),
      pct(capturePrice > 0 ? ((current - capturePrice) / capturePrice) * 100 : null),
      pct(periodReturnAtIndex(hist.closes, i, 1)),
      pct(periodReturnAtIndex(hist.closes, i, 5)),
      pct(periodReturnAtIndex(hist.closes, i, 21)),
      pct(periodReturnAtIndex(hist.closes, i, 63)),
      pct(periodReturnAtIndex(hist.closes, i, 126)),
      pct(periodReturnAtIndex(hist.closes, i, 252)),
      isLatestPoint ? roundN(latestMetrics.rsi, 1) : "",
      isLatestPoint ? roundN(latestMetrics.adx, 1) : "",
      isLatestPoint ? roundN(latestMetrics.volRatio, 2) : "",
      isLatestPoint ? roundN(latestMetrics.macdLine, 2) : "",
      isLatestPoint ? roundN(latestMetrics.macdHist, 2) : "",
      isLatestPoint ? pct(latestMetrics.dist52wHigh) : "",
      isLatestPoint ? pct(latestMetrics.breakout20dPct) : "",
      isLatestPoint ? latestMetrics.signal : "",
    ]);

    if (existingRowKeys) existingRowKeys.add(rowKey);
  }

  return { stockKey: stockKey, rows: rows };
}

function buildEmptyRow(stock, resolvedSymbol, dateStr) {
  const row = new Array(C._COUNT).fill("");
  row[C.SYMBOL]         = resolvedSymbol || "";
  row[C.NAME]           = stock.name     || "";
  row[C.FIRST_CAPTURED] = dateStr;
  row[C.LAST_SEEN]      = dateStr;
  row[C.IN_SCREENER]    = "✅";
  // CAPTURE_PRICE left blank — set on first successful Yahoo Finance fetch
  return row;
}

/**
 * Load all tracked stocks from a sheet into a Map.
 * Keys: both the symbol AND the name (so lookup works regardless of which
 * identifier the screener result provides).
 * FIX: v3 only stored sym||name as ONE key; now both keys point to same entry.
 */
function loadExistingStocks(sheet, lastRow) {
  const map = new Map();
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, C._COUNT).getValues();
  data.forEach((row, idx) => {
    const sym  = String(row[C.SYMBOL] || "").trim();
    const name = String(row[C.NAME]   || "").trim();
    const entry = { rowIdx: idx + 2, symbol: sym, name };
    if (sym  && sym  !== NO_SYMBOL_SENTINEL) map.set(sym,  entry);
    if (name)                                map.set(name, entry);
  });
  return map;
}

/**
 * Find a screener stock in the existing map.
 * FIX: try symbol lookup first, then name — handles both ticker and name-keyed rows.
 * Returns the matching map key, or null if not found.
 */
function findExistingKey(existingMap, stock) {
  if (stock.symbol && existingMap.has(stock.symbol)) return stock.symbol;
  if (stock.name   && existingMap.has(stock.name))   return stock.name;
  return null;
}

// ═════════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═════════════════════════════════════════════════════════════════════════════
function updateDashboard(ss) {
  let dash = getSheetByNameSafe(ss, CONFIG.DASHBOARD_SHEET);
  const previousScanner = dash ? String(dash.getRange(DASH_UI.SCANNER_CELL).getDisplayValue() || "").trim() : "";
  const previousSymbol = dash ? String(dash.getRange(DASH_UI.SYMBOL_CELL).getDisplayValue() || "").trim() : "";
  const historySheet = getSheetByNameSafe(ss, CONFIG.PRICE_HISTORY_SHEET);
  const historyEndRow = historySheet ? Math.max(historySheet.getLastRow(), 2) : 2;
  if (!dash) {
    dash = ss.insertSheet(CONFIG.DASHBOARD_SHEET, 0);
    dash.setTabColor("#4A90D9");
  }
  clearDashboardCharts(dash);
  dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns()).breakApart();
  dash.clearContents();
  dash.clearFormats();
  dash.setHiddenGridlines(true);
  dash.setFrozenRows(2);
  applyDashboardColumnLayout(dash);

  const now = formatDate(new Date());
  const DASH_COLS = DASH_UI.TOTAL_COLS;

  // Title rows
  dash.getRange(1, 1, 1, DASH_COLS).merge()
      .setValue("📊 SCREENER MULTI-MTF TRACKER")
      .setBackground("#1A237E").setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");

  dash.getRange(2, 1, 1, DASH_COLS).merge()
      .setValue(`Updated: ${now}`)
      .setBackground("#E8EAF6").setFontColor("#3949AB").setHorizontalAlignment("center");

  dash.getRange(4, 1, 1, DASH_COLS).merge()
      .setValue("📈 History Explorer")
      .setBackground("#DDEAFB").setFontWeight("bold").setFontSize(12);

  dash.getRange("A5:E5").merge()
      .setValue("Selection")
      .setBackground("#ECEFF1").setFontWeight("bold").setHorizontalAlignment("left");
  dash.getRange("G5:N5").merge()
      .setValue("Summary")
      .setBackground("#ECEFF1").setFontWeight("bold").setHorizontalAlignment("left");

  dash.getRange("H6:I6").merge();
  dash.getRange("H7:I7").merge();
  dash.getRange("H8:I8").merge();
  dash.getRange("K6:N6").merge();
  dash.getRange("K7:N7").merge();
  dash.getRange("K8:N8").merge();

  dash.getRange("A6:A8").setValues([["Scanner"], ["Stock"], ["History Points"]])
      .setBackground("#F8FAFC").setFontWeight("bold").setFontColor("#455A64");
  dash.getRange("G6:G8").setValues([["Latest Price"], ["Since Capture"], ["Latest Signal"]])
      .setBackground("#F8FAFC").setFontWeight("bold").setFontColor("#455A64");
  dash.getRange("J6:J8").setValues([["First Snapshot"], ["Latest Snapshot"], ["Tips"]])
      .setBackground("#F8FAFC").setFontWeight("bold").setFontColor("#455A64");

  dash.getRange("B6:B8").setBackground("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("left");
  dash.getRange("H6:I8").setBackground("#FFFFFF");
  dash.getRange("K6:N8").setBackground("#FFFFFF");
  dash.getRange("K8:N8").setValue("Change the dropdowns to refresh charts")
      .setFontColor("#546E7A").setFontStyle("italic");

  dash.getRange("A5:E8").setBorder(true, true, true, true, true, true, "#CFD8DC", SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange("G5:N8").setBorder(true, true, true, true, true, true, "#CFD8DC", SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange("H6:I8").setFontWeight("bold");
  dash.getRange("K6:N7").setFontWeight("bold");
  dash.getRange("H6:I6").setNumberFormat("₹#,##0.00");
  dash.getRange("H7:I7").setNumberFormat("0.00");
  dash.getRange("K6:N7").setNumberFormat("dd mmm yyyy hh:mm");
  dash.setRowHeights(1, 2, 24);
  dash.setRowHeight(4, 24);
  dash.setRowHeights(5, 4, 24);

  let dashRow = DASH_UI.TABLE_START_ROW;
  const dashHeaders = [
    "Symbol","Name","In Screener?","Price ₹","Signal","Trend",
    "ADX","Vol x","MACD Hist","52W High %",
    "1D %","1W %","1M %","3M %",
  ];

  CONFIG.SCANNERS.forEach(scanner => {
    const scanSheet = getSheetByNameSafe(ss, scanner.name);
    if (!scanSheet) return;

    const sLast  = scanSheet.getLastRow();
    const tracked = Math.max(0, sLast - 1);

    // Count "in screener" without individual row reads
    const inNow = sLast >= 2
      ? scanSheet.getRange(2, C.IN_SCREENER + 1, tracked, 1)
                 .getValues().filter(r => r[0] === "✅").length
      : 0;

    // Scanner section header
    dash.getRange(dashRow, 1, 1, DASH_COLS).merge()
        .setValue(`🔍 ${scanner.name}  |  ${inNow} in screener now  |  ${tracked} total tracked`)
        .setBackground(scanner.color || "#E8EAF6").setFontWeight("bold").setFontSize(11);
    dashRow++;

    // Sub-header row
    dash.getRange(dashRow, 1, 1, DASH_COLS).setValues([dashHeaders])
        .setBackground("#C5CAE9").setFontWeight("bold").setWrap(true).setVerticalAlignment("middle");
    dashRow++;

    if (sLast < 2) {
      dash.getRange(dashRow, 1).setValue("(no stocks yet)");
      dashRow += 2;
      return;
    }

    const data = scanSheet.getRange(2, 1, tracked, C._COUNT).getValues();
    // Sort: ✅ first, ⬜ second
    data.sort((a, b) => (a[C.IN_SCREENER] === "✅" ? 0 : 1) - (b[C.IN_SCREENER] === "✅" ? 0 : 1));

    // Build value grid and format grids in memory, write in bulk
    const valueGrid = [];
    const bgGrid    = [];     // per-cell backgrounds (14 cols)
    const fontGrid  = [];     // per-cell font colors
    const sparklineFormulas = [];

    data.forEach(r => {
      const adx = numOrBlank(r[C.ADX14]);
      const volRatio = numOrBlank(r[C.VOL_RATIO20]);
      const macdHist = numOrBlank(r[C.MACD_HIST]);
      const dist52w = numOrBlank(r[C.DIST_52W_HIGH]);
      const ret1d = numOrBlank(r[C.RET_1D]);
      const ret1w = numOrBlank(r[C.RET_1W]);
      const ret1m = numOrBlank(r[C.RET_1M]);
      const ret3m = numOrBlank(r[C.RET_3M]);
      const sig   = String(r[C.SIGNAL] || "");
      const [sigBg, sigFont] = signalColor(sig);

      valueGrid.push([
        r[C.SYMBOL], r[C.NAME], r[C.IN_SCREENER],
        r[C.CURRENT_PRICE], sig, "",
        adx, volRatio, macdHist, dist52w,
        ret1d, ret1w, ret1m, ret3m,
      ]);

      bgGrid.push([
        null, null, null,
        null, sigBg, null,
        adxBg(adx), volumeBg(volRatio), retBg(macdHist), proximityBg(dist52w),
        retBg(ret1d), retBg(ret1w), retBg(ret1m), retBg(ret3m),
      ]);

      fontGrid.push([
        null, null, null,
        null, sigFont, null,
        adxFont(adx), volumeFont(volRatio), retFont(macdHist), proximityFont(dist52w),
        retFont(ret1d), retFont(ret1w), retFont(ret1m), retFont(ret3m),
      ]);

      sparklineFormulas.push([
        buildDashboardSparklineFormula(scanner.name, r, historyEndRow)
      ]);
    });

    if (valueGrid.length > 0) {
      const dataRange = dash.getRange(dashRow, 1, valueGrid.length, DASH_COLS);
      // FIX: 3 bulk calls instead of N×4 individual calls
      dataRange.setValues(valueGrid);
      dataRange.setBackgrounds(bgGrid);
      dataRange.setFontColors(fontGrid);
      dash.getRange(dashRow, 6, valueGrid.length, 1).setFormulas(sparklineFormulas);
      // Center the "In Screener?" column (col 3)
      dash.getRange(dashRow, 3, valueGrid.length, 1).setHorizontalAlignment("center");
      dash.getRange(dashRow, 4, valueGrid.length, 1).setNumberFormat("₹#,##0.00");
      dash.getRange(dashRow, 7, valueGrid.length, 1).setNumberFormat("0.0");
      dash.getRange(dashRow, 8, valueGrid.length, 7).setNumberFormat("0.00");
      dashRow += valueGrid.length;
    }

    dashRow += 2;  // gap between scanners
  });

  try {
    refreshDashboardHistoryView(ss, { scanner: previousScanner, symbol: previousSymbol });
  } catch (e) {
    Logger.log("Dashboard history explorer refresh failed: " + e.message);
  }
}

function buildDashboardSparklineFormula(scannerName, row, historyEndRow) {
  const stockKey = buildStockKey(row[C.SYMBOL], row[C.NAME]);
  if (!stockKey || historyEndRow < 2) return '=""';

  const color = Number(row[C.RET_CAPTURE]) >= 0 ? "#1B5E20" : "#B71C1C";
  const priceRange   = buildSheetRangeA1(CONFIG.PRICE_HISTORY_SHEET, H.CURRENT_PRICE + 1, 2, historyEndRow);
  const scannerRange = buildSheetRangeA1(CONFIG.PRICE_HISTORY_SHEET, H.SCANNER + 1, 2, historyEndRow);
  const keyRange     = buildSheetRangeA1(CONFIG.PRICE_HISTORY_SHEET, H.STOCK_KEY + 1, 2, historyEndRow);

  return `=IFERROR(SPARKLINE(FILTER(${priceRange},${scannerRange}="${escapeFormulaString(scannerName)}",${keyRange}="${escapeFormulaString(stockKey)}"),{"charttype","line";"linewidth",2;"color","${color}"}),"")`;
}

function refreshDashboardHistoryView(ss, preferredSelection) {
  const dash = getSheetByNameSafe(ss, CONFIG.DASHBOARD_SHEET);
  const historySheet = getSheetByNameSafe(ss, CONFIG.PRICE_HISTORY_SHEET);
  const chartSheet = getSheetByNameSafe(ss, CONFIG.CHART_DATA_SHEET);
  if (!dash || !historySheet || !chartSheet) return;

  clearDashboardCharts(dash);
  clearChartDataRows(chartSheet);

  const historyLastRow = historySheet.getLastRow();
  const historyData = historyLastRow >= 2
    ? historySheet.getRange(2, 1, historyLastRow - 1, H._COUNT).getValues()
    : [];

  const scannerCell = dash.getRange(DASH_UI.SCANNER_CELL);
  const symbolCell = dash.getRange(DASH_UI.SYMBOL_CELL);
  const runsCell = dash.getRange(DASH_UI.RUNS_CELL);
  const latestPriceRange = dash.getRange("H6:I6");
  const returnRange = dash.getRange("H7:I7");
  const signalRange = dash.getRange("H8:I8");
  const firstRunRange = dash.getRange("K6:N6");
  const lastRunRange = dash.getRange("K7:N7");

  const scannerOptions = getDashboardScannerOptions(historyData);
  if (scannerOptions.length === 0) {
    scannerCell.clearContent().clearDataValidations();
    symbolCell.clearContent().clearDataValidations();
    runsCell.setValue(0);
    latestPriceRange.clearContent();
    returnRange.clearContent();
    signalRange.setValue("No history yet").setBackground(null).setFontColor("#546E7A");
    firstRunRange.clearContent();
    lastRunRange.clearContent();
    return;
  }

  let selectedScanner = String((preferredSelection && preferredSelection.scanner) || scannerCell.getDisplayValue() || "").trim();
  if (scannerOptions.indexOf(selectedScanner) === -1) {
    selectedScanner = scannerOptions[0];
  }
  scannerCell.setDataValidation(buildValueInListRule(scannerOptions)).setValue(selectedScanner);

  const stockOptions = getDashboardStockOptions(ss, historyData, selectedScanner);
  const stockLabels = stockOptions.map(function(item) { return item.display; });
  let selectedLabel = String((preferredSelection && preferredSelection.symbol) || symbolCell.getDisplayValue() || "").trim();
  if (stockLabels.indexOf(selectedLabel) === -1) {
    selectedLabel = stockLabels[0] || "";
  }
  if (stockLabels.length > 0) {
    symbolCell.setDataValidation(buildValueInListRule(stockLabels)).setValue(selectedLabel);
  } else {
    symbolCell.clearContent().clearDataValidations();
  }

  const selectedStock = stockOptions.find(function(item) { return item.display === selectedLabel; }) || null;
  const points = selectedStock
    ? historyData.filter(function(row) {
        return String(row[H.SCANNER] || "").trim() === selectedScanner && row[H.STOCK_KEY] === selectedStock.key;
      }).sort(compareHistoryRows)
    : [];

  writeDashboardHistorySummary(dash, points);
  writeChartDataRows(chartSheet, points);
  if (points.length > 0) {
    SpreadsheetApp.flush();
    insertDashboardCharts(dash, chartSheet, selectedScanner, selectedStock.display, points.length);
  }
}

function getDashboardScannerOptions(historyData) {
  const options = [];
  const seen = new Set();

  CONFIG.SCANNERS.forEach(function(scanner) {
    const name = String(scanner && scanner.name || "").trim();
    if (!name || seen.has(name)) return;
    seen.add(name);
    options.push(name);
  });

  historyData.forEach(function(row) {
    const name = String(row[H.SCANNER] || "").trim();
    if (!name || seen.has(name)) return;
    seen.add(name);
    options.push(name);
  });

  return options;
}

function getDashboardStockOptions(ss, historyData, scannerName) {
  const latestByKey = new Map();

  historyData.forEach(function(row) {
    const scanner = String(row[H.SCANNER] || "").trim();
    const key = String(row[H.STOCK_KEY] || "").trim();
    if (scanner !== scannerName || !key) return;

    const time = getHistoryTime(row);
    const existing = latestByKey.get(key);
    if (!existing || time > existing.time) {
      latestByKey.set(key, {
        key: key,
        time: time,
        symbol: String(row[H.SYMBOL] || "").trim(),
        name: String(row[H.NAME] || "").trim(),
      });
    }
  });

  const sheet = getSheetByNameSafe(ss, scannerName);
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, C._COUNT).getValues();
    data.forEach(function(row) {
      const key = buildStockKey(row[C.SYMBOL], row[C.NAME]);
      if (!key || latestByKey.has(key)) return;
      latestByKey.set(key, {
        key: key,
        time: 0,
        symbol: String(row[C.SYMBOL] || "").trim(),
        name: String(row[C.NAME] || "").trim(),
      });
    });
  }

  return Array.from(latestByKey.values())
    .sort(function(a, b) {
      return buildDashboardStockLabel(a.symbol, a.name).localeCompare(buildDashboardStockLabel(b.symbol, b.name));
    })
    .map(function(item) {
      return {
        key: item.key,
        display: buildDashboardStockLabel(item.symbol, item.name),
      };
    });
}

function buildDashboardStockLabel(symbol, name) {
  const sym = String(symbol || "").trim();
  const hasSymbol = sym && sym !== NO_SYMBOL_SENTINEL && !sym.startsWith(NO_SYMBOL_SENTINEL + "|");
  if (hasSymbol && name) return sym + " | " + name;
  return hasSymbol ? sym : String(name || "").trim();
}

function writeDashboardHistorySummary(dash, points) {
  const latest = points.length > 0 ? points[points.length - 1] : null;
  const latestPriceRange = dash.getRange("H6:I6");
  const returnRange = dash.getRange("H7:I7");
  const signalRange = dash.getRange("H8:I8");
  const firstRunRange = dash.getRange("K6:N6");
  const lastRunRange = dash.getRange("K7:N7");

  dash.getRange(DASH_UI.RUNS_CELL).setValue(points.length);
  latestPriceRange.setValue(latest ? numOrBlank(latest[H.CURRENT_PRICE]) : "");
  returnRange.setValue(latest ? numOrBlank(latest[H.RET_CAPTURE]) : "");
  firstRunRange.setValue(points.length > 0 ? points[0][H.SNAPSHOT_AT] : "");
  lastRunRange.setValue(latest ? latest[H.SNAPSHOT_AT] : "");

  if (!latest) {
    signalRange.setValue("No history yet").setBackground(null).setFontColor("#546E7A");
    return;
  }

  const signal = String(latest[H.SIGNAL] || "");
  const colors = signalColor(signal);
  signalRange.setValue(signal).setBackground(colors[0]).setFontColor(colors[1]);
}

function writeChartDataRows(sheet, points) {
  if (!points || points.length === 0) return;

  const rows = points.map(function(row) {
    return [
      row[H.SNAPSHOT_AT],
      numOrBlank(row[H.CURRENT_PRICE]),
      row[H.SNAPSHOT_AT],
      numOrBlank(row[H.RET_CAPTURE]),
      row[H.SNAPSHOT_AT],
      numOrBlank(row[H.RET_1D]),
      numOrBlank(row[H.RET_1W]),
      numOrBlank(row[H.RET_1M]),
      row[H.SIGNAL],
    ];
  });

  const neededRows = rows.length + 1;
  if (sheet.getMaxRows() < neededRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
  }
  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
}

function insertDashboardCharts(dash, chartSheet, scannerName, stockLabel, pointCount) {
  const titleBase = stockLabel || scannerName || "Selected Stock";
  const pointSize = pointCount <= 20 ? 4 : 1;

  const priceChart = dash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartSheet.getRange(1, 1, pointCount + 1, 2))
    .setPosition(DASH_UI.PRICE_CHART_ROW, DASH_UI.PRICE_CHART_COL, 0, 0)
    .setOption("title", titleBase + " — Price History")
    .setOption("legend", { position: "none" })
    .setOption("lineWidth", 2)
    .setOption("pointSize", pointSize)
    .setOption("width", 1260)
    .setOption("height", 250)
    .setOption("hAxis", { title: "Run Time", slantedText: true, slantedTextAngle: 35 })
    .setOption("vAxis", { title: "Price ₹" })
    .build();
  dash.insertChart(priceChart);

  const returnChart = dash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartSheet.getRange(1, 3, pointCount + 1, 2))
    .setPosition(DASH_UI.RETURN_CHART_ROW, DASH_UI.RETURN_CHART_COL, 0, 0)
    .setOption("title", titleBase + " — Since Capture %")
    .setOption("legend", { position: "none" })
    .setOption("lineWidth", 2)
    .setOption("pointSize", pointSize)
    .setOption("width", 620)
    .setOption("height", 220)
    .setOption("hAxis", { title: "Run Time", slantedText: true, slantedTextAngle: 35 })
    .setOption("vAxis", { title: "Return %" })
    .build();
  dash.insertChart(returnChart);

  const comparisonChart = dash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartSheet.getRange(1, 5, pointCount + 1, 4))
    .setPosition(DASH_UI.MULTI_RETURN_CHART_ROW, DASH_UI.MULTI_RETURN_CHART_COL, 0, 0)
    .setOption("title", titleBase + " — Short-Term Returns")
    .setOption("legend", { position: "bottom" })
    .setOption("lineWidth", 2)
    .setOption("pointSize", pointSize)
    .setOption("width", 620)
    .setOption("height", 220)
    .setOption("hAxis", { title: "Run Time", slantedText: true, slantedTextAngle: 35 })
    .setOption("vAxis", { title: "Return %" })
    .setOption("series", {
      0: { color: "#1565C0" },
      1: { color: "#2E7D32" },
      2: { color: "#EF6C00" },
    })
    .build();
  dash.insertChart(comparisonChart);
}

function clearDashboardCharts(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
}

function clearChartDataRows(sheet) {
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).clearContent();
  }
}

function applyDashboardColumnLayout(sheet) {
  DASHBOARD_COL_WIDTHS.forEach(function(width, idx) {
    sheet.setColumnWidth(idx + 1, width);
  });
}

function buildValueInListRule(values) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
}

function uniqueStrings(values) {
  return Array.from(new Set(values.filter(function(value) { return value; })));
}

function compareHistoryRows(a, b) {
  return getHistoryTime(a) - getHistoryTime(b);
}

function getHistoryTime(row) {
  const value = row[H.SNAPSHOT_AT];
  if (value instanceof Date) return value.getTime();
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (normalizeSheetName(sheet.getName()) !== normalizeSheetName(CONFIG.DASHBOARD_SHEET)) return;
    const a1 = e.range.getA1Notation();
    if (a1 !== DASH_UI.SCANNER_CELL && a1 !== DASH_UI.SYMBOL_CELL) return;
    refreshDashboardHistoryView(e.source || SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    Logger.log("Dashboard history refresh failed: " + err.message);
  }
}

function refreshDashboardHistory() {
  refreshDashboardHistoryView(SpreadsheetApp.getActiveSpreadsheet());
}

function rebuildDashboardOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSystemSheets(ss);
  updateDashboard(ss);
}

/** Background color for a return cell */
function retBg(val)   { return (val > 0) ? "#E8F5E9" : (val < 0) ? "#FFEBEE" : null; }
/** Font color for a return cell */
function retFont(val) { return (val > 0) ? "#1B5E20" : (val < 0) ? "#B71C1C" : null; }
/** Background color for ADX strength */
function adxBg(val) { return (val >= 25) ? "#C8E6C9" : (val >= 20) ? "#E8F5E9" : (val !== "" && val < 15) ? "#FFEBEE" : null; }
/** Font color for ADX strength */
function adxFont(val) { return (val >= 20) ? "#1B5E20" : (val !== "" && val < 15) ? "#C62828" : null; }
/** Background color for participation / volume */
function volumeBg(val) { return (val >= 1.5) ? "#E8F5E9" : (val !== "" && val < 0.8) ? "#FFF8E1" : null; }
/** Font color for participation / volume */
function volumeFont(val) { return (val >= 1.5) ? "#1B5E20" : (val !== "" && val < 0.8) ? "#BF360C" : null; }
/** Background color when the stock is close to / far from 52W high */
function proximityBg(val) { return (val !== "" && val <= 2) ? "#E8F5E9" : (val >= 10) ? "#FFF3E0" : null; }
/** Font color when the stock is close to / far from 52W high */
function proximityFont(val) { return (val !== "" && val <= 2) ? "#1B5E20" : (val >= 10) ? "#E65100" : null; }
/** Parse a cell value to number, or return blank if the source is blank/NaN */
function numOrBlank(v) { const n = Number(v); return isNaN(n) ? "" : n; }

// ═════════════════════════════════════════════════════════════════════════════
// EMAIL ALERT
// ═════════════════════════════════════════════════════════════════════════════
function sendEmailAlert(ss, allNewStocks) {
  const total    = allNewStocks.reduce((a, b) => a + b.stocks.length, 0);
  const sheetUrl = ss.getUrl();
  const lines    = [];

  allNewStocks.forEach(g => {
    lines.push(`\n🔍 ${g.scanner}:`);
    g.stocks.forEach(s => lines.push(`   • ${s.symbol || ""}  ${s.name || ""}`));
  });

  MailApp.sendEmail({
    to:      CONFIG.ALERT_EMAIL,
    subject: `📈 ${total} new stock(s) in Screener — ${formatDate(new Date())}`,
    body: [
      "New stocks appeared in your scanners:",
      ...lines,
      `\nTracker: ${sheetUrl}`,
      "\n(Performance data + signals auto-update on next run)",
    ].join("\n"),
  });
}

// ═════════════════════════════════════════════════════════════════════════════
// SYSTEM SHEETS + LOGGING
// ═════════════════════════════════════════════════════════════════════════════
function ensureSystemSheets(ss) {
  let logSheet = getSheetByNameSafe(ss, CONFIG.LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET);
    logSheet.setTabColor("#757575");
    logSheet.appendRow(["Timestamp", "Message"]);
    logSheet.getRange(1, 1, 1, 2).setBackground("#424242").setFontColor("#FFFFFF").setFontWeight("bold");
    logSheet.setFrozenRows(1);
  }

  let historySheet = getSheetByNameSafe(ss, CONFIG.PRICE_HISTORY_SHEET);
  if (!historySheet) {
    historySheet = ss.insertSheet(CONFIG.PRICE_HISTORY_SHEET);
    historySheet.setTabColor("#00695C");
  }
  ensurePriceHistorySheetSchema(historySheet);

  let chartSheet = getSheetByNameSafe(ss, CONFIG.CHART_DATA_SHEET);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(CONFIG.CHART_DATA_SHEET);
    chartSheet.setTabColor("#455A64");
  }
  ensureChartDataSheetSchema(chartSheet);
  if (!chartSheet.isSheetHidden() && ss.getActiveSheet().getSheetId() !== chartSheet.getSheetId()) {
    chartSheet.hideSheet();
  }
}

// FIX: accepts `ss` to avoid calling getActiveSpreadsheet() on every log call
function log(ss, msg) {
  try {
    const sheet = getSheetByNameSafe(ss, CONFIG.LOG_SHEET);
    if (!sheet) return;
    sheet.appendRow([formatDate(new Date()), msg]);
    const rows = sheet.getLastRow();
    if (rows > CONFIG.MAX_LOG_ROWS + 1) {
      sheet.deleteRows(2, rows - CONFIG.MAX_LOG_ROWS - 1);
    }
  } catch (_) { /* log failure must never crash the main flow */ }
}

// ═════════════════════════════════════════════════════════════════════════════
// UTILITIES
// ═════════════════════════════════════════════════════════════════════════════

function getISTNow() {
  return toISTDate(new Date());
}

function toISTDate(inputDate) {
  return new Date(inputDate.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));
}

function buildRunContext(mode) {
  const snapshotTime = toISTDate(new Date());
  return {
    mode: String(mode || "AUTO"),
    snapshotTime: snapshotTime,
  };
}

function escapeFormulaString(value) {
  return String(value || "").replace(/"/g, '""');
}

function buildSheetRangeA1(sheetName, colNum, rowStart, rowEnd) {
  const col = columnToLetter(colNum);
  const safeSheet = "'" + String(sheetName || "").replace(/'/g, "''") + "'";
  return safeSheet + "!$" + col + "$" + rowStart + ":$" + col + "$" + rowEnd;
}

function columnToLetter(colNum) {
  let n = colNum;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function getClockWindowLabel() {
  return pad2(CONFIG.MARKET_START_HOUR) + ":" + pad2(CONFIG.MARKET_START_MINUTE) +
         " - " +
         pad2(CONFIG.MARKET_END_HOUR) + ":" + pad2(CONFIG.MARKET_END_MINUTE);
}

function getConfiguredDateRanges() {
  const ranges = Array.isArray(CONFIG.ACTIVE_DATE_RANGES) ? CONFIG.ACTIVE_DATE_RANGES : [];
  return ranges.map(function(range, idx) {
    const from = String(range && range.from || "").trim();
    const to = String(range && range.to || "").trim();
    const label = String(range && range.label || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      throw new Error("Invalid ACTIVE_DATE_RANGES entry #" + (idx + 1) + " — use YYYY-MM-DD for both from and to");
    }
    if (from > to) {
      throw new Error("Invalid ACTIVE_DATE_RANGES entry #" + (idx + 1) + " — from date must be <= to date");
    }
    return { from: from, to: to, label: label };
  });
}

function getDateFilterStatus(ist) {
  const dateStr = todayISTString(ist);
  const ranges = getConfiguredDateRanges();

  if (ranges.length > 0) {
    const matched = ranges.find(function(range) {
      return dateStr >= range.from && dateStr <= range.to;
    });
    if (matched) {
      const suffix = matched.label ? " (" + matched.label + ")" : "";
      return {
        allowed: true,
        mode: "date_ranges",
        reason: "inside configured date range " + matched.from + " to " + matched.to + suffix,
      };
    }
    return {
      allowed: false,
      mode: "date_ranges",
      reason: "outside configured date ranges (" + dateStr + ")",
    };
  }

  const day = ist.getDay();   // 0 = Sun, 6 = Sat
  if (day === 0 || day === 6) {
    return {
      allowed: false,
      mode: "weekdays",
      reason: "outside weekday schedule (" + dateStr + ")",
    };
  }

  return {
    allowed: true,
    mode: "weekdays",
    reason: "inside weekday schedule",
  };
}

function isWithinConfiguredTimeWindow(ist) {
  const nowMins   = ist.getHours() * 60 + ist.getMinutes();
  const startMins = CONFIG.MARKET_START_HOUR * 60 + CONFIG.MARKET_START_MINUTE;
  const endMins   = CONFIG.MARKET_END_HOUR * 60 + CONFIG.MARKET_END_MINUTE;
  return nowMins >= startMins && nowMins <= endMins;
}

/** Actual exchange-style live session check used for intraday volume handling. */
function isLiveExchangeSessionNow() {
  const ist = getISTNow();
  const day = ist.getDay();
  if (day === 0 || day === 6) return false;
  return isWithinConfiguredTimeWindow(ist);
}

function isMinuteAlignedToTriggerSlot(ist) {
  const nowMins = ist.getHours() * 60 + ist.getMinutes();
  const startMins = CONFIG.MARKET_START_HOUR * 60 + CONFIG.MARKET_START_MINUTE;
  const interval = CONFIG.TRIGGER_INTERVAL_MINUTES;
  const mod = ((nowMins - startMins) % interval + interval) % interval;
  return mod === 0;
}

function buildISTDate(dateStr, hour, minute) {
  return new Date(dateStr + "T" + pad2(hour) + ":" + pad2(minute) + ":00+05:30");
}

function getNextRunTime(afterDate) {
  const interval = CONFIG.TRIGGER_INTERVAL_MINUTES;
  const validIntervals = [1, 5, 10, 15, 30];
  if (validIntervals.indexOf(interval) === -1) {
    throw new Error("CONFIG.TRIGGER_INTERVAL_MINUTES must be one of: " + validIntervals.join(", "));
  }

  const ranges = getConfiguredDateRanges();
  const now = afterDate || new Date();
  const nowIst = toISTDate(now);
  const today = todayISTString(nowIst);
  let candidate = new Date(now.getTime() + 60 * 1000);

  if (ranges.length > 0) {
    const futureRanges = ranges.slice().sort(function(a, b) {
      return a.from.localeCompare(b.from);
    }).filter(function(range) {
      return range.to >= today;
    });
    if (futureRanges.length === 0) return null;

    if (today < futureRanges[0].from) {
      candidate = buildISTDate(futureRanges[0].from, CONFIG.MARKET_START_HOUR, CONFIG.MARKET_START_MINUTE);
    }
  }

  let aligned = toISTDate(candidate);
  for (let i = 0; i < interval + 1 && !isMinuteAlignedToTriggerSlot(aligned); i++) {
    candidate = new Date(candidate.getTime() + 60 * 1000);
    aligned = toISTDate(candidate);
  }

  const maxChecks = Math.ceil((366 * 24 * 60) / interval) + 10;
  for (let i = 0; i < maxChecks; i++) {
    const ist = toISTDate(candidate);
    if (getDateFilterStatus(ist).allowed && isWithinConfiguredTimeWindow(ist) && isMinuteAlignedToTriggerSlot(ist)) {
      return candidate;
    }
    candidate = new Date(candidate.getTime() + interval * 60 * 1000);
  }

  return null;
}

function ensureNextRunTrigger(ss) {
  deleteTriggersByHandler("runAllScanners");
  const nextRun = getNextRunTime(new Date());
  if (!nextRun) {
    const msg = "⚠️ No future valid run slot found — runAllScanners trigger not created";
    if (ss) log(ss, msg);
    Logger.log(msg);
    return null;
  }
  ScriptApp.newTrigger("runAllScanners").timeBased().at(nextRun).create();
  Logger.log("✅ Next runAllScanners trigger scheduled for " + formatDate(nextRun) + " IST");
  return nextRun;
}

function getRunScheduleStatus() {
  const ist = getISTNow();
  const dateFilter = getDateFilterStatus(ist);

  if (!dateFilter.allowed) {
    return {
      isLiveWindow: false,
      canUseOffHours: false,
      reason: dateFilter.reason,
      mode: dateFilter.mode,
    };
  }

  if (!isWithinConfiguredTimeWindow(ist)) {
    return {
      isLiveWindow: false,
      canUseOffHours: false,
      reason: "outside intraday time window " + getClockWindowLabel() + " IST",
      mode: dateFilter.mode,
    };
  }

  return {
    isLiveWindow: true,
    canUseOffHours: false,
    reason: "inside live window",
    mode: dateFilter.mode,
  };
}

function getScheduleSummaryLines() {
  const lines = [];
  const ranges = getConfiguredDateRanges();

  if (ranges.length > 0) {
    lines.push("ℹ️  Date filter mode: explicit date ranges");
    ranges.forEach(function(range, idx) {
      const suffix = range.label ? " (" + range.label + ")" : "";
      lines.push("ℹ️  Date range " + (idx + 1) + ": " + range.from + " to " + range.to + suffix);
    });
  } else {
    lines.push("ℹ️  Date filter mode: weekdays only (Mon-Fri)");
  }

  lines.push("ℹ️  Live intraday window: " + getClockWindowLabel() + " IST");
  return lines;
}

function isMarketHours() {
  return getRunScheduleStatus().isLiveWindow;
}

/** Returns today's date in IST as "YYYY-MM-DD" */
function todayISTString(inputDate) {
  const d = inputDate || getISTNow();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

/** True if the run has consumed more than CONFIG.MAX_RUN_MS milliseconds */
function isTimedOut(runStart) {
  return (Date.now() - runStart) > CONFIG.MAX_RUN_MS;
}

function msToSec(ms) { return (ms / 1000).toFixed(1); }

function pad2(n) { return String(n).padStart(2, "0"); }

/** Round to N decimal places; return "" for non-numeric values */
function roundN(val, dec) {
  if (val == null) return "";
  const n = Number(val);
  if (isNaN(n)) return "";
  return Math.round(n * Math.pow(10, dec)) / Math.pow(10, dec);
}

/** Store percentage as raw number (2 dp). Blank for null/NaN. */
function pct(val) {
  if (val == null) return "";
  const n = Number(val);
  return isNaN(n) ? "" : Math.round(n * 100) / 100;
}

function formatDate(d) {
  return d.toLocaleString("en-IN", {
    timeZone:   "Asia/Kolkata",
    day: "2-digit", month: "short", year: "numeric",
    hour: "2-digit", minute: "2-digit",
  });
}

/** Delete only triggers that point to the given handler. */
function deleteTriggersByHandler(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function hasTriggerForHandler(handlerName) {
  return ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handlerName;
  });
}

function getAutoBackfillStatus() {
  return PropertiesService.getScriptProperties().getProperty(AUTO_BACKFILL_STATUS_KEY) || "";
}

function setAutoBackfillStatus(status) {
  const props = PropertiesService.getScriptProperties();
  if (status) {
    props.setProperty(AUTO_BACKFILL_STATUS_KEY, status);
  } else {
    props.deleteProperty(AUTO_BACKFILL_STATUS_KEY);
  }
}

function hasTrackedStocksForBackfill(ss) {
  if (!ss) return false;
  return CONFIG.SCANNERS.some(function(scanner) {
    const sheet = getSheetByNameSafe(ss, scanner.name);
    return sheet && sheet.getLastRow() >= 2;
  });
}

function ensureBackfillTriggerScheduled(ss, reason) {
  deleteTriggersByHandler("backfillPriceHistoryFromYahoo");
  const runAt = new Date(Date.now() + AUTO_BACKFILL_TRIGGER_DELAY_MS);
  ScriptApp.newTrigger("backfillPriceHistoryFromYahoo").timeBased().at(runAt).create();

  const message = "🕰️ Price history backfill scheduled for " + formatDate(runAt) + " IST" +
    (reason ? " — " + reason : "");
  if (ss) log(ss, message);
  Logger.log(message);
  return runAt;
}

function maybeScheduleAutomaticBackfill(ss) {
  if (getAutoBackfillStatus() !== AUTO_BACKFILL_PENDING) return null;
  if (!hasTrackedStocksForBackfill(ss)) return null;
  if (hasTriggerForHandler("backfillPriceHistoryFromYahoo")) return null;
  return ensureBackfillTriggerScheduled(ss, "automatic history backfill");
}

// ═════════════════════════════════════════════════════════════════════════════
// SETUP  (run ONCE manually from the Apps Script editor)
// ═════════════════════════════════════════════════════════════════════════════
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  rebuildDashboardOnly();
  const nextRun = ensureNextRunTrigger(null);
  setAutoBackfillStatus(AUTO_BACKFILL_PENDING);
  const backfillRun = maybeScheduleAutomaticBackfill(ss);
  Logger.log("ℹ️  Trigger model: exact next-slot scheduling (no automatic off-hours triggers)");
  Logger.log("ℹ️  rebuildDashboardOnly() runs during setup and scheduled runs keep the dashboard updated");
  getScheduleSummaryLines().forEach(function(line) { Logger.log(line); });
  Logger.log(`ℹ️  Signal profile: ${getSignalProfileName()}`);
  if (nextRun) {
    Logger.log("ℹ️  First scheduled run: " + formatDate(nextRun) + " IST");
  } else {
    Logger.log("⚠️  No future run was scheduled — check your date filter configuration");
  }
  if (backfillRun) {
    Logger.log("ℹ️  Automatic price-history backfill scheduled for " + formatDate(backfillRun) + " IST");
  } else {
    Logger.log("ℹ️  Automatic price-history backfill armed — it will schedule itself after stock data exists");
  }
  Logger.log("ℹ️  Stocks persist in the sheet even after leaving the screener");
  Logger.log("ℹ️  Run testYahooFetch() first to confirm Yahoo Finance connectivity");
}

// ═════════════════════════════════════════════════════════════════════════════
// MANUAL TEST FUNCTIONS
// ═════════════════════════════════════════════════════════════════════════════

/** Full test run on SCANNERS[0] — ignores market hours check, no trigger, no email */
function testRunFirstScanner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSystemSheets(ss);
  const RUN_START = Date.now();
  const runContext = buildRunContext("TEST");
  log(ss, "🧪 TEST START");
  processScanner(ss, CONFIG.SCANNERS[0], RUN_START, runContext);
  updateDashboard(ss);
  maybeScheduleAutomaticBackfill(ss);
  log(ss, `🧪 TEST DONE — ${msToSec(Date.now() - RUN_START)}s`);
}

/** Manual full run across all scanners — ignores schedule, no trigger, no email */
function testRunAllScannersManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSystemSheets(ss);

  const RUN_START = Date.now();
  const runContext = buildRunContext("MANUAL");
  const allNewStocks = [];
  log(ss, "🧪 MANUAL FULL TEST START");

  CONFIG.SCANNERS.forEach(scanner => {
    if (isTimedOut(RUN_START)) {
      log(ss, "⏱ Manual test stopped — global timeout reached");
      return;
    }
    try {
      log(ss, `🧪 Running manually: ${scanner.name}`);
      const newOnes = processScanner(ss, scanner, RUN_START, runContext);
      if (newOnes.length > 0) allNewStocks.push({ scanner: scanner.name, stocks: newOnes });
    } catch (e) {
      log(ss, `❌ [MANUAL ${scanner.name}] ${e.message}\n${e.stack || ""}`);
    }
  });

  updateDashboard(ss);
  maybeScheduleAutomaticBackfill(ss);

  const total = allNewStocks.reduce((a, b) => a + b.stocks.length, 0);
  log(ss, `🧪 MANUAL FULL TEST DONE — ${total} new stock(s) — ${msToSec(Date.now() - RUN_START)}s elapsed`);
}

/** Verify Yahoo Finance history + calculations for one symbol */
function testYahooFetch() {
  const symbol = "RELIANCE";   // ← change to test any NSE symbol
  Logger.log(`Fetching: ${symbol}${CONFIG.YF_SUFFIX}`);
  Logger.log(`Signal profile: ${getSignalProfileName()}`);

  const hist = normalizeHistory(fetchYahooHistory(symbol));
  if (!hist) { Logger.log("❌ No data — check CONFIG.YF_SUFFIX"); return; }

  const closes = hist.closes;
  const n      = closes.length;
  const cur    = closes[n - 1];
  const ma20   = calcMA(closes, 20);
  const ma50   = calcMA(closes, 50);
  const ma200  = calcMA(closes, 200);
  const rsi    = calcRSI(closes, 14);
  const adx    = calcADX(hist.highs, hist.lows, closes, 14);
  const volRatio = calcVolumeRatio(hist.volumes, 20, hist.timestamps);
  const macd   = calcMACD(closes, 12, 26, 9);
  const dist52wHigh = calcDistanceFromHigh(hist.highs, closes, 252);
  const breakout20dPct = calcBreakoutPct(hist.highs, closes, 20);

  Logger.log(`✅ ${n} trading days  |  Current: ₹${roundN(cur, 2)}`);
  Logger.log(`   1D: ${pct(periodReturn(closes, n,   1))}%  |  1W: ${pct(periodReturn(closes, n, 5))}%`);
  Logger.log(`   1M: ${pct(periodReturn(closes, n,  21))}%  |  3M: ${pct(periodReturn(closes, n, 63))}%`);
  Logger.log(`   6M: ${pct(periodReturn(closes, n, 126))}%  |  1Y: ${pct(periodReturn(closes, n, 252))}%`);
  Logger.log(`   MA20: ₹${roundN(ma20,2)}  |  MA50: ₹${roundN(ma50,2)}  |  MA200: ₹${roundN(ma200,2)}`);
  Logger.log(`   RSI(14): ${roundN(rsi, 1)}  |  ADX(14): ${roundN(adx, 1)}`);
  Logger.log(`   Vol Ratio(20): ${roundN(volRatio, 2)}x  |  MACD Line: ${roundN(macd.macdLine, 2)}  |  MACD Hist: ${roundN(macd.histogram, 2)}`);
  Logger.log(`   52W High Dist: ${pct(dist52wHigh)}%`);
  Logger.log(`   20D Breakout: ${pct(breakout20dPct)}%`);
  Logger.log(`   Signal: ${calcSignal({
    price: cur,
    ma20: ma20,
    ma50: ma50,
    ma200: ma200,
    rsi: rsi,
    adx: adx,
    volRatio: volRatio,
    macdLine: macd.macdLine,
    macdHist: macd.histogram,
    dist52wHigh: dist52wHigh,
    breakout20dPct: breakout20dPct,
  })}`);
  Logger.log(`   Avg Weekly: ${pct(avgPeriodReturn(closes, 5, 252))}%`);
  Logger.log(`   Avg Monthly: ${pct(avgPeriodReturn(closes, 21, 504))}%`);
}

/** Test Yahoo Finance symbol search by company name */
function testSymbolSearch() {
  const name = "Reliance Industries";  // ← change to test any name
  const sym  = searchYahooSymbol(name);
  Logger.log(sym ? `✅ Found: ${sym}` : "❌ Not found");
}

/** Re-run performance update only (no screener fetch) */
function testUpdatePerformanceOnly() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  ensureSystemSheets(ss);
  const sheet = getSheetByNameSafe(ss, CONFIG.SCANNERS[0].name);
  if (!sheet) { Logger.log("Sheet not found — run testRunFirstScanner first"); return; }
  const RUN_START = Date.now();
  const runContext = buildRunContext("MANUAL PERF");
  log(ss, `🔄 Performance update: [${CONFIG.SCANNERS[0].name}]`);
  updateSheetPerformance(ss, sheet, RUN_START, CONFIG.SCANNERS[0].name, runContext);
  maybeScheduleAutomaticBackfill(ss);
  log(ss, `✅ Done — ${msToSec(Date.now() - RUN_START)}s`);
}

/**
 * Seed the Price History sheet with synthetic daily EOD snapshots from Yahoo.
 * This can run manually or via the automatic one-time backfill flow armed by
 * setupTriggers(). It is not part of the normal live scanning trigger cadence.
 *
 * Notes:
 * - starts from each stock's First Captured date
 * - skips stocks already backfilled on later reruns
 * - safe to rerun if GAS times out mid-process
 * - automatically re-schedules itself when it hits the GAS time limit
 */
function backfillPriceHistoryFromYahoo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSystemSheets(ss);

  const historySheet = getSheetByNameSafe(ss, CONFIG.PRICE_HISTORY_SHEET);
  const state = loadPriceHistoryBackfillState(historySheet);
  const RUN_START = Date.now();

  let scannedStocks = 0;
  let backfilledStocks = 0;
  let skippedBackfilled = 0;
  let skippedMissing = 0;
  let appendedRows = 0;
  let timedOut = false;

  log(ss, "🕰️ Backfill start — seeding Price History from Yahoo daily history");

  for (let s = 0; s < CONFIG.SCANNERS.length; s++) {
    if (isTimedOut(RUN_START)) {
      timedOut = true;
      break;
    }

    const scanner = CONFIG.SCANNERS[s];
    const sheet = getSheetByNameSafe(ss, scanner.name);
    if (!sheet || sheet.getLastRow() < 2) continue;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, C._COUNT).getValues();

    for (let i = 0; i < data.length; i++) {
      if (isTimedOut(RUN_START)) {
        timedOut = true;
        break;
      }

      const row = data[i];
      const name = String(row[C.NAME] || "").trim();
      let symbol = String(row[C.SYMBOL] || "").trim();
      if (!name && !symbol) continue;

      scannedStocks++;

      if (!symbol || symbol === NO_SYMBOL_SENTINEL || symbol.startsWith(NO_SYMBOL_SENTINEL + "|")) {
        const parts = symbol ? symbol.split("|") : [];
        const bseHint = parts[1] || "";
        const bseCode = bseHint.startsWith("BSE:") ? bseHint.slice(4) : "";
        const resolved = name ? searchYahooSymbol(name, bseCode) : "";
        if (resolved) {
          symbol = resolved;
          row[C.SYMBOL] = resolved;
          sheet.getRange(i + 2, C.SYMBOL + 1).setValue(resolved);
        } else {
          skippedMissing++;
          continue;
        }
      }

      const stockKey = buildStockKey(symbol, name);
      const backfillKey = buildBackfilledStockKey(scanner.name, stockKey);
      if (state.backfilledStocks.has(backfillKey)) {
        skippedBackfilled++;
        continue;
      }

      try {
        const hist = normalizeHistory(fetchYahooHistory(symbol));
        Utilities.sleep(350);
        if (!hist || hist.closes.length < 2) {
          skippedMissing++;
          continue;
        }

        row[C.SYMBOL] = symbol;
        const built = buildBackfillRowsForStock(scanner.name, row, hist, state.existingRowKeys);
        if (built.rows.length === 0) {
          continue;
        }

        appendPriceHistoryRows(ss, built.rows);
        state.backfilledStocks.add(backfillKey);
        backfilledStocks++;
        appendedRows += built.rows.length;
      } catch (e) {
        log(ss, `⚠️ Backfill failed [${scanner.name}] ${symbol || name}: ${e.message}`);
      }
    }
  }

  updateDashboard(ss);

  const summary = `🕰️ Backfill done — scanned ${scannedStocks} stock(s), added ${appendedRows} row(s), seeded ${backfilledStocks} stock(s), ${skippedBackfilled} already seeded, ${skippedMissing} skipped — ${msToSec(Date.now() - RUN_START)}s`;
  log(ss, summary);
  if (timedOut) {
    setAutoBackfillStatus(AUTO_BACKFILL_PENDING);
    const nextRun = ensureBackfillTriggerScheduled(ss, "continuing automatic history backfill");
    log(ss, "⏱ Backfill reached time limit — it will continue automatically at " + formatDate(nextRun) + " IST");
    return;
  }

  deleteTriggersByHandler("backfillPriceHistoryFromYahoo");

  if (scannedStocks > 0) {
    setAutoBackfillStatus(AUTO_BACKFILL_COMPLETE);
  } else if (getAutoBackfillStatus() === AUTO_BACKFILL_PENDING) {
    log(ss, "ℹ️ No tracked stocks exist yet — automatic backfill remains armed and will schedule after the first populated run");
  }
}

/** Apply conditional formatting to all scanner sheets */
function applyFormattingToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  CONFIG.SCANNERS.forEach(sc => {
    const s = getSheetByNameSafe(ss, sc.name);
    if (s) applyConditionalFormatting(s);
  });
  Logger.log("✅ Conditional formatting applied (rules replaced, not appended)");
}

// ═════════════════════════════════════════════════════════════════════════════
// DEBUG — run manually from Apps Script editor to diagnose Screener 404s
// ═════════════════════════════════════════════════════════════════════════════

/**
 * Tests every scanner URL through the proxy and logs:
 *   - HTTP status code
 *   - Whether the response is JSON or HTML (login redirect)
 *   - First 300 chars of the response body
 *
 * Run this from: Apps Script editor → select "testScreenerConnectivity" → ▶ Run
 */
function testScreenerConnectivity() {
  CONFIG.SCANNERS.forEach(scanner => {
    const idMatch = scanner.url.match(/\/screens\/(\d+)([\/][^?]*)*/i);
    if (!idMatch) { Logger.log("❓ " + scanner.name + ": cannot parse screen ID"); return; }
    const screenId = idMatch[1];
    const slug     = (idMatch[2] || "").replace(/\/$/, "");

    const candidates = [
      "https://www.screener.in/screens/" + screenId + slug + "/?format=json",
      "https://www.screener.in/screens/" + screenId + "/?format=json",
      "https://www.screener.in/api/screens/" + screenId + "/?format=json",
    ].filter((v, i, a) => a.indexOf(v) === i);

    Logger.log("\n──────────────────────────────────");
    Logger.log("🔍 " + scanner.name + "  (ID: " + screenId + "  slug: " + (slug||"none") + ")");

    candidates.forEach(function(target) {
      // Test via proxy
      if (CONFIG.PROXY_URL) {
        const proxyUrl = CONFIG.PROXY_URL.replace(/\/$/, "") + "?url=" + encodeURIComponent(target);
        const proxyHeaders = { "User-Agent": "Mozilla/5.0 (compatible; GAS/1.0)" };
        if (CONFIG.SCREENER_COOKIE) proxyHeaders["X-Screener-Cookie"] = CONFIG.SCREENER_COOKIE;
        try {
          const r    = UrlFetchApp.fetch(proxyUrl, { headers: proxyHeaders, muteHttpExceptions: true });
          const code = r.getResponseCode();
          const body = r.getContentText().trim();
          const type = body.startsWith("<") ? "⚠️  HTML" : (body.startsWith("{") ? "✅ JSON" : "❓ unknown");
          Logger.log("  [proxy] HTTP " + code + " | " + type + " | " + target);
          if (code !== 200 || !body.startsWith("{")) Logger.log("    RAW: " + body.substring(0, 400));
        } catch(e) { Logger.log("  [proxy] ❌ " + e.message); }
      }
    });

    // Quick direct test (expected to 404 due to GAS IP block — confirms proxy is needed)
    const directUrl = "https://www.screener.in/screens/" + screenId + slug + "/?format=json";
    try {
      const r = UrlFetchApp.fetch(directUrl, { muteHttpExceptions: true });
      Logger.log("  [direct] HTTP " + r.getResponseCode() + " (expected 404 — GAS IP blocked)");
    } catch(e) { Logger.log("  [direct] ❌ " + e.message); }
  });

  Logger.log("\n══ LEGEND ════════════════════════");
  Logger.log("✅ JSON via proxy  → working correctly");
  Logger.log("⚠️  HTML via proxy  → session cookie wrong/expired");
  Logger.log("HTTP 404 via proxy → slug or ID wrong, or Screener down");
  Logger.log("HTTP 404 direct    → GAS IP blocked (normal — use proxy)");
}

/**
 * Fetches the raw HTML of a scanner URL via proxy and logs the first 1000 chars.
 * Use this to diagnose what Screener.in is actually returning for problem screens.
 * Change SCANNER_INDEX to 0, 1, or 2 to test different scanners.
 */
function testRawPage() {
  const SCANNER_INDEX = 1;  // ← 0=Monthly Opus, 1=Daily Scanner, 2=Below Book Value
  const scanner = CONFIG.SCANNERS[SCANNER_INDEX];
  if (!scanner) { Logger.log("No scanner at index " + SCANNER_INDEX); return; }

  const idMatch = scanner.url.match(/\/screens\/(\d+)(\/[^?]*)?/i);
  const screenId = idMatch ? idMatch[1] : "?";
  const slug     = idMatch && idMatch[2] ? idMatch[2].replace(/\/$/, "") : "";
  const htmlUrl  = "https://www.screener.in/screens/" + screenId + slug + "/";

  const proxyBase = CONFIG.PROXY_URL ? CONFIG.PROXY_URL.replace(/\/$/, "") : "";
  let fetchUrl = htmlUrl;
  const headers = { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" };
  if (proxyBase) {
    fetchUrl = proxyBase + "?url=" + encodeURIComponent(htmlUrl);
    headers["User-Agent"] = "Mozilla/5.0 (compatible; GAS/1.0)";
    if (CONFIG.SCREENER_COOKIE) headers["X-Screener-Cookie"] = CONFIG.SCREENER_COOKIE;
  }

  Logger.log("Fetching: " + fetchUrl);
  const resp = UrlFetchApp.fetch(fetchUrl, { muteHttpExceptions: true, headers: headers });
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  Logger.log("HTTP " + code + " | Length: " + body.length + " chars");
  Logger.log("hasRows=" + (body.indexOf("data-row-company-id") !== -1));
  Logger.log("hasQuery=" + (body.indexOf("query-builder") !== -1));
  Logger.log("hasLogin=" + (body.indexOf('action="/login/"') !== -1));
  Logger.log("--- First 1000 chars ---");
  Logger.log(body.substring(0, 1000));
  Logger.log("--- Last 500 chars ---");
  Logger.log(body.substring(Math.max(0, body.length - 500)));
}