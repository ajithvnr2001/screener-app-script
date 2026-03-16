# Detailed Guide for `screenerv2.gs`

## Purpose
This is the full reference guide for the current `screenerv2.gs` workflow.

Use this document when you want one place that explains:

- what the project does end to end
- how the automation model works
- what every important sheet is for
- how the data moves through the system
- how to read the signals and charts
- what to do when something breaks

If you only want the shortest setup flow, read `read.md` first.

## What The Project Does
`screenerv2.gs` is a Google Apps Script that:

- reads one or more `screener.in` scanners
- stores every stock that appears in those scanners
- preserves old stocks even after they leave the scanner
- fetches Yahoo Finance history for tracked stocks
- calculates returns and technical indicators
- assigns a final signal label
- writes rolling snapshot history to a dedicated history sheet
- builds a dashboard and chart explorer inside Google Sheets

This is not just a simple alert script anymore. It is a small tracking system built inside Apps Script and Sheets.

## Main Outputs
After the script is running, the spreadsheet usually contains:

- one tab per entry in `CONFIG.SCANNERS`
- `📊 Dashboard`
- `📈 Price History`
- `📉 Chart Data`
- `🗒️ Log`

### Scanner tabs
Each scanner tab is your detailed stock table for that one Screener source.

### `📊 Dashboard`
This is the fast reading and monitoring surface.

### `📈 Price History`
This is the append-only snapshot history for stocks.

### `📉 Chart Data`
This is an internal helper sheet used to feed dashboard charts. It is not meant for direct manual use.

### `🗒️ Log`
This records important operational messages, failures, and summary lines.

## High-Level Architecture
The project has five main layers.

### 1. Screener fetch layer
Gets stocks from `screener.in` using JSON or HTML scraping through a proxy if needed.

### 2. Tracking layer
Adds new stocks to scanner tabs, keeps old ones, marks whether they are still in the source scanner, and stores first/last seen timestamps.

### 3. Market data layer
Uses Yahoo Finance to fetch price, volume, and daily history.

### 4. Calculation layer
Computes returns, moving averages, RSI, ADX, MACD, breakout metrics, and final signal labels.

### 5. Presentation layer
Writes data to scanner tabs, appends `📈 Price History`, and refreshes `📊 Dashboard`.

## Main Runtime Flow
The normal automatic flow is:

```text
setupTriggers()
  -> runAllScanners()
      -> processScanner()
          -> fetchScreenerScan()
          -> updateSheetPerformance()
              -> fetchYahooHistory()
              -> indicator calculations
              -> appendPriceHistoryRows()
      -> updateDashboard()
      -> ensureNextRunTrigger()
```

This means the system continuously loops forward by scheduling its own next run.

## The Trigger Model
This script does not use a permanent fixed repeating trigger.

Instead:

- `setupTriggers()` creates the next valid one-time `runAllScanners()` trigger
- when that run finishes, it schedules the next valid one

So in practice it behaves like recurring automation, but internally it is a chain of one-time triggers.

### Why this model is used
It makes schedule control easier when the run window is restricted by:

- weekdays or explicit date ranges
- a start time
- an end time
- a small set of allowed intervals

It also avoids permanent off-hours wakeups.

## Automatic vs Manual Functions
### Functions that become automatic after setup
- `runAllScanners()`
- dashboard refresh inside scheduled runs
- future `📈 Price History` row creation
- `backfillPriceHistoryFromYahoo()` until backfill is complete

### Functions that still exist mainly for manual use
- `testYahooFetch()`
- `testRunFirstScanner()`
- `testRunAllScannersManual()`
- `testUpdatePerformanceOnly()`
- `refreshDashboardHistory()`
- `rebuildDashboardOnly()`

### Important nuance
`rebuildDashboardOnly()` is not a forever trigger of its own. But:

- `setupTriggers()` calls it once immediately
- every scheduled `runAllScanners()` run also refreshes the dashboard

So in day-to-day use, the dashboard still updates automatically.

## Setup From Scratch
## Minimum hands-off setup
If you are okay waiting until the next scheduled slot:

```text
testYahooFetch()
setupTriggers()
```

## Minimum immediate-data setup
If you want to populate the sheet immediately:

```text
testYahooFetch()
testRunAllScannersManual()
setupTriggers()
```

## What `setupTriggers()` now handles
When you run `setupTriggers()`:

- system sheets are ensured
- the dashboard is rebuilt immediately
- the next automatic `runAllScanners()` slot is scheduled
- automatic historical price backfill is armed
- logs are written for schedule and signal profile status

## Configuration Reference
The most important configuration lives in the `CONFIG` object inside `screenerv2.gs`.

### `SCANNERS`
An array of scanner definitions.

Each item has:

- `name`
- `url`
- optional `color`

Example:

```js
{ name: "Monthly Opus V3", url: "https://www.screener.in/screens/3534290/monthly-opus-v3/?format=json", color: "#E0F7FA" }
```

Use a unique `name` for each scanner.

### `SCREENER_COOKIE`
Use this when:

- the screener is private
- the screener requires login
- direct anonymous access fails

This is usually your `sessionid` cookie from Screener.

### `PROXY_URL`
Used because Apps Script requests to Screener can be blocked.

This should point to your Cloudflare Worker or equivalent proxy endpoint.

### `ALERT_EMAIL`
If blank, email alerts are disabled.

### `YF_SUFFIX`
Default Yahoo exchange suffix.

Current intended behavior:

- normal flow tries NSE with `.NS`
- BSE fallback is handled where needed
- BSE-only symbols can still resolve with `.BO`

### Market window settings
- `MARKET_START_HOUR`
- `MARKET_START_MINUTE`
- `MARKET_END_HOUR`
- `MARKET_END_MINUTE`

These define the allowed intraday live run window in IST.

### `ACTIVE_DATE_RANGES`
If empty:

- the script uses normal weekday mode (`Mon-Fri`)

If populated:

- the script uses only those explicit date windows

Example:

```js
ACTIVE_DATE_RANGES: [
  { from: "2026-03-17", to: "2026-03-31", label: "March Window" },
  { from: "2026-04-10", to: "2026-04-25", label: "April Window" },
]
```

Rules:

- format is `YYYY-MM-DD`
- ranges are inclusive
- you can use multiple windows
- these override weekday mode

### `TRIGGER_INTERVAL_MINUTES`
Allowed values:

- `1`
- `5`
- `10`
- `15`
- `30`

### `SIGNAL_PROFILE`
Allowed values:

- `conservative`
- `balanced`
- `aggressive`

## Signal Profiles
These profiles change the thresholds used by the signal engine.

### `conservative`
Best when you want:

- stricter breakout confirmation
- fewer but higher-threshold signals

### `balanced`
Best default for general use.

### `aggressive`
Best when you want:

- earlier signals
- looser confirmation
- more noise accepted

You can also switch profiles with helper functions:

```text
setConservativeProfile()
setBalancedProfile()
setAggressiveProfile()
resetSignalProfile()
```

## How Stocks Are Tracked
The design is intentionally append-preserving.

When a stock first appears:

- it is added to the scanner tab
- `First Captured` is set
- `Last Seen` is set
- `In Screener?` becomes `✅`

When a stock disappears from the source scanner later:

- the row is not deleted
- `In Screener?` becomes `⬜`
- the historical tracking remains

This means the project is designed to preserve discovered names and their later performance.

## Yahoo Data and Technical Calculation Flow
For each tracked stock:

1. the script resolves the Yahoo symbol if needed
2. it fetches Yahoo history
3. it normalizes the arrays
4. it calculates returns
5. it calculates technical indicators
6. it writes back the latest snapshot into the scanner sheet
7. it appends a history row to `📈 Price History`

## Important Technical Indicators
### Moving averages
- `MA 20`
- `MA 50`
- `MA 200`

Main idea:

- `Price > MA20 > MA50 > MA200` is the strongest trend structure

### `RSI (14)`
Used to detect:

- oversold pullbacks
- bullish strength
- overbought extensions

### `ADX (14)`
Measures trend strength, not direction.

### `Vol Ratio (20)`
Measures participation.

Current design reduces intraday distortion by using the latest completed session during live hours.

### `MACD Line`
Used mainly for directional bias.

### `MACD Hist`
Used mainly for momentum acceleration / crossover context.

### `52W High Dist %`
Distance from the highest traded high in the last 52 weeks.

### `20D Breakout %`
Measures whether price is above or below the previous 20-session high.

## Signal Logic Overview
The signal engine combines:

- trend structure
- RSI context
- ADX trend strength
- volume confirmation
- MACD direction
- MACD acceleration
- breakout context
- 52-week-high proximity

This means the `Signal` field is a decision layer, not a raw indicator.

It should be read as:

- a summarized classification
- not a guarantee
- not a perfect predictor

## Reading The Dashboard
The dashboard has two main parts.

### 1. History Explorer
Top area with:

- scanner dropdown
- stock dropdown
- history point count
- latest price
- since-capture return
- latest signal
- first snapshot
- latest snapshot
- charts

Charts include:

- `Price History`
- `Since Capture %`
- `Short-Term Returns`

### 2. Scanner summary tables
Lower area listing each scanner with:

- current presence state
- signal
- sparkline trend
- ADX
- volume ratio
- MACD histogram
- 52-week-high proximity
- short-term returns

## How The History Explorer Picks Data
### Scanner dropdown
This is driven from `CONFIG.SCANNERS`, not only from the history sheet.

That means:

- newly added scanners can show up even before they have deep history

### Stock dropdown
This is built from:

- `📈 Price History` if history exists
- the scanner tab itself if history is still missing

That means a new scanner can appear quickly even before all charts are fully populated.

## `📈 Price History` Explained
This sheet is append-only.

Each row stores a snapshot of:

- timestamp
- run mode
- scanner
- stock identity
- current price
- capture price
- returns
- some technical values
- signal

Rows can come from:

- normal scheduled live runs
- manual tests
- automatic backfill

## Automatic Historical Backfill
`backfillPriceHistoryFromYahoo()` is now integrated into the automation flow.

What it does:

- takes existing tracked stocks
- fetches historical Yahoo daily data
- seeds old history rows from `First Captured` onward

Important:

- it is automatic after `setupTriggers()`
- if it hits Apps Script time limit, it re-schedules itself
- when the historical seeding is complete, it stops

So it is automatic, but not forever recurring.

## What Happens When You Add A New Screener
If you add a new screener to `CONFIG.SCANNERS`:

1. save `screenerv2.gs`
2. run `setupTriggers()`

That will:

- refresh the dashboard explorer scanner list
- keep normal live scheduling correct
- re-arm automatic backfill for the new scanner

## What Is Actually Recurring
### Forever-recurring practical behavior
- `runAllScanners()`
- dashboard refresh inside scheduled runs
- future `📈 Price History` snapshots

### Automatic until complete
- `backfillPriceHistoryFromYahoo()`

### One-shot convenience helper
- `rebuildDashboardOnly()`

## Troubleshooting Guide
### Problem: no scanner data appears
Check:

- `PROXY_URL`
- `SCREENER_COOKIE`
- the scanner URL

Use:

```text
testScreenerConnectivity()
testRawPage()
```

### Problem: Yahoo data is missing
Check:

- `YF_SUFFIX`
- symbol resolution

Use:

```text
testYahooFetch()
testSymbolSearch()
```

### Problem: dashboard charts are blank
Try:

```text
rebuildDashboardOnly()
refreshDashboardHistory()
```

Also check:

- whether `History Points` is very low
- whether automatic backfill has finished
- whether the selected scanner/stock has history yet

### Problem: trigger exists but nothing updates
Check:

- Apps Script execution history
- `🗒️ Log`
- date range settings
- market window settings

### Problem: new scanner does not appear
Do:

```text
setupTriggers()
```

Then wait for the next scheduled run, or run:

```text
rebuildDashboardOnly()
```

### Problem: backfill seems partial
Usually this means the script split the work across multiple runs due to Apps Script time limits.

Current intended behavior:

- it should auto-resume until complete

## Recommended Operator Routine
### If you are starting from scratch

```text
testYahooFetch()
setupTriggers()
```

Optional for immediate population:

```text
testRunAllScannersManual()
```

### If you changed scanners or schedule

```text
setupTriggers()
```

### If the dashboard looks wrong

```text
rebuildDashboardOnly()
```

### If you want to inspect just one part of the pipeline

```text
testRunFirstScanner()
testUpdatePerformanceOnly()
refreshDashboardHistory()
```

## Operational Notes
- scanner rows are preserved even after stocks leave the scanner
- dashboard data is a quick decision layer, not the full story
- signals are heuristics, not guarantees
- the history explorer becomes more useful as more snapshot rows accumulate
- if a new scanner has no history yet, that does not mean the scanner is broken

## Document Map
Use the docs like this:

- `read.md` for the quickest overview
- `SETUP_GUIDE.md` for setup and config
- `TRIGGERS_GUIDE.md` for automation behavior
- `ANALYSIS_LOOKUP.md` for reading signals and charts
- `audit.md` for technical indicator review
- this file for the complete operational reference

## Final Summary
The current `screenerv2.gs` system is designed so that after `setupTriggers()`:

- scheduled runs continue automatically
- dashboard updates continue automatically
- future price history keeps growing automatically
- historical backfill runs automatically until it is finished

Manual helper functions still exist, but they are mainly for:

- testing
- repair
- immediate refresh
- troubleshooting
