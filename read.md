# `screenerv2.gs` Quick Read

## What This Project Does
`screenerv2.gs` tracks multiple `screener.in` scanners in Google Sheets, enriches them with Yahoo Finance price and technical data, stores price-history snapshots, and builds a dashboard with charts.

It creates and maintains:

- one sheet tab per scanner in `CONFIG.SCANNERS`
- `📊 Dashboard`
- `📈 Price History`
- `📉 Chart Data`
- `🗒️ Log`

## Fastest Setup
If you want the system to take over automatically:

```text
testYahooFetch()
setupTriggers()
```

If you want data immediately instead of waiting for the next scheduled slot:

```text
testYahooFetch()
testRunAllScannersManual()
setupTriggers()
```

## What `setupTriggers()` Now Handles
After you run `setupTriggers()`:

- dashboard rebuild starts immediately
- the next `runAllScanners()` slot is scheduled automatically
- historical price backfill is armed automatically
- future dashboard updates happen automatically
- future price-history snapshots are captured automatically

In normal use, you should not need to manually run:

- `rebuildDashboardOnly()`
- `backfillPriceHistoryFromYahoo()`

## Important Automation Notes
- `runAllScanners()` is the recurring live engine, implemented as a self-scheduling chain of one-time triggers
- `backfillPriceHistoryFromYahoo()` runs automatically until historical seeding is complete, then stops
- if you add a new screener later, save the script and run `setupTriggers()` once again

## Main Docs
- `FRESH_SETUP_CHECKLIST.md`: shortest step-by-step setup list
- `SETUP_GUIDE.md`: full setup and configuration flow
- `TRIGGERS_GUIDE.md`: how automation, scheduling, and backfill triggers work
- `ANALYSIS_LOOKUP.md`: how to read the dashboard, charts, signals, and scanner sheets
- `DETAILED_GUIDE.md`: heavy end-to-end operational reference
- `audit.md`: technical audit of the indicator and signal engine

## Most Useful Manual Functions
These are still available for testing or repair:

```text
testRunFirstScanner()
testRunAllScannersManual()
testUpdatePerformanceOnly()
rebuildDashboardOnly()
refreshDashboardHistory()
backfillPriceHistoryFromYahoo()
```

## If Something Looks Wrong
Check these in order:

1. `🗒️ Log`
2. Apps Script execution history
3. `PROXY_URL`
4. `SCREENER_COOKIE`
5. `testYahooFetch()`
6. `testScreenerConnectivity()`

## Recommended Rule
Whenever you change:

- scanner links
- trigger timing
- date ranges
- market window

save the script and run:

```text
setupTriggers()
```
