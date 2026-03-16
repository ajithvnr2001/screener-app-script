# Setup Guide

## Use This Script File
Use `screenerv2.gs`.

This is the active version that includes:

- automatic self-scheduling intraday runs
- automatic dashboard rebuild during setup and scheduled runs
- automatic historical price backfill until complete
- `📈 Price History` and history charts

## What The Project Creates
After setup and the first successful data run, the sheet will contain:

- one tab per scanner in `CONFIG.SCANNERS`
- `📊 Dashboard`
- `📈 Price History`
- `📉 Chart Data` (internal helper sheet, usually hidden)
- `🗒️ Log`

## Before You Start
You should have:

- a Google Sheet
- access to `Extensions -> Apps Script`
- your `screener.in` URLs
- a valid Screener `sessionid` cookie if you use private or login-gated screens
- a working proxy URL if direct Screener requests are blocked

## Edit `CONFIG` First
Open `screenerv2.gs` and review the `CONFIG` block.

Most important settings:

- `SCANNERS`: list of screener links to track
- `SCREENER_COOKIE`: needed for private/login-protected screens
- `PROXY_URL`: your Cloudflare Worker or equivalent proxy
- `ALERT_EMAIL`: set `""` to disable mail
- `YF_SUFFIX`: keep `.NS` for normal NSE-first behavior; the script also tries `.BO` when needed
- `MARKET_START_HOUR`, `MARKET_START_MINUTE`
- `MARKET_END_HOUR`, `MARKET_END_MINUTE`
- `ACTIVE_DATE_RANGES`: leave `[]` for normal `Mon-Fri` mode
- `TRIGGER_INTERVAL_MINUTES`: only `1`, `5`, `10`, `15`, `30`
- `SIGNAL_PROFILE`: `conservative`, `balanced`, or `aggressive`

Do not rename these system sheets once created:

- `📊 Dashboard`
- `📈 Price History`
- `📉 Chart Data`
- `🗒️ Log`

## Date Range Example
If you want explicit calendar windows instead of the default weekday mode:

```js
ACTIVE_DATE_RANGES: [
  { from: "2026-03-17", to: "2026-03-31", label: "March Window" },
  { from: "2026-04-10", to: "2026-04-25", label: "April Window" },
],
```

Rules:

- each range is inclusive
- you can add multiple ranges
- when this list is not empty, it overrides the default `Mon-Fri` date filter
- the intraday time window still applies

## Fastest Hands-Off Setup
If you are okay waiting until the next valid scheduled slot for the first real run:

```text
testYahooFetch()
setupTriggers()
```

What `setupTriggers()` now does automatically:

- rebuilds the dashboard immediately
- schedules the next `runAllScanners()` slot
- arms automatic historical backfill
- keeps future dashboard and history updates automatic

## Fastest Immediate Setup
If you want data right away instead of waiting for the next scheduled slot:

```text
testYahooFetch()
testRunAllScannersManual()
setupTriggers()
```

Why this version is useful:

- `testYahooFetch()` confirms Yahoo Finance access and indicator calculations
- `testRunAllScannersManual()` fills sheets immediately
- `setupTriggers()` then takes over the ongoing automation

## What Becomes Automatic After `setupTriggers()`
Once `setupTriggers()` is run successfully:

- `runAllScanners()` continues automatically at the configured interval
- the dashboard refreshes automatically during scheduled runs
- `📈 Price History` keeps appending new snapshots automatically
- `backfillPriceHistoryFromYahoo()` runs automatically until old history seeding is complete

Important nuance:

- `runAllScanners()` behaves like recurring automation, but internally it is a chain of self-scheduled one-time triggers
- `backfillPriceHistoryFromYahoo()` is automatic until it finishes, then it stops

## If You Add A New Screener Later
If you edit `CONFIG.SCANNERS` and add a new screener:

1. save `screenerv2.gs`
2. run `setupTriggers()` once

That refreshes the dashboard, keeps scheduling correct, and re-arms automatic historical backfill for the new screener.

## Manual Helper Functions
These still exist for testing or repair:

```text
testRunFirstScanner()
testRunAllScannersManual()
testUpdatePerformanceOnly()
rebuildDashboardOnly()
refreshDashboardHistory()
backfillPriceHistoryFromYahoo()
```

In normal operation, you usually do not need to run the last three manually.

## Signal Profiles
Available profiles:

- `conservative`
- `balanced`
- `aggressive`

Profile helper functions:

```text
setConservativeProfile()
setBalancedProfile()
setAggressiveProfile()
resetSignalProfile()
```

Notes:

- `resetSignalProfile()` removes the stored override and falls back to `CONFIG.SIGNAL_PROFILE`
- the active profile affects signal thresholds used during runs

## Common Setup Problems
### Screener returns no data
Check:

- `PROXY_URL`
- `SCREENER_COOKIE`
- the screener URL itself

Useful functions:

```text
testScreenerConnectivity()
testRawPage()
```

### Yahoo Finance data is missing
Check:

- `YF_SUFFIX`
- whether the symbol was resolved correctly

Useful functions:

```text
testYahooFetch()
testSymbolSearch()
```

### Dashboard exists but charts look empty
Run:

```text
rebuildDashboardOnly()
```

If the selected stock still has no history points, wait for automatic backfill or the next scheduled data run.

### Existing sheets do not show new schema
Run:

```text
testUpdatePerformanceOnly()
```

or:

```text
rebuildDashboardOnly()
```

## Day-To-Day Usage
Normally you only need:

- automatic trigger runs
- `📊 Dashboard` for the quick view
- scanner tabs for full row-level data
- `🗒️ Log` for troubleshooting

If you change schedule settings, date windows, or add/remove scanners, save the script and run:

```text
setupTriggers()
```
