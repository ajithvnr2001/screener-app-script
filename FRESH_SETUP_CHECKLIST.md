# Fresh Setup Checklist

## Use This File
Use `screenerv2.gs`.

## Before You Run Anything
1. Create a Google Sheet.
2. Open `Extensions -> Apps Script`.
3. Paste `screenerv2.gs` into the project.
4. Save the project.

## Update `CONFIG`
Edit these first:

- `SCANNERS`
- `SCREENER_COOKIE`
- `PROXY_URL`
- `ALERT_EMAIL`
- `YF_SUFFIX`
- `MARKET_START_HOUR`
- `MARKET_START_MINUTE`
- `MARKET_END_HOUR`
- `MARKET_END_MINUTE`
- `ACTIVE_DATE_RANGES`
- `TRIGGER_INTERVAL_MINUTES`
- `SIGNAL_PROFILE`

## Fastest Normal Setup
Run these in order:

```text
testYahooFetch()
setupTriggers()
```

Use this when you are okay waiting for the next valid scheduled slot.

## Fastest Immediate Setup
Run these in order:

```text
testYahooFetch()
testRunAllScannersManual()
setupTriggers()
```

Use this when you want the sheet populated immediately.

## What `setupTriggers()` Handles Automatically
After `setupTriggers()`:

- dashboard rebuild starts immediately
- the next `runAllScanners()` slot is scheduled automatically
- historical backfill is armed automatically
- future dashboard updates happen automatically
- future price-history snapshots happen automatically

## What You Normally Do Not Need To Run Manually
In normal use, you usually do not need:

```text
rebuildDashboardOnly()
refreshDashboardHistory()
backfillPriceHistoryFromYahoo()
```

## After Setup
Open:

- `📊 Dashboard`
- scanner tabs
- `🗒️ Log`

## If You Add A New Screener Later
1. Add the new scanner in `CONFIG.SCANNERS`
2. Save the script
3. Run:

```text
setupTriggers()
```

## If Something Looks Wrong
Try these in order:

```text
testYahooFetch()
rebuildDashboardOnly()
testRunAllScannersManual()
```

Then check:

- `🗒️ Log`
- Apps Script execution history
- `PROXY_URL`
- `SCREENER_COOKIE`
