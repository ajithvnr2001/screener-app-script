# Triggers Guide

## Trigger Model
`screenerv2.gs` does not use a permanent `everyMinutes(...)` trigger.

Instead, it uses a self-scheduling model:

- `setupTriggers()` creates the next valid `runAllScanners()` trigger
- when `runAllScanners()` finishes, it schedules the next one
- only one future `runAllScanners()` trigger is normally visible at a time

That is expected.

## What Is Automatic Now
After you run `setupTriggers()`:

- the dashboard is rebuilt immediately
- the next `runAllScanners()` slot is scheduled automatically
- historical backfill is armed automatically
- future dashboard updates happen automatically on every scheduled run
- future price history snapshots are appended automatically on every scheduled run

## Which Functions Are Recurring
### `runAllScanners()`
This is the main recurring automation.

It behaves like recurring execution because each run schedules the next valid one.

### `backfillPriceHistoryFromYahoo()`
This is automatic, but not forever recurring.

It will:

- start automatically after `setupTriggers()` when appropriate
- re-schedule itself automatically if it hits the Apps Script time limit
- stop once historical backfill is complete

### `rebuildDashboardOnly()`
This is not a separate forever trigger.

However, you normally do not need to run it manually because:

- `setupTriggers()` runs it once immediately
- each scheduled `runAllScanners()` run rebuilds the dashboard anyway

## What `setupTriggers()` Does
When you run `setupTriggers()` it:

- rebuilds the dashboard immediately
- validates schedule-related configuration
- deletes only old triggers that point to `runAllScanners`
- creates the next valid one-time trigger for `runAllScanners`
- arms automatic history backfill
- logs the current date filter, live window, and signal profile

It does not delete unrelated triggers for other handler functions.

## Schedule Controls In `CONFIG`
These settings control when live scans are allowed:

```text
ACTIVE_DATE_RANGES
MARKET_START_HOUR
MARKET_START_MINUTE
MARKET_END_HOUR
MARKET_END_MINUTE
TRIGGER_INTERVAL_MINUTES
```

Default behavior:

- `ACTIVE_DATE_RANGES: []` means weekday mode (`Mon-Fri`)
- live window is `9:00 AM` to `3:40 PM` IST
- default cadence is every `15` minutes

## Allowed Intervals
Only these values are allowed for `TRIGGER_INTERVAL_MINUTES`:

- `1`
- `5`
- `10`
- `15`
- `30`

If you change the interval, save the file and run:

```text
setupTriggers()
```

again.

## Why You See Only One Trigger
This project intentionally keeps only the next future run scheduled.

That means the Apps Script trigger list will usually show:

- one future `runAllScanners()` trigger
- sometimes a temporary `backfillPriceHistoryFromYahoo()` trigger while history seeding is still in progress

This is normal.

## No Off-Hours Trigger Wakeups
The current design avoids permanent off-hours wakeups.

It works by:

- calculating the next valid slot in code
- scheduling only that slot

So:

- no fixed off-hours repeating trigger should exist
- no separate market-open or market-close trigger is required

## Date Filter Modes
### Default mode
If `ACTIVE_DATE_RANGES` is empty, the script runs only on:

```text
Mon-Fri
```

within the configured intraday time window.

### Explicit date windows
If `ACTIVE_DATE_RANGES` is populated, that list overrides weekday mode.

Example:

```js
ACTIVE_DATE_RANGES: [
  { from: "2026-03-01", to: "2026-03-15", label: "Window 1" },
  { from: "2026-04-05", to: "2026-04-20", label: "Window 2" },
],
```

Rules:

- each range is inclusive
- date format must be `YYYY-MM-DD`
- `from` must be less than or equal to `to`
- the intraday time window still applies inside each valid date range

## Recommended Usage
### From scratch

```text
testYahooFetch()
setupTriggers()
```

Optional:

```text
testRunAllScannersManual()
```

Use that only if you want immediate data instead of waiting for the next scheduled slot.

### After changing schedule settings
If you change:

- scanner links
- date ranges
- market start/end time
- trigger interval

save the file and run:

```text
setupTriggers()
```

again.

## If You Add A New Screener
When you add a new entry to `CONFIG.SCANNERS`:

1. save `screenerv2.gs`
2. run `setupTriggers()`

Why:

- the dashboard explorer will refresh
- the live scheduling stays valid
- automatic historical backfill gets armed for the new screener

## Troubleshooting
### No trigger appears
Run:

```text
setupTriggers()
```

again and check the Apps Script logs.

### Trigger exists but rows are not updating
Check:

- execution history in Apps Script
- `🗒️ Log`
- date windows in `ACTIVE_DATE_RANGES`
- live window settings
- `PROXY_URL`
- `SCREENER_COOKIE`
- Yahoo connectivity via `testYahooFetch()`

### Backfill seems to stop midway
That is usually an Apps Script time-limit split.

Current behavior:

- it re-schedules itself automatically until complete

If needed, you can still run:

```text
backfillPriceHistoryFromYahoo()
```

manually, but normal setup should not require it.

## Quick Summary
- `runAllScanners()` is the recurring live engine.
- `backfillPriceHistoryFromYahoo()` is automatic until finished, then stops.
- `rebuildDashboardOnly()` runs during setup; scheduled runs keep the dashboard fresh.
- run `setupTriggers()` whenever schedule or scanner configuration changes.
