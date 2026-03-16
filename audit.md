# Technical Audit for `screenerv2.gs`

This audit remains focused on the technical-analysis and signal engine.

It does not attempt to fully document the newer dashboard, history explorer, auto-backfill, or trigger setup UX. For operational guidance, use:

- `SETUP_GUIDE.md`
- `TRIGGERS_GUIDE.md`
- `ANALYSIS_LOOKUP.md`

## Audit Summary
This audit reviews the current technical-analysis engine in `screenerv2.gs` after the latest fixes for:

- true `52W High Dist %` based on highs
- completed-session `Vol Ratio (20)`
- separate `MACD Line` vs `MACD Hist` handling

### Overall conclusion
The current version is materially stronger than the earlier revisions. I did **not** find a critical arithmetic defect in the current implementations of:

- `SMA`
- `RSI`
- `EMA`
- `MACD`
- `ADX`
- `20D breakout`
- `52W high distance`

However, it is still not possible to guarantee literal `100% accuracy` in a trading sense, because:

- technical indicators are probabilistic, not deterministic
- live market bars are partial and can change until the session closes
- the final `Signal` is a heuristic decision layer on top of valid indicators

So the right target is:

- **high programmatic correctness**
- **clear signal semantics**
- **low avoidable distortion**

On those three goals, the current version is in a good state.

## Scope Reviewed
This audit focused on:

- Yahoo OHLCV fetch and normalization
- history shaping before calculations
- indicator implementations
- signal-combination logic
- intraday data semantics
- scheduler behavior only where it affects technical updates

Primary functions reviewed:

- `fetchYahooHistory()`
- `normalizeHistory()`
- `updateSheetPerformance()`
- `calcMA()`
- `calcRSI()`
- `calcEMA()`
- `calcMACD()`
- `calcVolumeRatio()`
- `calcDistanceFromHigh()`
- `calcBreakoutPct()`
- `calcADX()`
- `calcSignal()`

## Method Used
The audit was done using:

- source inspection of the current `screenerv2.gs`
- targeted synthetic sanity tests for bullish, bearish, sideways, and intraday scenarios
- spot tests of edge cases that previously caused distortion

### Sanity checks performed
1. Verified `52W High Dist %` now uses highs instead of closes.
2. Verified intraday `Vol Ratio (20)` ignores the partial current session during live market hours.
3. Verified a strong downtrend still classifies as `🔴 SELL`.
4. Verified a bullish trend with positive `MACD Line` but near-zero `MACD Hist` no longer loses bullish bias incorrectly.
5. Verified syntax and diagnostics remain clean after the fixes.

## What Is Now Correct

### 1. `52W High Dist %` semantics are now correct
Current behavior:

- uses the highest **traded high** in the lookback window
- compares current **close** to that high

This matches the practical meaning traders usually expect when reading “distance from 52-week high.”

### 2. Intraday volume distortion has been reduced
Current behavior:

- during live market hours, the current partial session is ignored
- volume ratio is based on the latest completed session versus prior completed sessions

This is much better than comparing partial current-day volume against full historical sessions.

### 3. MACD trend vs acceleration is now separated
Current behavior:

- `MACD Line` is used for trend bias
- `MACD Hist` is used for acceleration / crossover context

This is more accurate than treating histogram sign alone as full bullish or bearish confirmation.

### 4. ADX implementation is structurally standard
The current ADX implementation:

- computes `TR`, `+DM`, `-DM`
- applies Wilder smoothing
- computes `DI`, `DX`, and final `ADX`

The structure is consistent with common implementations.

### 5. RSI implementation is structurally standard
The RSI logic:

- builds consecutive changes
- seeds average gain/loss
- applies Wilder smoothing

This is mathematically sound for a 14-period RSI.

## Remaining Findings

### Medium: intraday signals still mix live price with completed-session volume
Current behavior is intentionally conservative:

- price-based indicators use the latest available bar
- volume confirmation uses the last completed session during live hours

This is better than the old partial-volume distortion, but it still means the signal can mix:

- current live price structure
- previous completed-session volume confirmation

Impact:

- intraday breakout signals may appear slightly “late” on volume confirmation
- the system is more stable, but not perfectly synchronized intraday

This is not a formula bug. It is a design trade-off.

### Medium: moving average logic is tolerant, not textbook-strict
`calcMA()` accepts a period if at least `80%` of bars in the lookback window are valid.

Impact:

- on normal Yahoo daily data, this is usually harmless
- on gap-heavy or poor-quality histories, the result can differ from a strict textbook SMA

This is a resilience feature, not a hard bug, but it does trade purity for continuity.

### Low: missing-confirmation fallback is permissive
In `calcSignal()`, bullish confirmation can still pass when both MACD line and volume confirmation are unavailable.

Intent:

- avoid losing all signal output on sparse-data or incomplete-data rows

Risk:

- some low-quality rows may look more confirmed than they truly are

This is low risk because most tracked symbols with usable history should have both MACD and volume data.

### Low: no automated regression test suite exists
The repo has useful manual test functions, but not a deterministic unit-test harness.

Impact:

- formula regressions are still possible during future refactors
- confidence depends on manual spot testing instead of repeatable assertions

## Signal Logic Assessment
The signal engine is now substantially better than the earlier version.

### What it gets right
- requires `MA200` for full “perfect bull” classification
- uses `ADX` as trend-strength gating
- uses `MACD Line` for direction bias
- uses `MACD Hist` for acceleration-sensitive states
- uses volume and breakout context as confirmation

### Important interpretation note
The `Signal` column is a **decision layer**, not a raw indicator. That means:

- the indicators can all be mathematically valid
- the final label can still be debatable depending on strategy style

For example:

- a momentum trader may want earlier `BUY`
- a swing trader may prefer stricter pullback confirmation
- a long-only investor may ignore overbought caution labels

So “signal accuracy” is partly strategic, not purely mathematical.

## Data Quality Assumptions
The current logic assumes Yahoo daily data is mostly clean.

Key assumptions:

- trailing incomplete close may be `NaN` and should be trimmed
- the remaining OHLCV arrays are aligned enough to use together
- daily timestamps are trustworthy for session-completion checks

If Yahoo changes response structure or starts returning inconsistent high/low/close availability intraday, some calculations may degrade even if the formulas themselves remain correct.

## Things I Would Not Change Right Now
I would **not** add more indicators just to make the system look more advanced.

Indicators that are not currently necessary:

- Bollinger Bands
- Stochastic Oscillator
- CCI
- Supertrend
- Williams %R
- Ichimoku

Reason:

- they would increase overlap
- they would complicate interpretation
- they would not guarantee better practical accuracy

The current stack is already broad enough:

- trend: `MA20`, `MA50`, `MA200`, `ADX`
- momentum: `RSI`, `MACD Line`, `MACD Hist`
- participation: `Vol Ratio (20)`
- price location: `52W High Dist %`, `20D Breakout %`

## Best Next Improvements
If the goal is even higher reliability, these are the next best upgrades.

### 1. Add an indicator mode: `live` vs `eod`
This would be the highest-value enhancement.

Suggested behavior:

- `live` mode: allow current live bar in price indicators
- `eod` mode: ignore the current session entirely until market close

Benefit:

- eliminates mixed-state intraday semantics when desired
- gives a cleaner “confirmed close” version of the model

### 2. Add a small test harness
Suggested tests:

- monotonic uptrend
- monotonic downtrend
- sideways range
- partial intraday volume case
- 52-week high/high-vs-close check
- missing-data tolerance cases

Benefit:

- future changes become safer
- indicator correctness becomes easier to preserve

### 3. Add an optional signal explanation column
Example values:

- `Trend+ADX+MACD`
- `Breakout+Volume`
- `Overbought`
- `Weak trend`

Benefit:

- easier debugging
- easier trust-building for discretionary review

## Final Assessment

### Current rating
- **Formula correctness**: strong
- **Live-market semantics**: good, with known design trade-offs
- **Signal-engine quality**: good and materially improved
- **Predictive certainty**: inherently limited by market behavior

### Final verdict
The current `screenerv2.gs` technical engine is in a **good production-ready state** for a practical screener.

There are no critical remaining formula defects found in this audit.

The biggest remaining gap is not indicator math. It is the unavoidable difference between:

- mathematically correct indicators
- and perfectly predictive trading decisions

If you want the next step after this audit, the best improvement is **not another indicator**. It is a **`live` vs `eod` technical mode** plus a **small regression test harness**.
