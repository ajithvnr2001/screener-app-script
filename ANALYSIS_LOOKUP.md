# Analysis Lookup Guide

## Purpose
This guide explains how to read:

- the `📊 Dashboard`
- the `📈 History Explorer`
- the scanner sheets
- the main signal and technical columns in `screenerv2.gs`

## Best Reading Order
Use this order for any stock:

1. Check `In Screener?`
2. Check `Signal`
3. Check trend structure with `Price`, `MA 20`, `MA 50`, `MA 200`
4. Check confirmation with `ADX`, `Vol Ratio (20)`, `MACD Line`, `MACD Hist`
5. Check location with `52W High Dist %` and `20D Breakout %`
6. Check return columns and the history charts

## Dashboard Quick View
The main dashboard table is for fast shortlisting.

It shows:

- `Symbol`
- `Name`
- `In Screener?`
- `Price ₹`
- `Signal`
- `Trend` sparkline
- `ADX`
- `Vol x`
- `MACD Hist`
- `52W High %`
- `1D %`, `1W %`, `1M %`, `3M %`

Use the dashboard table to decide what deserves deeper review. Use the scanner tabs for full detail.

## History Explorer
The top section of `📊 Dashboard` is the `History Explorer`.

### What it does
It lets you choose:

- one scanner
- one stock inside that scanner

Then it shows:

- `History Points`: number of stored history rows used for that stock
- `Latest Price`
- `Since Capture`
- `Latest Signal`
- `First Snapshot`
- `Latest Snapshot`

It also draws:

- `Price History`
- `Since Capture %`
- `Short-Term Returns` (`1D`, `1W`, `1M`)

### Important note
The scanner dropdown is driven from `CONFIG.SCANNERS`, so newly added scanners can appear there even before they have full history.

If a selected stock has little or no history yet:

- summary values may be limited
- charts may be blank or very short
- automatic backfill or the next scheduled run will fill them in

## Price History Sheet
`📈 Price History` is append-only.

Each row is one stored snapshot for one stock.

Sources of rows:

- scheduled live runs
- manual test runs
- automatic historical Yahoo backfill

This sheet is the source for the history explorer charts and trend sparklines.

## Scanner Sheet Columns
### Tracking columns

| Column | Meaning | How to use it |
| --- | --- | --- |
| `Symbol` | Exchange ticker used for Yahoo Finance lookups | Main stock identifier |
| `Name` | Company name from Screener | Useful when symbol is unclear |
| `First Captured` | When the stock first entered tracking | Shows how early the stock was found |
| `Last Seen` | Last time the stock appeared in the scanner | Helps detect drop-offs |
| `In Screener?` | `✅` if still in the source screen, `⬜` if not | First filter to check |

### Price and return columns

| Column | Meaning | How to use it |
| --- | --- | --- |
| `Capture Price ₹` | Price when the script first got valid Yahoo data | Base price for tracking performance |
| `Current Price ₹` | Latest fetched price | Current reference |
| `Since Capture %` | Return vs capture price | Measures outcome since discovery |
| `1D %`, `1W %`, `1M %`, `3M %`, `6M %`, `1Y %`, `2Y %`, `3Y %` | Point-to-point returns | Shows short and long horizon trend |
| `Avg Weekly %`, `Avg Monthly %`, `Avg 3M %`, `Avg 6M %`, `Avg 1Y %` | Average non-overlapping period returns | Helps compare consistency, not just the latest move |

### Technical columns

| Column | Meaning | How to use it |
| --- | --- | --- |
| `RSI (14)` | Momentum oscillator | Low can mean oversold, high can mean extended |
| `MA 20` | 20-day moving average | Short trend |
| `MA 50` | 50-day moving average | Intermediate trend |
| `MA 200` | 200-day moving average | Long trend health |
| `ADX (14)` | Trend strength | Higher means stronger trend |
| `Vol Ratio (20)` | Latest completed-session volume versus the previous 20-session average | Above `1.0` means better participation |
| `MACD Line` | Direction / trend bias | Positive is bullish bias, negative is bearish bias |
| `MACD Hist` | Acceleration / crossover context | Positive is improving momentum, negative is weakening momentum |
| `52W High Dist %` | Distance from the highest traded high in the last 52 weeks | Smaller is stronger for momentum names |
| `20D Breakout %` | Distance vs previous 20-session high | Positive means price is above the recent breakout level |
| `Signal` | Final synthesized rating | Primary action summary |
| `Last Updated` | Last successful technical update time | Freshness check |

## How To Read The Main Technical Fields
### `MA 20`, `MA 50`, `MA 200`
The strongest trend structure is usually:

```text
Price > MA20 > MA50 > MA200
```

That means:

- price is above short-term trend
- short-term trend is above medium-term trend
- medium-term trend is above long-term trend

### `RSI (14)`
Under the default `balanced` profile:

- `<= 32` is treated as oversold
- `52` to `78` is the main bullish-strength band
- `>= 78` is treated as overbought

Interpretation:

- low RSI during a strong uptrend can mean pullback opportunity
- very high RSI can mean the move is extended and needs tighter risk control

### `ADX (14)`
ADX measures trend strength, not direction.

Under the default `balanced` profile:

- `>= 20` = strong enough trend confirmation
- `< 16` = weak trend

Interpretation:

- strong ADX plus bullish MA structure is better than bullish MAs alone
- low ADX usually means range or weak follow-through

### `Vol Ratio (20)`
This compares the latest completed-session volume with the average of the previous 20 completed sessions.

Interpretation under the default `balanced` profile:

- `> 1.25` = solid participation
- much higher values can help confirm breakouts
- low values mean weaker participation

### `MACD Line`
Use this as the main directional bias:

- positive = bullish direction bias
- negative = bearish direction bias
- near zero = transition area

### `MACD Hist`
Use this as the momentum-acceleration view:

- positive = improving momentum
- negative = weakening momentum
- near zero = little acceleration either way

### `52W High Dist %`
Interpretation:

- near `0` = at or near the 52-week high
- low values = strong momentum location
- larger values = farther below major highs

### `20D Breakout %`
Interpretation:

- positive = above the previous 20-session high
- near zero = testing breakout area
- negative = below the breakout level

## Signal Meanings
### Strongest bullish signals

| Signal | Meaning |
| --- | --- |
| `🚀 BREAKOUT BUY` | Strong trend, breakout context, and confirmation all aligned |
| `🚀 STRONG BUY (Pullback)` | Strong uptrend with pullback-style setup |
| `🚀 STRONG BUY` | Strong uptrend with solid confirmation |

### Bullish but less ideal

| Signal | Meaning |
| --- | --- |
| `✅ BUY` | Bullish structure with enough confirmation |
| `🟡 BUY (Overbought — trail SL)` | Still bullish, but extended |
| `🟡 HOLD (Overbought — watch)` | Trend intact but stretched |
| `🟡 HOLD (Weak Trend)` | Structure exists, but confirmation is weaker |

### Neutral / transition

| Signal | Meaning |
| --- | --- |
| `⏸️ HOLD (Pullback)` | Above medium trend but pulling back |
| `⏸️ HOLD (Oversold — possible bounce)` | Pullback setup, but not strong enough yet |
| `⏸️ HOLD` | Mixed evidence |

### Bearish / weak

| Signal | Meaning |
| --- | --- |
| `⚠️ WEAK (Oversold — possible bounce)` | Weak structure, but short-term bounce possible |
| `⚠️ SELL (Oversold — watch reversal)` | Bearish structure with oversold conditions |
| `🔴 SELL` | Clear bearish structure |
| `⚪ Insufficient Data` | Not enough clean history for a reliable signal |

## Practical Reading Examples
### Strong continuation candidate
Look for:

- `In Screener? = ✅`
- `Signal = 🚀 BREAKOUT BUY` or `✅ BUY`
- `Price > MA20 > MA50 > MA200`
- `ADX` strong
- `Vol Ratio (20)` above average
- `MACD Line` positive
- `MACD Hist` positive
- `52W High Dist %` low
- `20D Breakout %` positive

### Trend is good but extended
Look for:

- `Signal = 🟡 BUY (Overbought — trail SL)` or `🟡 HOLD (Overbought — watch)`
- very high RSI
- price far above short-term averages

These may still work, but risk/reward is usually worse than an earlier entry.

### De-prioritize
Look for:

- `In Screener? = ⬜`
- `Signal = 🔴 SELL`
- `ADX` weak
- `MACD Line` negative
- `MACD Hist` negative
- price below `MA20` and `MA50`

## Signal Profiles
The signal engine depends on the active profile:

- `conservative`: stricter confirmation, fewer signals
- `balanced`: default
- `aggressive`: earlier signals, more noise

Profile helper functions:

```text
setConservativeProfile()
setBalancedProfile()
setAggressiveProfile()
resetSignalProfile()
```

## Rule Of Thumb
Do not rely on one column alone.

The best reads usually come from agreement between:

- `Signal`
- MA structure
- `ADX`
- `Vol Ratio (20)`
- `MACD Line`
- `MACD Hist`
- `52W High Dist %`
- `20D Breakout %`
- the history charts

If most of those agree, the setup is usually much better than one driven by only RSI or only one recent return value.
