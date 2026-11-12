
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scanner 7 Historical Backtest
Kraken Futures | 30 days | 5m SQZMOM + 15m Pure Momentum
Independent symbols and independent 54 scenarios.
No BTC direction filter, no BTC grouping, no BTC dependency.
"""

from __future__ import annotations

import math
import time
import logging
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
import pandas as pd
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIG
# ============================================================

BASE = "https://futures.kraken.com"
CHARTS = f"{BASE}/api/charts/v1"
INSTRUMENTS = f"{BASE}/derivatives/api/v3/instruments"
TICKERS = f"{BASE}/derivatives/api/v3/tickers"

DAYS = 30
TOP_CONTRACTS = 10
INITIAL_EQUITY = 100.0
LEVERAGE = 10.0

# Explicit simulation assumptions; these are NOT historical fee data.
FEE_RATE = 0.0005          # 0.05% per side
SLIPPAGE_RATE = 0.0002     # 0.02% per side
FUNDING_RATE = 0.0         # historical funding not modeled

SQZ_MS = 5 * 60 * 1000
PMO_MS = 15 * 60 * 1000

BB_LENGTH = 20
KC_LENGTH = 20
PMO_LENGTH = 45
PMO_UP = 0.05
PMO_DOWN = -0.05

SL_BUFFERS = (0.0, 0.001, 0.002)
MARGIN_FRACTIONS = (0.01, 0.10)
RR_VALUES = (1, 2, 3, 4, 8, 10)

OUTPUT = Path("backtest_report.xlsx")
LOG = logging.getLogger("scanner7_backtest")

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Scanner7-Historical-Backtest/1.0",
    "Accept": "application/json",
})


# ============================================================
# DATA STRUCTURES
# ============================================================

@dataclass
class Candle:
    time: int
    open: float
    high: float
    low: float
    close: float
    volume: float = 0.0


@dataclass(frozen=True)
class Scenario:
    name: str
    margin_fraction: float
    buffer: float
    exit_type: str
    rr: float | None = None


# ============================================================
# API
# ============================================================

def api_get(url, params=None, attempts=5):
    error = None
    for attempt in range(attempts):
        try:
            r = SESSION.get(url, params=params, timeout=30)
            if r.status_code in (429, 500, 502, 503, 504):
                raise RuntimeError(f"Temporary HTTP {r.status_code}")
            r.raise_for_status()
            data = r.json()
            if not isinstance(data, dict):
                raise RuntimeError("Unexpected API response")
            if data.get("result") == "error":
                raise RuntimeError(str(data))
            return data
        except Exception as exc:
            error = exc
            if attempt + 1 < attempts:
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"API request failed: {error}")


def get_market_metadata():
    instruments = api_get(INSTRUMENTS).get("instruments", [])
    tickers = api_get(TICKERS).get("tickers", [])

    ticker_map = {
        str(x.get("symbol", "")).upper(): x
        for x in tickers if x.get("symbol")
    }

    rows = []
    for item in instruments:
        symbol = str(item.get("symbol", "")).upper()
        kind = str(item.get("type", "")).lower()

        if not symbol or item.get("tradeable") is False:
            continue
        if item.get("suspended") is True:
            continue

        # Avoid inverse and quanto settlement unless their contract
        # conversion can be verified. This version accepts linear PF
        # perpetual contracts only.
        if not symbol.startswith("PF_"):
            continue
        if kind not in ("futures_vanilla", "flexible_futures"):
            continue

        ticker = ticker_map.get(symbol)
        if not ticker:
            continue

        last = number(ticker.get("last"))
        volume = number(ticker.get("vol24h"))

        if last is None or volume is None or last <= 0 or volume <= 0:
            continue

        # PF linear contracts: vol24h is treated as base-asset volume.
        # If Kraken does not provide a verifiable base-volume field,
        # the contract is excluded rather than guessed.
        volume_usd = volume * last

        rows.append({
            "symbol": symbol,
            "type": kind,
            "last_price": last,
            "volume_24h": volume,
            "estimated_volume_usd": volume_usd,
            "volume_method": "vol24h ¡Á last; linear PF contract",
        })

    if not rows:
        raise RuntimeError(
            "No eligible linear PF contracts with verifiable volume data."
        )

    rows.sort(key=lambda x: x["estimated_volume_usd"], reverse=True)
    return rows[:TOP_CONTRACTS], ticker_map


def number(value):
    try:
        if value is None:
            return None
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def fetch_candles(symbol, resolution, interval_ms, start_ms, end_ms):
    """
    Fetch historical candles in small time windows.
    Kraken chart timestamps are milliseconds; from/to are seconds.
    """
    url = f"{CHARTS}/trade/{symbol}/{resolution}"
    result = {}
    errors = []

    # Three-day windows keep each request comfortably below
    # the chart endpoint's maximum candle count.
    window_seconds = 3 * 24 * 60 * 60
    cursor = start_ms // 1000
    end_seconds = end_ms // 1000

    while cursor <= end_seconds:
        window_end = min(cursor + window_seconds, end_seconds)

        try:
            payload = api_get(url, {
                "from": cursor,
                "to": window_end,
            })

            for row in payload.get("candles", []):
                try:
                    stamp = int(float(row["time"]))
                    if stamp < 100_000_000_000:
                        stamp *= 1000

                    c = Candle(
                        time=stamp,
                        open=float(row["open"]),
                        high=float(row["high"]),
                        low=float(row["low"]),
                        close=float(row["close"]),
                        volume=float(row.get("volume", 0) or 0),
                    )

                    values = (c.open, c.high, c.low, c.close)
                    if not all(math.isfinite(x) for x in values):
                        continue
                    if c.high < max(c.open, c.close, c.low):
                        continue
                    if c.low > min(c.open, c.close, c.high):
                        continue

                    # Only fully closed candles.
                    if c.time + interval_ms <= end_ms:
                        result[c.time] = c

                except (KeyError, ValueError, TypeError):
                    continue

        except Exception as exc:
            errors.append(f"{resolution} {cursor}-{window_end}: {exc}")

        cursor = window_end + 1

    candles = [result[k] for k in sorted(result)]
    return candles, errors


# ============================================================
# INDICATOR HELPERS
# ============================================================

def sma(values, length):
    out = [None] * len(values)
    for i in range(length - 1, len(values)):
        w = values[i - length + 1:i + 1]
        if all(x is not None for x in w):
            out[i] = sum(w) / length
    return out


def rolling_extreme(values, length, highest=True):
    out = [None] * len(values)
    fn = max if highest else min
    for i in range(length - 1, len(values)):
        out[i] = fn(values[i - length + 1:i + 1])
    return out


def linreg_endpoint(values, length):
    out = [None] * len(values)
    xm = (length - 1) / 2
    den = sum((x - xm) ** 2 for x in range(length))

    for i in range(length - 1, len(values)):
        w = values[i - length + 1:i + 1]
        if any(x is None for x in w):
            continue
        ym = sum(w) / length
        slope = sum(
            (x - xm) * (w[x] - ym) for x in range(length)
        ) / den
        out[i] = ym + slope * (length - 1 - xm)

    return out


def sqzmom_values(candles):
    highs = [c.high for c in candles]
    lows = [c.low for c in candles]
    closes = [c.close for c in candles]

    hh = rolling_extreme(highs, KC_LENGTH, True)
    ll = rolling_extreme(lows, KC_LENGTH, False)
    av = sma(closes, KC_LENGTH)

    source = [None] * len(candles)
    for i in range(len(candles)):
        if hh[i] is None or ll[i] is None or av[i] is None:
            continue
        midpoint = ((hh[i] + ll[i]) / 2 + av[i]) / 2
        source[i] = closes[i] - midpoint

    return linreg_endpoint(source, KC_LENGTH)


def sqzmom_colors(candles):
    values = sqzmom_values(candles)
    colors = [None] * len(values)

    for i, value in enumerate(values):
        if value is None:
            continue

        previous = values[i - 1] if i > 0 else 0.0
        previous = 0.0 if previous is None else previous

        if value > 0:
            colors[i] = "LIGHT_GREEN" if value > previous else "DARK_GREEN"
        else:
            colors[i] = "LIGHT_RED" if value < previous else "DARK_RED"

    return colors


BUY_PATTERNS = {
    ("LIGHT_RED", "DARK_RED", "DARK_RED"),
    ("DARK_GREEN", "LIGHT_GREEN", "LIGHT_GREEN"),
}

SELL_PATTERNS = {
    ("LIGHT_GREEN", "DARK_GREEN", "DARK_GREEN"),
    ("DARK_RED", "LIGHT_RED", "LIGHT_RED"),
}


def sqz_direction(pattern):
    p = tuple(pattern)
    if p in BUY_PATTERNS:
        return "BUY"
    if p in SELL_PATTERNS:
        return "SELL"
    return None


def ema(values, length):
    out = [None] * len(values)
    if not values:
        return out

    alpha = 2 / (length + 1)
    current = values[0]
    out[0] = current

    for i in range(1, len(values)):
        current = alpha * values[i] + (1 - alpha) * current
        out[i] = current

    return out


def wma(values, length):
    out = [None] * len(values)
    den = length * (length + 1) / 2

    for i in range(length - 1, len(values)):
        w = values[i - length + 1:i + 1]
        if any(x is None for x in w):
            continue
        out[i] = sum((j + 1) * w[j] for j in range(length)) / den

    return out


def rma(values, length):
    out = [None] * len(values)
    if len(values) < length:
        return out

    current = sum(values[:length]) / length
    out[length - 1] = current

    for i in range(length, len(values)):
        current = (current * (length - 1) + values[i]) / length
        out[i] = current

    return out


def pure_momentum_colors(candles):
    n = len(candles)
    if not n:
        return []

    raw = {
        "open": [c.open for c in candles],
        "high": [c.high for c in candles],
        "low": [c.low for c in candles],
        "close": [c.close for c in candles],
    }

    doubled = {}
    for key, values in raw.items():
        smooth = ema(values, PMO_LENGTH)
        doubled[key] = [
            values[i] * 2 - smooth[i] for i in range(n)
        ]

    cda = wma(doubled["close"], PMO_LENGTH)

    tr = []
    for i, c in enumerate(candles):
        if i == 0:
            tr.append(c.high - c.low)
        else:
            pc = candles[i - 1].close
            tr.append(max(
                c.high - c.low,
                abs(c.high - pc),
                abs(c.low - pc),
            ))

    atr = rma(tr, PMO_LENGTH)

    trend = 0
    colors = [None] * n

    for i in range(n):
        change = None
        if (
            i > 0
            and cda[i] is not None
            and cda[i - 1] is not None
            and atr[i] not in (None, 0)
        ):
            change = (cda[i] - cda[i - 1]) / atr[i]

        if change is not None:
            if change > PMO_UP:
                trend = 1
            elif change < PMO_DOWN:
                trend = -1

        colors[i] = "GREEN" if trend == 1 else "RED"

    return colors


# ============================================================
# SCENARIOS
# ============================================================

def make_scenarios():
    scenarios = []

    # 36 fixed TP scenarios:
    for margin in MARGIN_FRACTIONS:
        for buffer in SL_BUFFERS:
            for rr in RR_VALUES:
                scenarios.append(Scenario(
                    name=f"FIXED_M{margin}_B{buffer}_RR{rr}",
                    margin_fraction=margin,
                    buffer=buffer,
                    exit_type="FIXED",
                    rr=float(rr),
                ))

    # 18 floating-exit scenarios:
    floating = (
        "OPPOSITE_PMO",
        "OPPOSITE_SQZ",
        "STRUCTURE_TRAIL",
    )

    for margin in MARGIN_FRACTIONS:
        for buffer in SL_BUFFERS:
            for exit_type in floating:
                scenarios.append(Scenario(
                    name=f"{exit_type}_M{margin}_B{buffer}",
                    margin_fraction=margin,
                    buffer=buffer,
                    exit_type=exit_type,
                ))

    assert len(scenarios) == 54
    return scenarios


# ============================================================
# SIGNAL TIMELINE
# ============================================================

def build_events(c5, c15):
    """
    At each 5m open:
    - use the previous three CLOSED 5m candles;
    - use the latest CLOSED 15m PMO candle;
    - signal becomes executable at this 5m open.
    """
    colors5 = sqzmom_colors(c5)
    colors15 = pure_momentum_colors(c15)

    events = {}
    pmo_index = -1

    for i in range(3, len(c5)):
        entry_time = c5[i].time

        while (
            pmo_index + 1 < len(c15)
            and c15[pmo_index + 1].time + PMO_MS <= entry_time
        ):
            pmo_index += 1

        if pmo_index < 0:
            continue

        pattern = colors5[i - 3:i]
        if len(pattern) != 3 or any(x is None for x in pattern):
            continue

        direction = sqz_direction(pattern)
        pmo_color = colors15[pmo_index]
        pmo_direction = (
            "BUY" if pmo_color == "GREEN"
            else "SELL" if pmo_color == "RED"
            else None
        )

        if direction and direction == pmo_direction:
            events[entry_time] = {
                "direction": direction,
                "pattern": tuple(pattern),
                "pmo_index": pmo_index,
                "sqz_index": i - 1,
            }

    return events, colors5, colors15


# ============================================================
# STOP / PIVOT HELPERS
# ============================================================

def initial_stop(c15, entry_time, direction, buffer):
    closed = [
        c for c in c15
        if c.time + PMO_MS <= entry_time
    ]

    if len(closed) < 10:
        return None

    prior = closed[-10:]

    if direction == "BUY":
        return min(c.low for c in prior) * (1 - buffer)

    return max(c.high for c in prior) * (1 + buffer)


def confirmed_pivots(c15, left=2, right=2):
    """
    A pivot is usable only after its right-side candles have closed.
    Returns (confirmation_time, pivot_low, pivot_high).
    """
    result = []

    for i in range(left, len(c15) - right):
        window = c15[i - left:i + right + 1]
        pivot = c15[i]

        is_low = all(
            pivot.low < x.low
            for j, x in enumerate(window)
            if j != left
        )
        is_high = all(
            pivot.high > x.high
            for j, x in enumerate(window)
            if j != left
        )

        if is_low or is_high:
            confirmation_time = c15[i + right].time + PMO_MS
            result.append((
                confirmation_time,
                pivot.low if is_low else None,
                pivot.high if is_high else None,
            ))

    return result


# ============================================================
# TRADE ENGINE
# ============================================================

def adverse_entry(price, direction):
    return price * (
        1 + SLIPPAGE_RATE if direction == "BUY"
        else 1 - SLIPPAGE_RATE
    )


def adverse_exit(price, direction):
    # direction is the position direction, not the exit order side.
    return price * (
        1 - SLIPPAGE_RATE if direction == "BUY"
        else 1 + SLIPPAGE_RATE
    )


def run_scenario(symbol, c5, c15, events, colors15, scenario):
    equity = INITIAL_EQUITY
    peak_equity = equity
    max_drawdown = 0.0

    position = None
    trades = []
    skipped = 0

    pivots = confirmed_pivots(c15)
    pivot_cursor = 0

    # Map PMO candle close times to their indices.
    pmo_closed_at = [
        c.time + PMO_MS for c in c15
    ]

    for i, candle in enumerate(c5):
        now = candle.time

        # Latest PMO candle known at this 5m open.
        pmo_i = -1
        for j, close_time in enumerate(pmo_closed_at):
            if close_time <= now:
                pmo_i = j
            else:
                break

        # Apply newly confirmed structure pivots.
        while pivot_cursor < len(pivots) and pivots[pivot_cursor][0] <= now:
            _, pivot_low, pivot_high = pivots[pivot_cursor]
            pivot_cursor += 1

            if position and scenario.exit_type == "STRUCTURE_TRAIL":
                if position["direction"] == "BUY" and pivot_low is not None:
                    candidate = pivot_low
                    if candidate > position["stop"]:
                        position["stop"] = candidate

                elif position["direction"] == "SELL" and pivot_high is not None:
                    candidate = pivot_high
                    if candidate < position["stop"]:
                        position["stop"] = candidate

        # Manage an existing position.
        if position:
            direction = position["direction"]
            exit_reason = None
            exit_price = None

            # Floating PMO exit: opposite closed PMO candle,
            # executed at the current 5m open.
            if (
                scenario.exit_type == "OPPOSITE_PMO"
                and pmo_i >= 0
            ):
                color = colors15[pmo_i]
                opposite = (
                    direction == "BUY" and color == "RED"
                ) or (
                    direction == "SELL" and color == "GREEN"
                )
                if opposite:
                    exit_reason = "OPPOSITE_PMO"
                    exit_price = candle.open

            # Floating SQZMOM exit: confirmed opposite 3-bar pattern
            # from the last three closed 5m candles, next 5m open.
            if (
                exit_reason is None
                and scenario.exit_type == "OPPOSITE_SQZ"
                and i >= 3
            ):
                recent = sqzmom_direction(
                    [position["colors5"][i - 3],
                     position["colors5"][i - 2],
                     position["colors5"][i - 1]]
                )
                if recent and recent != direction:
                    exit_reason = "OPPOSITE_SQZ"
                    exit_price = candle.open

            # Intrabar stop/target checks. If both hit, stop wins.
            if exit_reason is None:
                stop_hit = (
                    candle.low <= position["stop"]
                    if direction == "BUY"
                    else candle.high >= position["stop"]
                )

                target_hit = (
                    position["target"] is not None
                    and (
                        candle.high >= position["target"]
                        if direction == "BUY"
                        else candle.low <= position["target"]
                    )
                )

                if stop_hit:
                    exit_reason = "STOP"
                    exit_price = position["stop"]
                elif target_hit:
                    exit_reason = "TAKE_PROFIT"
                    exit_price = position["target"]

            if exit_reason:
                fill = adverse_exit(exit_price, direction)
                qty = position["qty"]

                gross = (
                    (fill - position["entry"]) * qty
                    if direction == "BUY"
                    else (position["entry"] - fill) * qty
                )

                exit_fee = abs(fill * qty) * FEE_RATE
                net = gross - exit_fee
                equity += net

                trades.append({
                    "symbol": symbol,
                    "scenario": scenario.name,
                    "direction": direction,
                    "entry_time": position["entry_time"],
                    "exit_time": now,
                    "entry_price": position["entry"],
                    "exit_price": fill,
                    "stop_initial": position["initial_stop"],
                    "stop_final": position["stop"],
                    "target": position["target"],
                    "margin": position["margin"],
                    "notional": position["notional"],
                    "qty": qty,
                    "gross_pnl": gross,
                    "entry_fee": position["entry_fee"],
                    "exit_fee": exit_fee,
                    "net_pnl": net - position["entry_fee"],
                    "equity_after": equity,
                    "exit_reason": exit_reason,
                })

                position = None

                peak_equity = max(peak_equity, equity)
                if peak_equity > 0:
                    max_drawdown = max(
                        max_drawdown,
                        (peak_equity - equity) / peak_equity
                    )

        # One open trade per symbol per scenario.
        if position is not None:
            continue

        event = events.get(now)
        if not event:
            continue

        direction = event["direction"]
        stop = initial_stop(
            c15, now, direction, scenario.buffer
        )

        if stop is None:
            skipped += 1
            continue

        raw_entry = candle.open
        entry = adverse_entry(raw_entry, direction)

        if direction == "BUY" and stop >= entry:
            skipped += 1
            continue
        if direction == "SELL" and stop <= entry:
            skipped += 1
            continue

        margin = equity * scenario.margin_fraction
        notional = margin * LEVERAGE

        if margin <= 0 or notional <= 0:
            skipped += 1
            continue

        qty = notional / entry
        entry_fee = notional * FEE_RATE
        equity -= entry_fee

        risk = abs(entry - stop)
        target = None

        if scenario.exit_type == "FIXED":
            if direction == "BUY":
                target = entry + risk * scenario.rr
            else:
                target = entry - risk * scenario.rr

        position = {
            "direction": direction,
            "entry_time": now,
            "entry": entry,
            "initial_stop": stop,
            "stop": stop,
            "target": target,
            "margin": margin,
            "notional": notional,
            "qty": qty,
            "entry_fee": entry_fee,
            "colors5": sqzmom_colors(c5),
        }

    # Close any remaining trade at the last available close.
    if position and c5:
        last = c5[-1]
        direction = position["direction"]
        fill = adverse_exit(last.close, direction)
        qty = position["qty"]

        gross = (
            (fill - position["entry"]) * qty
            if direction == "BUY"
            else (position["entry"] - fill) * qty
        )

        exit_fee = abs(fill * qty) * FEE_RATE
        net = gross - exit_fee
        equity += net

        trades.append({
            "symbol": symbol,
            "scenario": scenario.name,
            "direction": direction,
            "entry_time": position["entry_time"],
            "exit_time": last.time + SQZ_MS,
            "entry_price": position["entry"],
            "exit_price": fill,
            "stop_initial": position["initial_stop"],
            "stop_final": position["stop"],
            "target": position["target"],
            "margin": position["margin"],
            "notional": position["notional"],
            "qty": qty,
            "gross_pnl": gross,
            "entry_fee": position["entry_fee"],
            "exit_fee": exit_fee,
            "net_pnl": net - position["entry_fee"],
            "equity_after": equity,
            "exit_reason": "END_OF_DATA_MARK_TO_MARKET",
        })

    wins = sum(1 for t in trades if t["net_pnl"] > 0)
    losses = sum(1 for t in trades if t["net_pnl"] < 0)
    total = len(trades)

    summary = {
        "symbol": symbol,
        "scenario": scenario.name,
        "exit_type": scenario.exit_type,
        "margin_fraction": scenario.margin_fraction,
        "sl_buffer": scenario.buffer,
        "rr": scenario.rr,
        "initial_equity": INITIAL_EQUITY,
        "final_equity": equity,
        "net_profit": equity - INITIAL_EQUITY,
        "return_pct": (equity / INITIAL_EQUITY - 1) * 100,
        "trades": total,
        "wins": wins,
        "losses": losses,
        "win_rate_pct": wins / total * 100 if total else 0,
        "max_drawdown_pct": max_drawdown * 100,
        "skipped_entries": skipped,
    }

    return summary, trades


# ============================================================
# DATA QUALITY
# ============================================================

def quality_report(symbol, timeframe, candles, interval_ms, errors):
    gaps = []
    for i in range(1, len(candles)):
        delta = candles[i].time - candles[i - 1].time
        if delta != interval_ms:
            gaps.append({
                "symbol": symbol,
                "timeframe": timeframe,
                "previous_time": candles[i - 1].time,
                "current_time": candles[i].time,
                "gap_ms": delta,
                "missing_intervals": max(0, delta // interval_ms - 1),
            })

    return {
        "symbol": symbol,
        "timeframe": timeframe,
        "candles": len(candles),
        "gaps": len(gaps),
        "api_errors": len(errors),
        "error_details": " | ".join(errors[:20]),
    }, gaps


# ============================================================
# MAIN
# ============================================================

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    now = datetime.now(timezone.utc)
    test_start = now - timedelta(days=DAYS)

    # Warmup data is fetched before the test window.
    warmup_start = test_start - timedelta(days=10)

    start_ms = int(warmup_start.timestamp() * 1000)
    end_ms = int(now.timestamp() * 1000)

    selected, ticker_map = get_market_metadata()
    scenarios = make_scenarios()

    all_summaries = []
    all_trades = []
    all_quality = []
    all_gaps = []
    selected_rows = []

    for rank, contract in enumerate(selected, start=1):
        symbol = contract["symbol"]
        LOG.info("Processing %s (%s/%s)", symbol, rank, len(selected))

        c5, e5 = fetch_candles(
            symbol, "5m", SQZ_MS, start_ms, end_ms
        )
        c15, e15 = fetch_candles(
            symbol, "15m", PMO_MS, start_ms, end_ms
        )

        q5, g5 = quality_report(symbol, "5m", c5, SQZ_MS, e5)
        q15, g15 = quality_report(symbol, "15m", c15, PMO_MS, e15)

        all_quality.extend([q5, q15])
        all_gaps.extend(g5 + g15)
        selected_rows.append(contract)

        if len(c5) < 200 or len(c15) < 100:
            LOG.error("%s: insufficient historical data; skipped", symbol)
            continue

        # Only entries inside the 30-day test period are allowed.
        events, colors5, colors15 = build_events(c5, c15)
        events = {
            stamp: event
            for stamp, event in events.items()
            if stamp >= int(test_start.timestamp() * 1000)
        }

        # Restrict price data to the test period plus warmup already used
        # for indicator calculations.
        test_c5 = [
            c for c in c5
            if c.time >= int(test_start.timestamp() * 1000)
        ]

        if not test_c5:
            LOG.error("%s: no candles in test window", symbol)
            continue

        # Keep the full 15m history for the initial 10-bar stop and
        # indicator state. The engine only enters on test-window events.
        for scenario in scenarios:
            summary, trades = run_scenario(
                symbol,
                c5,
                c15,
                events,
                colors15,
                scenario,
            )
            all_summaries.append(summary)
            all_trades.extend(trades)

    assumptions = [
        {"item": "Exchange", "value": "Kraken Futures"},
        {"item": "Historical window", "value": "30 days"},
        {"item": "Contracts", "value": "Top 10 eligible linear PF perpetuals"},
        {"item": "Signal", "value": "5m SQZMOM + closed 15m Pure Momentum"},
        {"item": "BTC filter", "value": "None; every symbol is independent"},
        {"item": "Scenarios", "value": "54 per symbol"},
        {"item": "Initial equity", "value": "$100 per symbol and scenario"},
        {"item": "Leverage", "value": "10x"},
        {"item": "Margin fractions", "value": "1% and 10% of current equity"},
        {"item": "SL buffers", "value": "0%, 0.1%, 0.2%"},
        {"item": "Fixed RR", "value": "1, 2, 3, 4, 8, 10"},
        {"item": "Fee", "value": f"{FEE_RATE * 100:.4f}% per side (assumption)"},
        {"item": "Slippage", "value": f"{SLIPPAGE_RATE * 100:.4f}% per side (assumption)"},
        {"item": "Funding", "value": "0; historical funding not modeled"},
        {"item": "Same-candle SL and TP", "value": "Stop first"},
        {"item": "Open final position", "value": "Marked to final close"},
        {"item": "Volume selection", "value": "vol24h ¡Á last for eligible linear PF contracts"},
        {"item": "Contract limitation", "value": "Inverse/quanto excluded; not approximated"},
        {"item": "Important", "value": "No results are fabricated if API data is missing"},
    ]

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        pd.DataFrame(all_summaries).to_excel(
            writer, sheet_name="Summary", index=False
        )
        pd.DataFrame(all_trades).to_excel(
            writer, sheet_name="Trade Ledger", index=False
        )
        pd.DataFrame(all_quality).to_excel(
            writer, sheet_name="Data Quality", index=False
        )
        pd.DataFrame(all_gaps).to_excel(
            writer, sheet_name="Data Gaps", index=False
        )
        pd.DataFrame(selected_rows).to_excel(
            writer, sheet_name="Selected Contracts", index=False
        )
        pd.DataFrame(assumptions).to_excel(
            writer, sheet_name="Assumptions", index=False
        )

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for column in ws.columns:
                max_len = 0
                col_letter = get_column_letter(column[0].column)
                for cell in column[:5000]:
                    try:
                        max_len = max(max_len, len(str(cell.value or "")))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

    LOG.info("Finished. Report: %s", OUTPUT.resolve())
    LOG.info(
        "Contracts=%s | summaries=%s | trades=%s",
        len(selected_rows), len(all_summaries), len(all_trades)
    )


if __name__ == "__main__":
    main()

