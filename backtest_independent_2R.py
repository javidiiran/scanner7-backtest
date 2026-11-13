#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Independent candle-by-candle BTC backtest: SQZMOM 5m + Pure Momentum 15m.
No scanner imports, scanner signals, scanner state, or scanner scheduling.
"""
from __future__ import annotations
import math, time, logging, sys
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
import requests
import pandas as pd
from openpyxl.utils import get_column_letter

BASE = "https://futures.kraken.com"
CHARTS = f"{BASE}/api/charts/v1"
INSTRUMENTS = f"{BASE}/derivatives/api/v3/instruments"
TICKERS = f"{BASE}/derivatives/api/v3/tickers"
SYMBOL = "PF_XBTUSD"
DAYS, WARMUP_DAYS = 30, 10
INITIAL_EQUITY = 100.0
LEVERAGE = 50.0
MARGIN_FRACTION = 0.10             # 10% of currently FREE wallet balance
STOP_BUFFER = 0.001                # 0.1% beyond prior 10 closed 15m candle extreme
REWARD_RISK = 2.0
FEE_RATE = 0.0005                  # estimated taker fee each side
SLIPPAGE_RATE = 0.0002             # estimated adverse slippage each side
MAINTENANCE_MARGIN_RATE = 0.005    # generic estimate; not exchange liquidation engine
LIQUIDATION_FEE_RATE = 0.005       # illustrative estimate
SQZ_MS, PMO_MS = 300_000, 900_000
KC_LENGTH, PMO_LENGTH = 20, 45
PMO_UP, PMO_DOWN = 0.05, -0.05
OUTPUT = Path("independent_backtest_report_2R.xlsx")
LOG = logging.getLogger("independent_candle_backtest")
SESSION = requests.Session()
SESSION.headers.update({"User-Agent":"Independent-Candle-Backtest/1.0", "Accept":"application/json"})

@dataclass(frozen=True)
class Candle:
    time: int
    open: float
    high: float
    low: float
    close: float
    volume: float = 0.0

@dataclass
class Position:
    trade_id: int
    direction: str
    signal_time: int
    entry_time: int
    entry: float
    stop: float
    target: float
    margin: float
    notional: float
    qty: float
    entry_fee: float
    liq_price: float
    pattern: str
    pmo_color: str
    funding_pnl: float = 0.0
    contracts: float = 0.0
    contract_size: float = 1.0


def num(v):
    try:
        x = float(v)
        return x if math.isfinite(x) else None
    except (TypeError, ValueError):
        return None


def api_get(url, params=None, attempts=5):
    err = None
    for k in range(attempts):
        try:
            r = SESSION.get(url, params=params, timeout=35)
            if r.status_code in (429, 500, 502, 503, 504):
                raise RuntimeError(f"HTTP {r.status_code}")
            r.raise_for_status()
            d = r.json()
            if not isinstance(d, dict) or d.get("result") == "error":
                raise RuntimeError(f"Bad API payload: {d}")
            return d
        except Exception as e:
            err = e
            if k + 1 < attempts:
                time.sleep(1.5 * (k + 1))
    raise RuntimeError(f"Kraken API failed: {err}")


def validate_market():
    instruments = api_get(INSTRUMENTS).get("instruments", [])
    tickers = api_get(TICKERS).get("tickers", [])
    inst = next((x for x in instruments if str(x.get("symbol", "")).upper() == SYMBOL), None)
    tick = next((x for x in tickers if str(x.get("symbol", "")).upper() == SYMBOL), None)
    if not inst or not tick:
        raise RuntimeError(f"Required Kraken market {SYMBOL} or its ticker is unavailable")
    if inst.get("tradeable") is False or inst.get("suspended") is True:
        raise RuntimeError(f"{SYMBOL} is not currently tradeable")
    contract_size = num(inst.get("contractSize"))
    if contract_size is None or contract_size <= 0:
        raise RuntimeError(f"Kraken contractSize metadata is missing/invalid for {SYMBOL}: {inst.get('contractSize')!r}")
    return {"symbol": SYMBOL, "type": inst.get("type"), "contract_size": contract_size,
            "tick_size": inst.get("tickSize"), "quote_currency": inst.get("quoteCurrency"),
            "underlying": inst.get("underlying"), "last_price": num(tick.get("last"))}


def fetch_candles(resolution, interval_ms, start_ms, end_ms, tick_type="trade"):
    if tick_type not in ("trade", "mark"):
        raise ValueError("tick_type must be trade or mark")
    url = f"{CHARTS}/{tick_type}/{SYMBOL}/{resolution}"
    found, errors = {}, []
    cursor, end_s, window = start_ms // 1000, end_ms // 1000, 3 * 86400
    while cursor <= end_s:
        stop = min(cursor + window, end_s)
        try:
            payload = api_get(url, {"from": cursor, "to": stop})
            rows = payload.get("candles")
            if not isinstance(rows, list):
                raise RuntimeError("Missing candles array")
            for row in rows:
                try:
                    ts = int(float(row["time"]))
                    if ts < 100_000_000_000:
                        ts *= 1000
                    c = Candle(ts, float(row["open"]), float(row["high"]), float(row["low"]),
                               float(row["close"]), float(row.get("volume", 0) or 0))
                    vals = (c.open, c.high, c.low, c.close)
                    if not all(math.isfinite(v) and v > 0 for v in vals):
                        raise ValueError("invalid OHLC")
                    if c.high < max(c.open, c.close, c.low) or c.low > min(c.open, c.close, c.high):
                        raise ValueError("invalid OHLC range")
                    if c.time + interval_ms <= end_ms:
                        found[c.time] = c
                except (KeyError, TypeError, ValueError, OverflowError) as e:
                    errors.append(f"malformed {resolution} candle: {e}")
        except Exception as e:
            errors.append(f"{resolution} {cursor}-{stop}: {e}")
        cursor = stop + 1
    return [found[k] for k in sorted(found)], errors


def fetch_funding_rates(start_ms, end_ms):
    """Fetch Kraken historical relative funding settlements; fail closed on bad data."""
    url = f"{BASE}/derivatives/api/v4/historicalfundingrates"
    payload = api_get(url, {"symbol": SYMBOL})
    rows = payload.get("rates")
    if not isinstance(rows, list) or not rows:
        raise RuntimeError("Kraken historical funding response has no rates array")
    out = []
    malformed = 0
    for row in rows:
        rate = num(row.get("relativeFundingRate"))
        stamp = row.get("timestamp")
        if rate is None or not stamp:
            malformed += 1
            continue
        try:
            ts = int(datetime.fromisoformat(str(stamp).replace("Z", "+00:00")).timestamp() * 1000)
        except Exception:
            malformed += 1
            continue
        if ts % SQZ_MS != 0:
            raise RuntimeError(f"Funding settlement {stamp} is not aligned to a 5m candle boundary; refusing ambiguous timing")
        if start_ms < ts <= end_ms:
            out.append((ts, rate))
    out = sorted(set(out))
    if malformed:
        raise RuntimeError(f"Funding response contains {malformed} malformed settlement record(s)")
    if not out:
        raise RuntimeError("No historical funding settlements found in requested period")
    expected = max(1, int((end_ms - start_ms) / (8 * 3600000)))
    if len(out) < expected * 0.80:
        raise RuntimeError(f"Funding history incomplete: received {len(out)} settlements; expected approximately {expected}")
    # Funding is normally settled every 8 hours. Do not silently accept a large hole.
    gaps = [(b[0] - a[0]) / 3600000 for a, b in zip(out, out[1:])]
    if gaps and max(gaps) > 12.5:
        raise RuntimeError(f"Historical funding has a settlement gap of {max(gaps):.2f} hours")
    return out


def sma(a, n):
    out = [None] * len(a)
    for i in range(n - 1, len(a)):
        w = a[i - n + 1:i + 1]
        if all(x is not None for x in w):
            out[i] = sum(w) / n
    return out


def extrema(a, n, high=True):
    out = [None] * len(a)
    fn = max if high else min
    for i in range(n - 1, len(a)):
        out[i] = fn(a[i - n + 1:i + 1])
    return out


def linreg(a, n):
    out = [None] * len(a)
    xm = (n - 1) / 2
    den = sum((j - xm) ** 2 for j in range(n))
    for i in range(n - 1, len(a)):
        w = a[i - n + 1:i + 1]
        if any(x is None for x in w):
            continue
        ym = sum(w) / n
        slope = sum((j - xm) * (w[j] - ym) for j in range(n)) / den
        out[i] = ym + slope * (n - 1 - xm)
    return out


def sqzmom_colors(cs):
    """LazyBear-style SQZMOM histogram color states, calculated from OHLC history."""
    highs = extrema([c.high for c in cs], KC_LENGTH, True)
    lows = extrema([c.low for c in cs], KC_LENGTH, False)
    avgs = sma([c.close for c in cs], KC_LENGTH)
    source = [None] * len(cs)
    for i, c in enumerate(cs):
        if highs[i] is not None and lows[i] is not None and avgs[i] is not None:
            source[i] = c.close - ((highs[i] + lows[i]) / 2 + avgs[i]) / 2
    values = linreg(source, KC_LENGTH)
    colors = []
    for i, v in enumerate(values):
        if v is None:
            colors.append(None)
            continue
        prev = values[i - 1] if i and values[i - 1] is not None else 0.0
        colors.append("LIGHT_GREEN" if v > 0 and v > prev else
                      "DARK_GREEN" if v > 0 else
                      "LIGHT_RED" if v < prev else "DARK_RED")
    return colors


BUY_PATTERNS = {
    ("LIGHT_RED", "DARK_RED", "DARK_RED"),
    ("DARK_GREEN", "LIGHT_GREEN", "LIGHT_GREEN"),
}
SELL_PATTERNS = {
    ("LIGHT_GREEN", "DARK_GREEN", "DARK_GREEN"),
    ("DARK_RED", "LIGHT_RED", "LIGHT_RED"),
}


def pattern_direction(pattern):
    p = tuple(pattern)
    if p in BUY_PATTERNS:
        return "BUY"
    if p in SELL_PATTERNS:
        return "SELL"
    return None


def ema(a, n):
    if not a:
        return []
    out = [a[0]]
    alpha = 2 / (n + 1)
    for x in a[1:]:
        out.append(alpha * x + (1 - alpha) * out[-1])
    return out


def wma(a, n):
    out = [None] * len(a)
    den = n * (n + 1) / 2
    for i in range(n - 1, len(a)):
        w = a[i - n + 1:i + 1]
        if all(x is not None for x in w):
            out[i] = sum((j + 1) * w[j] for j in range(n)) / den
    return out


def rma(a, n):
    out = [None] * len(a)
    if len(a) < n:
        return out
    out[n - 1] = sum(a[:n]) / n
    for i in range(n, len(a)):
        out[i] = (out[i - 1] * (n - 1) + a[i]) / n
    return out


def pure_momentum_colors(cs):
    """15m Pure Momentum color-state calculation used by this standalone model."""
    n = len(cs)
    raw = {k: [getattr(c, k) for c in cs] for k in ("open", "high", "low", "close")}
    dbl = {k: [v[i] * 2 - e[i] for i in range(n)]
           for k, v in raw.items() for e in [ema(v, PMO_LENGTH)]}
    cda = wma(dbl["close"], PMO_LENGTH)
    tr = []
    for i, c in enumerate(cs):
        tr.append(c.high - c.low if i == 0 else
                  max(c.high - c.low, abs(c.high - cs[i - 1].close), abs(c.low - cs[i - 1].close)))
    atr = rma(tr, PMO_LENGTH)
    trend, colors = 0, []
    for i in range(n):
        change = ((cda[i] - cda[i - 1]) / atr[i]
                  if i and cda[i] is not None and cda[i - 1] is not None and atr[i] else None)
        if change is not None:
            if change > PMO_UP:
                trend = 1
            elif change < PMO_DOWN:
                trend = -1
        colors.append("GREEN" if trend == 1 else "RED")
    return colors


def build_confirmed_signals(c5, c15):
    """Create an audit event for every SQZMOM pattern, including rejected confirmations."""
    sqz = sqzmom_colors(c5)
    pmo = pure_momentum_colors(c15)
    events, pmo_j = {}, -1
    for i in range(2, len(c5) - 1):
        signal_close = c5[i].time + SQZ_MS
        while pmo_j + 1 < len(c15) and c15[pmo_j + 1].time + PMO_MS <= signal_close:
            pmo_j += 1
        pattern = tuple(sqz[i - 2:i + 1])
        direction = pattern_direction(pattern)
        if direction is None:
            continue
        pmo_color = pmo[pmo_j] if pmo_j >= 0 else None
        confirmed = ((direction == "BUY" and pmo_color == "GREEN") or
                     (direction == "SELL" and pmo_color == "RED"))
        events[i + 1] = {
            "direction": direction, "signal_time": signal_close,
            "pattern": " -> ".join(pattern), "pmo_color": pmo_color or "UNAVAILABLE",
            "pmo_candle_close": c15[pmo_j].time + PMO_MS if pmo_j >= 0 else None,
            "confirmed": confirmed,
            "rejection_reason": "" if confirmed else ("NO_CLOSED_15M_PMO_CANDLE" if pmo_j < 0 else "PURE_MOMENTUM_DIRECTION_MISMATCH"),
        }
    return events, sqz, pmo


def stop_from_previous_10_closed_15m(c15, signal_time, direction):
    closed = [c for c in c15 if c.time + PMO_MS <= signal_time]
    if len(closed) < 10:
        return None
    last10 = closed[-10:]
    if direction == "BUY":
        return min(c.low for c in last10) * (1 - STOP_BUFFER)
    return max(c.high for c in last10) * (1 + STOP_BUFFER)


def adverse_entry(price, direction):
    return price * (1 + SLIPPAGE_RATE if direction == "BUY" else 1 - SLIPPAGE_RATE)


def adverse_exit(price, direction):
    return price * (1 - SLIPPAGE_RATE if direction == "BUY" else 1 + SLIPPAGE_RATE)


def isolated_liquidation_price(entry, qty, margin, direction):
    reserve = entry * qty * (MAINTENANCE_MARGIN_RATE + LIQUIDATION_FEE_RATE)
    loss_capacity = max(0.0, margin - reserve)
    move = loss_capacity / qty
    return max(1e-12, entry - move) if direction == "BUY" else entry + move


def close_position(pos, raw_exit, timestamp, reason, free_balance, time_basis="INTRABAR_TIME_UNKNOWN_RECORDED_AT_CANDLE_CLOSE"):
    fill = adverse_exit(raw_exit, pos.direction)
    sign = 1 if pos.direction == "BUY" else -1
    gross = (fill - pos.entry) * pos.qty * sign
    exit_fee = abs(fill * pos.qty) * FEE_RATE
    liq_fee = abs(fill * pos.qty) * LIQUIDATION_FEE_RATE if reason.startswith("LIQUIDATION") else 0.0
    # Isolated loss cannot consume another position's collateral or free wallet.
    pnl_after_costs = gross - exit_fee - liq_fee + pos.funding_pnl
    if pnl_after_costs < -pos.margin:
        pnl_after_costs = -pos.margin
    returned = max(0.0, pos.margin + pnl_after_costs)
    free_balance += returned
    net = pnl_after_costs - pos.entry_fee
    row = {
        "trade_id": pos.trade_id, "direction": pos.direction,
        "signal_time_utc": datetime.fromtimestamp(pos.signal_time / 1000, timezone.utc).isoformat(),
        "entry_time_utc": datetime.fromtimestamp(pos.entry_time / 1000, timezone.utc).isoformat(),
        "exit_time_utc": datetime.fromtimestamp(timestamp / 1000, timezone.utc).isoformat(),
        "exit_time_basis": time_basis,
        "entry_price": pos.entry, "exit_price": fill, "stop_loss": pos.stop,
        "take_profit": pos.target, "isolated_margin": pos.margin, "notional": pos.notional,
        "quantity_base_estimate": pos.qty, "contracts": pos.contracts, "contract_size": pos.contract_size, "gross_pnl": gross,
        "entry_fee": pos.entry_fee, "exit_fee": exit_fee, "liquidation_fee": liq_fee,
        "funding_pnl": pos.funding_pnl, "net_pnl": net, "free_balance_after": free_balance, "exit_reason": reason,
        "sqzmom_pattern": pos.pattern, "pure_momentum_color": pos.pmo_color,
    }
    return free_balance, row


def run_backtest(c5, c15, mark5, funding_rates, test_start_ms, metadata):
    events, sqz_colors, pmo_colors = build_confirmed_signals(c5, c15)
    free_balance = INITIAL_EQUITY
    open_positions: list[Position] = []
    ledger, signal_log = [], []
    next_id = 1
    mark_by_time = {c.time: c for c in mark5}
    funding_i = 0
    contract_size = float(metadata["contract_size"])
    peak_equity, max_drawdown = INITIAL_EQUITY, 0.0
    for i, candle in enumerate(c5):
        now = candle.time
        # Apply each settlement to positions open before that exact funding timestamp.
        while funding_i < len(funding_rates) and funding_rates[funding_i][0] <= now:
            fts, frate = funding_rates[funding_i]
            mark_ts = ((fts - 1) // SQZ_MS) * SQZ_MS
            funding_candle = mark_by_time.get(mark_ts)
            if funding_candle is None:
                raise RuntimeError(f"Missing preceding 5m mark candle for funding settlement {fts}")
            for fp in open_positions:
                if fp.entry_time < fts:
                    # Positive relative funding: longs pay, shorts receive.
                    payment = -(1 if fp.direction == "BUY" else -1) * fp.contracts * fp.contract_size * frate
                    fp.funding_pnl += payment
                    fp.liq_price = isolated_liquidation_price(fp.entry, fp.qty, max(0.0, fp.margin + fp.funding_pnl), fp.direction)
            funding_i += 1
        mark_candle = mark_by_time.get(now)
        if mark_candle is None:
            raise RuntimeError(f"Missing 5m mark-price candle at {now}")
        # First process all already-open isolated positions against this 5m candle.
        survivors = []
        for pos in open_positions:
            d = pos.direction
            stop_hit = candle.low <= pos.stop if d == "BUY" else candle.high >= pos.stop
            target_hit = candle.high >= pos.target if d == "BUY" else candle.low <= pos.target
            liq_hit = mark_candle.low <= pos.liq_price if d == "BUY" else mark_candle.high >= pos.liq_price
            gap_stop = (candle.open <= pos.stop if d == "BUY" else candle.open >= pos.stop)
            gap_target = (candle.open >= pos.target if d == "BUY" else candle.open <= pos.target)
            gap_liq = (mark_candle.open <= pos.liq_price if d == "BUY" else mark_candle.open >= pos.liq_price)
            reason, raw = None, None
            if gap_liq:
                reason = "LIQUIDATION_GAP_AMBIGUOUS_STOP" if gap_stop else "LIQUIDATION_GAP"
                if gap_target:
                    reason += "_AMBIGUOUS_TAKE_PROFIT"
                raw = mark_candle.open
            elif gap_stop:
                reason, raw = "STOP_GAP", candle.open
            elif gap_target:
                reason, raw = "TAKE_PROFIT_GAP", candle.open
            elif liq_hit and stop_hit:
                reason, raw = "LIQUIDATION_AMBIGUOUS_STOP", pos.liq_price
            elif liq_hit and target_hit:
                reason, raw = "LIQUIDATION_AMBIGUOUS_TAKE_PROFIT", pos.liq_price
            elif stop_hit and target_hit:
                reason, raw = "STOP_LOSS_AMBIGUOUS_TAKE_PROFIT", pos.stop
            elif liq_hit:
                reason, raw = "LIQUIDATION", pos.liq_price
            elif stop_hit:
                reason, raw = "STOP_LOSS", pos.stop
            elif target_hit:
                reason, raw = "TAKE_PROFIT", pos.target
            if reason:
                is_gap = "_GAP" in reason
                exit_ts = now if is_gap else now + SQZ_MS
                basis = "CANDLE_OPEN_GAP" if is_gap else "INTRABAR_TIME_UNKNOWN_RECORDED_AT_CANDLE_CLOSE"
                free_balance, trade = close_position(pos, raw, exit_ts, reason, free_balance, basis)
                ledger.append(trade)
            else:
                survivors.append(pos)
        open_positions = survivors

        # Signals detected at the previous 5m close enter at this candle's open.
        event = events.get(i)
        if event and now >= test_start_ms:
            if now != event["signal_time"]:
                signal_log.append({**event, "entry_bar_time": now, "status": "SKIPPED_NONCONTIGUOUS_ENTRY_BAR",
                                   "reason": "NEXT_5M_OPEN_DOES_NOT_MATCH_SIGNAL_CLOSE"})
            elif not event["confirmed"]:
                signal_log.append({**event, "entry_bar_time": now,
                                   "status": "REJECTED_CONFIRMATION", "reason": event["rejection_reason"]})
                equity = free_balance + sum(pos.margin + pos.funding_pnl + (mark_candle.close - pos.entry) * pos.qty * (1 if pos.direction == "BUY" else -1) for pos in open_positions)
                peak_equity = max(peak_equity, equity)
                if peak_equity > 0:
                    max_drawdown = max(max_drawdown, (peak_equity - equity) / peak_equity)
                continue
            direction = event["direction"]
            stop = stop_from_previous_10_closed_15m(c15, event["signal_time"], direction)
            entry = adverse_entry(candle.open, direction)
            if stop is None or (direction == "BUY" and stop >= entry) or (direction == "SELL" and stop <= entry):
                signal_log.append({**event, "entry_bar_time": now, "status": "SKIPPED_INVALID_STOP_OR_HISTORY"})
            elif free_balance <= 0:
                signal_log.append({**event, "entry_bar_time": now, "status": "SKIPPED_NO_FREE_BALANCE"})
            else:
                margin = free_balance * MARGIN_FRACTION
                notional = margin * LEVERAGE
                contracts = notional / contract_size
                qty = contracts * contract_size / entry
                entry_fee = notional * FEE_RATE
                if margin <= 0 or notional <= 0 or free_balance < margin + entry_fee:
                    signal_log.append({**event, "entry_bar_time": now, "status": "SKIPPED_INSUFFICIENT_FREE_BALANCE"})
                else:
                    free_balance -= margin + entry_fee
                    risk = abs(entry - stop)
                    target = entry + risk * REWARD_RISK if direction == "BUY" else entry - risk * REWARD_RISK
                    pos = Position(next_id, direction, event["signal_time"], now, entry, stop, target,
                                   margin, notional, qty, entry_fee,
                                   isolated_liquidation_price(entry, qty, margin, direction),
                                   event["pattern"], event["pmo_color"], 0.0, contracts, contract_size)
                    open_positions.append(pos)
                    signal_log.append({**event, "entry_bar_time": now, "status": "OPENED", "trade_id": next_id,
                                       "margin": margin, "free_balance_after_entry": free_balance})
                    next_id += 1
                    # Entry bar is checked too, but each new trade is isolated.
                    d = pos.direction
                    stop_hit = candle.low <= pos.stop if d == "BUY" else candle.high >= pos.stop
                    target_hit = candle.high >= pos.target if d == "BUY" else candle.low <= pos.target
                    liq_hit = mark_candle.low <= pos.liq_price if d == "BUY" else mark_candle.high >= pos.liq_price
                    reason, raw = None, None
                    if liq_hit and stop_hit:
                        reason, raw = "LIQUIDATION_AMBIGUOUS_STOP_ENTRY_BAR", pos.liq_price
                    elif liq_hit and target_hit:
                        reason, raw = "LIQUIDATION_AMBIGUOUS_TAKE_PROFIT_ENTRY_BAR", pos.liq_price
                    elif stop_hit and target_hit:
                        reason, raw = "STOP_LOSS_AMBIGUOUS_TAKE_PROFIT_ENTRY_BAR", pos.stop
                    elif liq_hit:
                        reason, raw = "LIQUIDATION_ENTRY_BAR", pos.liq_price
                    elif stop_hit:
                        reason, raw = "STOP_LOSS_ENTRY_BAR", pos.stop
                    elif target_hit:
                        reason, raw = "TAKE_PROFIT_ENTRY_BAR", pos.target
                    if reason:
                        free_balance, trade = close_position(pos, raw, now + SQZ_MS, reason, free_balance,
                                                             "INTRABAR_TIME_UNKNOWN_RECORDED_AT_CANDLE_CLOSE")
                        ledger.append(trade)
                        open_positions.remove(pos)

        # Equity includes each isolated margin plus its own unrealized PnL.
        equity = free_balance
        for pos in open_positions:
            sign = 1 if pos.direction == "BUY" else -1
            unrealized = (mark_candle.close - pos.entry) * pos.qty * sign
            equity += pos.margin + unrealized
        peak_equity = max(peak_equity, equity)
        if peak_equity > 0:
            max_drawdown = max(max_drawdown, (peak_equity - equity) / peak_equity)

    # Close all remaining positions at the final available 5m close (end-of-test mark).
    if c5:
        last = c5[-1]
        last_mark = mark5[-1]
        for pos in list(open_positions):
            free_balance, trade = close_position(pos, last_mark.close, last.time + SQZ_MS,
                                                 "END_OF_DATA_MARK_TO_MARKET", free_balance)
            ledger.append(trade)
            open_positions.remove(pos)
    wins = sum(t["net_pnl"] > 0 for t in ledger)
    losses = sum(t["net_pnl"] < 0 for t in ledger)
    total = len(ledger)
    summary = {
        "symbol": SYMBOL, "period_days": DAYS, "initial_equity": INITIAL_EQUITY,
        "final_equity": free_balance, "net_profit": free_balance - INITIAL_EQUITY,
        "return_pct": (free_balance / INITIAL_EQUITY - 1) * 100,
        "trades": total, "wins": wins, "losses": losses,
        "win_rate_pct": wins / total * 100 if total else 0,
        "max_drawdown_pct": max_drawdown * 100,
        "buy_trades": sum(t["direction"] == "BUY" for t in ledger),
        "sell_trades": sum(t["direction"] == "SELL" for t in ledger),
        "signals_detected": len(signal_log),
        "signals_confirmed": sum(x["status"] == "OPENED" for x in signal_log),
        "signals_rejected_confirmation": sum(x["status"] == "REJECTED_CONFIRMATION" for x in signal_log),
        "signals_skipped": sum(x["status"].startswith("SKIPPED") for x in signal_log),
        "ambiguous_liquidation_stop_exits": sum("LIQUIDATION" in str(t["exit_reason"]) and "AMBIGUOUS_STOP" in str(t["exit_reason"]) for t in ledger),
        "ambiguous_liquidation_take_profit_exits": sum("LIQUIDATION" in str(t["exit_reason"]) and "AMBIGUOUS_TAKE_PROFIT" in str(t["exit_reason"]) for t in ledger),
        "ambiguous_stop_take_profit_exits": sum("STOP_LOSS_AMBIGUOUS_TAKE_PROFIT" in str(t["exit_reason"]) for t in ledger),
    }
    return summary, ledger, signal_log, sqz_colors, pmo_colors


def data_quality(candles, interval, errors, start_ms, end_ms, tf):
    gaps = []
    for a, b in zip(candles, candles[1:]):
        if b.time - a.time != interval and b.time >= start_ms and a.time < end_ms:
            gaps.append({"timeframe": tf, "previous_ms": a.time, "current_ms": b.time,
                         "missing_intervals": max(0, (b.time - a.time) // interval - 1)})
    return {"timeframe": tf, "candles": len(candles), "gaps_in_test_window": len(gaps),
            "api_errors": len(errors), "error_details": " | ".join(errors[:20]),
            "valid": len(errors) == 0 and len(gaps) == 0}, gaps


def write_report(summary, ledger, signal_log, quality_rows, gap_rows, metadata):
    assumptions = [
        ("Engine", "Independent historical candle-by-candle calculation; no scanner code or scanner signals"),
        ("Market", "Kraken Futures BTC perpetual PF_XBTUSD only"),
        ("History", "30 test days + 10 warmup days for indicator initialization"),
        ("SQZMOM", "5m OHLC; calculate histogram colors; scan every closed candle oldest-to-newest"),
        ("Patterns", "BUY: LIGHT_RED>DARK_RED>DARK_RED or DARK_GREEN>LIGHT_GREEN>LIGHT_GREEN; SELL: inverse patterns"),
        ("Confirmation", "Latest 15m Pure Momentum candle whose close time is no later than the 5m signal close"),
        ("Entry timing", "Signal time is the 5m candle close boundary; entry is the immediately following 5m candle open at that same boundary. Non-contiguous next bars are skipped and logged."),
        ("Independence", "Every confirmed signal opens a separate position; BUY/SELL and same-direction trades may overlap; no signal closes another trade"),
        ("Margin", "Isolated per position; no cross-margin and no use of another position's collateral"),
        ("Sizing", "10% of currently free wallet balance for each new isolated margin; reserved margin is not free"),
        ("Leverage", "50x assumed; not a statement of account-specific leverage eligibility"),
        ("Stop loss", "BUY below lowest low / SELL above highest high of previous 10 closed 15m candles, with 0.1% buffer"),
        ("Take profit", "2:1 reward-to-risk based on entry-to-stop distance"),
        ("Fees", f"Estimated taker fee {FEE_RATE*100:.3f}% each side"),
        ("Slippage", f"Estimated adverse slippage {SLIPPAGE_RATE*100:.3f}% each side"),
        ("Liquidation", "Simplified isolated liquidation estimate; mark-price OHLC triggers liquidation; maintenance tier and liquidation fee remain estimates, not Kraken exact engine"),
        ("Funding", "Historical relativeFundingRate settlements applied to positions open at settlement; positive rates debit longs and credit shorts"),
        ("Funding price", "The preceding 5m mark close is required at each settlement; relative rate is applied to the contract USD notional; missing mark data invalidates the run"),
        ("Mark price", "Historical 5m mark OHLC used for liquidation checks and end-of-data mark-to-market; trade OHLC used for SL/TP"),
        ("Contract sizing", "Kraken instrument contractSize metadata is used to derive contract count; base-equivalent quantity is reported"),
        ("Contract metadata", str(metadata)),
        ("Same candle ambiguity", "5m OHLC cannot determine intrabar event order. Conservative priority is applied but explicitly labeled: liquidation before SL/TP if mark-price liquidation and another exit are touched; SL before TP if both trade-price levels are touched. Summary counts ambiguous cases; this is a stress assumption, not observed exchange execution."),
        ("Exit timestamps", "Gap exits use candle-open time. Intrabar exits are timestamped at candle close only as a reporting boundary; exact trigger time is unknown from 5m OHLC."),
    ]
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame(ledger).to_excel(writer, sheet_name="Trade Ledger", index=False)
        pd.DataFrame(signal_log).to_excel(writer, sheet_name="Signal Log", index=False)
        pd.DataFrame(quality_rows).to_excel(writer, sheet_name="Data Quality", index=False)
        pd.DataFrame(gap_rows).to_excel(writer, sheet_name="Data Gaps", index=False)
        pd.DataFrame([{"item": k, "value": v} for k, v in assumptions] +
                     [{"item": "Market metadata", "value": str(metadata)}]).to_excel(writer, sheet_name="Assumptions", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                width = max((len(str(c.value or "")) for c in col[:5000]), default=10)
                ws.column_dimensions[letter].width = min(width + 2, 48)


def self_test():
    global build_confirmed_signals
    assert pattern_direction(("LIGHT_RED", "DARK_RED", "DARK_RED")) == "BUY"
    assert pattern_direction(("DARK_GREEN", "LIGHT_GREEN", "LIGHT_GREEN")) == "BUY"
    assert pattern_direction(("LIGHT_GREEN", "DARK_GREEN", "DARK_GREEN")) == "SELL"
    assert pattern_direction(("DARK_RED", "LIGHT_RED", "LIGHT_RED")) == "SELL"
    # Independent long and short positions coexist; each locks its own 10% of free cash.
    free = 100.0
    m1 = free * MARGIN_FRACTION
    free -= m1
    m2 = free * MARGIN_FRACTION
    free -= m2
    assert abs(m1 - 10.0) < 1e-9 and abs(m2 - 9.0) < 1e-9 and abs(free - 81.0) < 1e-9
    # Both sides of a 5m pattern are retained in the audit stream when PMO rejects.
    assert pattern_direction(("LIGHT_RED", "DARK_RED", "DARK_RED")) == "BUY"
    # A funding debit is charged only to its own isolated position.
    p = Position(1, "BUY", 0, 0, 100.0, 90.0, 140.0, 10.0, 500.0, 5.0,
                 0.25, 70.0, "pattern", "GREEN", -1.0, 500.0, 1.0)
    fb, row = close_position(p, 100.0, 300000, "END_OF_DATA_MARK_TO_MARKET", 80.0)
    assert abs(row["funding_pnl"] + 1.0) < 1e-9
    assert fb < 90.0  # initial free cash plus returned collateral after costs/funding
    # Synthetic candle replay: a BUY and then a SELL coexist; the second signal
    # must not close or reverse the first position. Both are independently settled.
    base = 9_000_000
    c15_test = [Candle(i * PMO_MS, 100.0, 110.0, 90.0, 100.0) for i in range(10)]
    c5_test = [Candle(base + i * SQZ_MS, 100.0, 100.5, 99.5, 100.0) for i in range(3)]
    mark_test = [Candle(c.time, 100.0, 100.5, 99.5, 100.0) for c in c5_test]
    saved_builder = build_confirmed_signals
    try:
        build_confirmed_signals = lambda a, b: ({
            1: {"direction":"BUY", "signal_time":base+SQZ_MS, "pattern":"BUY_TEST", "pmo_color":"GREEN", "confirmed":True, "rejection_reason":""},
            2: {"direction":"SELL", "signal_time":base+2*SQZ_MS, "pattern":"SELL_TEST", "pmo_color":"RED", "confirmed":True, "rejection_reason":""},
        }, [], [])
        sm, tl, sl, _, _ = run_backtest(c5_test, c15_test, mark_test, [], 0, {"contract_size":1.0})
        assert len(sl) == 2 and all(x["status"] == "OPENED" for x in sl)
        assert len(tl) == 2 and {x["direction"] for x in tl} == {"BUY", "SELL"}
    finally:
        build_confirmed_signals = saved_builder
    # No position limit or direction filter is present in the engine.
    assert "scanner" not in run_backtest.__code__.co_names
    print("Independent backtest self-tests: OK")


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    self_test()
    if "--self-test" in sys.argv:
        return
    now = datetime.now(timezone.utc)
    test_start = now - timedelta(days=DAYS)
    warm_start = test_start - timedelta(days=WARMUP_DAYS)
    start_ms, test_ms, end_ms = int(warm_start.timestamp() * 1000), int(test_start.timestamp() * 1000), int(now.timestamp() * 1000)
    metadata = validate_market()
    c5, e5 = fetch_candles("5m", SQZ_MS, start_ms, end_ms)
    c15, e15 = fetch_candles("15m", PMO_MS, start_ms, end_ms)
    mark5, em = fetch_candles("5m", SQZ_MS, start_ms, end_ms, tick_type="mark")
    funding_rates = fetch_funding_rates(test_ms, end_ms)
    q5, g5 = data_quality(c5, SQZ_MS, e5, test_ms, end_ms, "5m")
    q15, g15 = data_quality(c15, PMO_MS, e15, test_ms, end_ms, "15m")
    qm, gm = data_quality(mark5, SQZ_MS, em, test_ms, end_ms, "5m MARK")
    if len(c5) < 200 or len(c15) < 100 or len(mark5) < 200 or not q5["valid"] or not q15["valid"] or not qm["valid"]:
        raise RuntimeError(f"Data quality failed; refusing to publish performance results. 5m={q5}; 15m={q15}; mark={qm}")
    summary, ledger, signals, _, _ = run_backtest(c5, c15, mark5, funding_rates, test_ms, metadata)
    funding_quality = {"timeframe": "funding settlements", "records": len(funding_rates),
                       "first_settlement_utc": datetime.fromtimestamp(funding_rates[0][0] / 1000, timezone.utc).isoformat(),
                       "last_settlement_utc": datetime.fromtimestamp(funding_rates[-1][0] / 1000, timezone.utc).isoformat(),
                       "max_gap_hours": max([(b[0]-a[0])/3600000 for a,b in zip(funding_rates, funding_rates[1:])] or [0]),
                       "valid": True}
    write_report(summary, ledger, signals, [q5, q15, qm, funding_quality], g5 + g15 + gm, metadata)
    LOG.info("Completed: %s | trades=%s | final equity=%.4f", OUTPUT.resolve(), summary["trades"], summary["final_equity"])


if __name__ == "__main__":
    main()
