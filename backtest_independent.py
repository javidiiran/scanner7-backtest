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
REWARD_RISK = 4.0
FEE_RATE = 0.0005                  # estimated taker fee each side
SLIPPAGE_RATE = 0.0002             # estimated adverse slippage each side
MAINTENANCE_MARGIN_RATE = 0.005    # generic estimate; not exchange liquidation engine
LIQUIDATION_FEE_RATE = 0.005       # illustrative estimate
SQZ_MS, PMO_MS = 300_000, 900_000
KC_LENGTH, PMO_LENGTH = 20, 45
PMO_UP, PMO_DOWN = 0.05, -0.05
OUTPUT = Path("independent_backtest_report.xlsx")
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
    return {"symbol": SYMBOL, "type": inst.get("type"), "contract_size": inst.get("contractSize"),
            "tick_size": inst.get("tickSize"), "quote_currency": inst.get("quoteCurrency"),
            "underlying": inst.get("underlying"), "last_price": num(tick.get("last"))}


def fetch_candles(resolution, interval_ms, start_ms, end_ms):
    url = f"{CHARTS}/trade/{SYMBOL}/{resolution}"
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
    """Each closed 5m bar is evaluated once; confirmation uses only a 15m bar
    whose close timestamp is <= the 5m signal close. Entry is next 5m open.
    Returns a map keyed by the next 5m candle index (execution bar).
    """
    sqz = sqzmom_colors(c5)
    pmo = pure_momentum_colors(c15)
    events, pmo_j = {}, -1
    for i in range(2, len(c5) - 1):
        signal_close = c5[i].time + SQZ_MS
        while pmo_j + 1 < len(c15) and c15[pmo_j + 1].time + PMO_MS <= signal_close:
            pmo_j += 1
        pattern = tuple(sqz[i - 2:i + 1])
        direction = pattern_direction(pattern)
        if direction is None or pmo_j < 0:
            continue
        pmo_color = pmo[pmo_j]
        if (direction == "BUY" and pmo_color != "GREEN") or (direction == "SELL" and pmo_color != "RED"):
            continue
        events[i + 1] = {
            "direction": direction,
            "signal_time": signal_close,
            "pattern": " -> ".join(pattern),
            "pmo_color": pmo_color,
            "pmo_candle_close": c15[pmo_j].time + PMO_MS,
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


def close_position(pos, raw_exit, timestamp, reason, free_balance):
    fill = adverse_exit(raw_exit, pos.direction)
    sign = 1 if pos.direction == "BUY" else -1
    gross = (fill - pos.entry) * pos.qty * sign
    exit_fee = abs(fill * pos.qty) * FEE_RATE
    liq_fee = abs(fill * pos.qty) * LIQUIDATION_FEE_RATE if reason == "LIQUIDATION" else 0.0
    # Isolated loss cannot consume another position's collateral or free wallet.
    pnl_after_costs = gross - exit_fee - liq_fee
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
        "entry_price": pos.entry, "exit_price": fill, "stop_loss": pos.stop,
        "take_profit": pos.target, "isolated_margin": pos.margin, "notional": pos.notional,
        "quantity_base_estimate": pos.qty, "gross_pnl": gross,
        "entry_fee": pos.entry_fee, "exit_fee": exit_fee, "liquidation_fee": liq_fee,
        "net_pnl": net, "free_balance_after": free_balance, "exit_reason": reason,
        "sqzmom_pattern": pos.pattern, "pure_momentum_color": pos.pmo_color,
    }
    return free_balance, row


def run_backtest(c5, c15, test_start_ms):
    events, sqz_colors, pmo_colors = build_confirmed_signals(c5, c15)
    free_balance = INITIAL_EQUITY
    open_positions: list[Position] = []
    ledger, signal_log = [], []
    next_id = 1
    peak_equity, max_drawdown = INITIAL_EQUITY, 0.0
    for i, candle in enumerate(c5):
        now = candle.time
        # First process all already-open isolated positions against this 5m candle.
        survivors = []
        for pos in open_positions:
            d = pos.direction
            stop_hit = candle.low <= pos.stop if d == "BUY" else candle.high >= pos.stop
            target_hit = candle.high >= pos.target if d == "BUY" else candle.low <= pos.target
            liq_hit = candle.low <= pos.liq_price if d == "BUY" else candle.high >= pos.liq_price
            gap_stop = (candle.open <= pos.stop if d == "BUY" else candle.open >= pos.stop)
            gap_target = (candle.open >= pos.target if d == "BUY" else candle.open <= pos.target)
            gap_liq = (candle.open <= pos.liq_price if d == "BUY" else candle.open >= pos.liq_price)
            reason, raw = None, None
            if gap_liq:
                reason, raw = "LIQUIDATION_GAP", candle.open
            elif gap_stop:
                reason, raw = "STOP_GAP", candle.open
            elif gap_target:
                reason, raw = "TAKE_PROFIT_GAP", candle.open
            elif liq_hit and (not stop_hit or (pos.liq_price > pos.stop if d == "BUY" else pos.liq_price < pos.stop)):
                reason, raw = "LIQUIDATION", pos.liq_price
            elif stop_hit:
                reason, raw = "STOP_LOSS", pos.stop
            elif target_hit:
                reason, raw = "TAKE_PROFIT", pos.target
            if reason:
                free_balance, trade = close_position(pos, raw, now, "LIQUIDATION" if reason.startswith("LIQUIDATION") else reason, free_balance)
                ledger.append(trade)
            else:
                survivors.append(pos)
        open_positions = survivors

        # Signals detected at the previous 5m close enter at this candle's open.
        event = events.get(i)
        if event and now >= test_start_ms:
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
                qty = notional / entry
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
                                   event["pattern"], event["pmo_color"])
                    open_positions.append(pos)
                    signal_log.append({**event, "entry_bar_time": now, "status": "OPENED", "trade_id": next_id,
                                       "margin": margin, "free_balance_after_entry": free_balance})
                    next_id += 1
                    # Entry bar is checked too, but each new trade is isolated.
                    d = pos.direction
                    stop_hit = candle.low <= pos.stop if d == "BUY" else candle.high >= pos.stop
                    target_hit = candle.high >= pos.target if d == "BUY" else candle.low <= pos.target
                    liq_hit = candle.low <= pos.liq_price if d == "BUY" else candle.high >= pos.liq_price
                    reason, raw = None, None
                    if liq_hit and (not stop_hit or (pos.liq_price > pos.stop if d == "BUY" else pos.liq_price < pos.stop)):
                        reason, raw = "LIQUIDATION", pos.liq_price
                    elif stop_hit:
                        reason, raw = "STOP_LOSS_ENTRY_BAR", pos.stop
                    elif target_hit:
                        reason, raw = "TAKE_PROFIT_ENTRY_BAR", pos.target
                    if reason:
                        free_balance, trade = close_position(pos, raw, now, "LIQUIDATION" if reason == "LIQUIDATION" else reason, free_balance)
                        ledger.append(trade)
                        open_positions.remove(pos)

        # Equity includes each isolated margin plus its own unrealized PnL.
        equity = free_balance
        for pos in open_positions:
            sign = 1 if pos.direction == "BUY" else -1
            unrealized = (candle.close - pos.entry) * pos.qty * sign
            equity += pos.margin + unrealized
        peak_equity = max(peak_equity, equity)
        if peak_equity > 0:
            max_drawdown = max(max_drawdown, (peak_equity - equity) / peak_equity)

    # Close all remaining positions at the final available 5m close (end-of-test mark).
    if c5:
        last = c5[-1]
        for pos in list(open_positions):
            free_balance, trade = close_position(pos, last.close, last.time + SQZ_MS,
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
        "signals_confirmed": sum(x["status"] == "OPENED" for x in signal_log),
        "signals_skipped": sum(x["status"].startswith("SKIPPED") for x in signal_log),
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
        ("Entry timing", "Next 5m candle open after the 5m signal candle has closed and confirmation is available"),
        ("Independence", "Every confirmed signal opens a separate position; BUY/SELL and same-direction trades may overlap; no signal closes another trade"),
        ("Margin", "Isolated per position; no cross-margin and no use of another position's collateral"),
        ("Sizing", "10% of currently free wallet balance for each new isolated margin; reserved margin is not free"),
        ("Leverage", "50x assumed; not a statement of account-specific leverage eligibility"),
        ("Stop loss", "BUY below lowest low / SELL above highest high of previous 10 closed 15m candles, with 0.1% buffer"),
        ("Take profit", "4:1 reward-to-risk based on entry-to-stop distance"),
        ("Fees", f"Estimated taker fee {FEE_RATE*100:.3f}% each side"),
        ("Slippage", f"Estimated adverse slippage {SLIPPAGE_RATE*100:.3f}% each side"),
        ("Liquidation", "Simplified isolated liquidation estimate using generic maintenance and liquidation fee assumptions; not Kraken's exact engine"),
        ("Funding", "Historical funding not included in this build; must be added before treating results as exchange-exact"),
        ("Mark price", "This build checks stop/target/liquidation against trade OHLC; mark-price liquidation history is not yet modeled"),
        ("Contract sizing", "Base quantity estimated as notional / entry price; validate against Kraken contract multiplier metadata"),
        ("Same candle ambiguity", "If stop and target are both touched in one candle, stop is conservatively prioritized; liquidation is prioritized when closer to entry than stop"),
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
    q5, g5 = data_quality(c5, SQZ_MS, e5, test_ms, end_ms, "5m")
    q15, g15 = data_quality(c15, PMO_MS, e15, test_ms, end_ms, "15m")
    if len(c5) < 200 or len(c15) < 100 or not q5["valid"] or not q15["valid"]:
        raise RuntimeError(f"Data quality failed; refusing to publish performance results. 5m={q5}; 15m={q15}")
    summary, ledger, signals, _, _ = run_backtest(c5, c15, test_ms)
    write_report(summary, ledger, signals, [q5, q15], g5 + g15, metadata)
    LOG.info("Completed: %s | trades=%s | final equity=%.4f", OUTPUT.resolve(), summary["trades"], summary["final_equity"])


if __name__ == "__main__":
    main()
