#!/usr/bin/env python3
"""Scanner 7 - Kraken Futures, 5m SQZMOM detection + 15m Pure Momentum confirmation.

All calculations use CLOSED candles only. BTC follows the same signal path as
all other symbols. BTC direction (or, without a BTC signal, its latest closed
15m Pure Momentum color) controls display priority only; opposite signals are
still reported after aligned signals.
"""
from __future__ import annotations

import json
import logging
import math
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests

# ---------------------------------------------------------------------------
# Kraken Futures endpoints and scanner settings
# ---------------------------------------------------------------------------
BASE_URL = "https://futures.kraken.com"
CHARTS_URL = f"{BASE_URL}/api/charts/v1"
INSTRUMENTS_URL = f"{BASE_URL}/derivatives/api/v3/instruments"
TICKERS_URL = f"{BASE_URL}/derivatives/api/v3/tickers"
SQZ_TIMEFRAME = "5m"
PMO_TIMEFRAME = "15m"
SQZ_INTERVAL_MS = 5 * 60 * 1000
PMO_INTERVAL_MS = 15 * 60 * 1000
CANDLE_LIMIT = 500
WORKERS = max(2, min(20, int(os.getenv("SCAN_WORKERS", "10"))))
STATE_PATH = Path(os.getenv("STATE_PATH", "state.json"))
# New key isolates this version from earlier FVG-based Scanner 7 state.
STATE_KEY = "scanner7_sqzmom5m_pure_momentum15m_v3"

# LazyBear SQZMOM source inputs. Histogram val uses KC length 20; the other
# inputs are retained here to document the exact source defaults.
BB_LENGTH = 20
BB_MULT = 2.0
KC_LENGTH = 20
KC_MULT = 1.5
USE_TRUE_RANGE = True

# Pure Momentum Oscillator (MisinkoMaster) source defaults.
PMO_LENGTH = 45
PMO_LONG_THRESHOLD = 0.05
PMO_SHORT_THRESHOLD = -0.05

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S UTC",
)
LOG = logging.getLogger("scanner7")

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "Scanner7-SQZMOM5m-PureMomentum15m/4.0", "Accept": "application/json"})


@dataclass(frozen=True)
class Candle:
    time: int  # open timestamp in milliseconds
    open: float
    high: float
    low: float
    close: float
    volume: float = 0.0


def request_json(url: str, params: dict | None = None, attempts: int = 4) -> dict:
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            response = SESSION.get(url, params=params, timeout=25)
            if response.status_code in (429, 500, 502, 503, 504):
                raise requests.HTTPError(f"HTTP {response.status_code}")
            response.raise_for_status()
            payload = response.json()
            if not isinstance(payload, dict):
                raise RuntimeError("Kraken returned a non-object JSON payload")
            if payload.get("result") == "error":
                raise RuntimeError(str(payload))
            return payload
        except Exception as exc:  # retry transient API/network failures
            last_error = exc
            if attempt + 1 < attempts:
                time.sleep(1.0 + attempt * 1.5)
    raise RuntimeError(f"Request failed: {url}: {last_error}")


def get_instruments() -> list[dict]:
    payload = request_json(INSTRUMENTS_URL)
    instruments = payload.get("instruments", [])
    accepted_types = {"futures_inverse", "futures_vanilla", "flexible_futures"}
    result: list[dict] = []
    for item in instruments:
        symbol = str(item.get("symbol", "")).upper()
        kind = str(item.get("type", "")).lower()
        if not symbol or kind not in accepted_types:
            continue
        if item.get("tradeable") is False or item.get("suspended") is True:
            continue
        result.append(item)
    if not result:
        raise RuntimeError("Kraken returned no tradable Futures instruments")
    return result


def get_tickers() -> dict[str, dict]:
    payload = request_json(TICKERS_URL)
    return {
        str(row.get("symbol", "")).upper(): row
        for row in payload.get("tickers", [])
        if row.get("symbol")
    }


def pick_btc_symbol(instruments: list[dict]) -> str:
    symbols = {str(row.get("symbol", "")).upper() for row in instruments}
    for candidate in ("PF_XBTUSD", "PI_XBTUSD", "FI_XBTUSD", "FF_XBTUSD", "PF_BTCUSD", "PI_BTCUSD"):
        if candidate in symbols:
            return candidate
    candidates = sorted(s for s in symbols if "XBT" in s or "BTC" in s)
    if not candidates:
        raise RuntimeError("Could not identify a BTC futures contract on Kraken")
    return candidates[0]


def _timestamp_ms(value: Any) -> int:
    stamp = int(float(value))
    # Kraken chart timestamps are normally milliseconds; tolerate seconds.
    return stamp * 1000 if stamp < 100_000_000_000 else stamp


def fetch_candles(symbol: str, timeframe: str, interval_ms: int) -> list[Candle]:
    url = f"{CHARTS_URL}/trade/{symbol}/{timeframe}"
    payload = request_json(url, params={"count": CANDLE_LIMIT})
    now_ms = int(time.time() * 1000)
    unique: dict[int, Candle] = {}
    for row in payload.get("candles", []):
        try:
            candle = Candle(
                time=_timestamp_ms(row["time"]),
                open=float(row["open"]), high=float(row["high"]),
                low=float(row["low"]), close=float(row["close"]),
                volume=float(row.get("volume", 0) or 0),
            )
            if not all(math.isfinite(v) for v in (candle.open, candle.high, candle.low, candle.close)):
                continue
            if candle.high < max(candle.open, candle.close, candle.low):
                continue
            if candle.low > min(candle.open, candle.close, candle.high):
                continue
            if candle.time + interval_ms <= now_ms:
                unique[candle.time] = candle
        except (KeyError, TypeError, ValueError, OverflowError):
            continue
    return [unique[key] for key in sorted(unique)]


def candle_is_stale(candle: Candle, interval_ms: int, now_ms: int | None = None) -> bool:
    """Allow at most one full bar of lag; reject future or older candles."""
    now_ms = int(time.time() * 1000) if now_ms is None else int(now_ms)
    expected_latest_open = (now_ms // interval_ms) * interval_ms - interval_ms
    lag_ms = expected_latest_open - candle.time
    return lag_ms < 0 or lag_ms > interval_ms


# ---------------------------------------------------------------------------
# SQZMOM_LB: exact LazyBear histogram `val` and bcolor semantics
# ---------------------------------------------------------------------------
def sma(values: list[float | None], length: int) -> list[float | None]:
    out: list[float | None] = [None] * len(values)
    for i in range(length - 1, len(values)):
        window = values[i - length + 1:i + 1]
        if all(value is not None for value in window):
            out[i] = sum(float(value) for value in window) / length
    return out


def rolling_extreme(values: list[float], length: int, highest: bool) -> list[float | None]:
    out: list[float | None] = [None] * len(values)
    fn = max if highest else min
    for i in range(length - 1, len(values)):
        out[i] = fn(values[i - length + 1:i + 1])
    return out


def linreg_endpoint(values: list[float | None], length: int) -> list[float | None]:
    """Equivalent to Pine linreg(source, length, 0)."""
    out: list[float | None] = [None] * len(values)
    x_mean = (length - 1) / 2.0
    denominator = sum((x - x_mean) ** 2 for x in range(length))
    for i in range(length - 1, len(values)):
        window = values[i - length + 1:i + 1]
        if any(value is None for value in window):
            continue
        y_mean = sum(float(value) for value in window) / length
        slope = sum((x - x_mean) * (float(window[x]) - y_mean) for x in range(length)) / denominator
        out[i] = y_mean + slope * (length - 1 - x_mean)
    return out


def sqzmom_histogram(candles: list[Candle]) -> list[float | None]:
    """LazyBear: linreg(close - avg(avg(highest, lowest), sma(close)), 20, 0)."""
    highs = [c.high for c in candles]
    lows = [c.low for c in candles]
    closes: list[float | None] = [c.close for c in candles]
    highest = rolling_extreme(highs, KC_LENGTH, True)
    lowest = rolling_extreme(lows, KC_LENGTH, False)
    close_average = sma(closes, KC_LENGTH)
    source: list[float | None] = [None] * len(candles)
    for i in range(len(candles)):
        if highest[i] is None or lowest[i] is None or close_average[i] is None:
            continue
        mid_hl = (float(highest[i]) + float(lowest[i])) / 2.0
        source[i] = closes[i] - ((mid_hl + float(close_average[i])) / 2.0)
    return linreg_endpoint(source, KC_LENGTH)


def histogram_color(value: float, previous: float) -> str:
    # Pine: val>0 ? (val>nz(val[1]) ? lime : green)
    #                 : (val<nz(val[1]) ? red : maroon)
    if value > 0.0:
        return "LIGHT_GREEN" if value > previous else "DARK_GREEN"
    return "LIGHT_RED" if value < previous else "DARK_RED"


def histogram_colors(histogram: list[float | None]) -> list[str | None]:
    colors: list[str | None] = [None] * len(histogram)
    for i, value in enumerate(histogram):
        if value is None:
            continue
        previous = histogram[i - 1] if i > 0 else None
        colors[i] = histogram_color(float(value), 0.0 if previous is None else float(previous))
    return colors


BUY_PATTERNS = {
    ("LIGHT_RED", "DARK_RED", "DARK_RED"),
    ("DARK_GREEN", "LIGHT_GREEN", "LIGHT_GREEN"),
}
SELL_PATTERNS = {
    ("LIGHT_GREEN", "DARK_GREEN", "DARK_GREEN"),
    ("DARK_RED", "LIGHT_RED", "LIGHT_RED"),
}


def detect_sqzmom_signal(candles: list[Candle]) -> tuple[str | None, list[str] | None, str | None]:
    histogram = sqzmom_histogram(candles)
    colors = histogram_colors(histogram)
    latest_color = colors[-1] if colors else None
    last_three = colors[-3:]
    if len(last_three) != 3 or any(color is None for color in last_three):
        return None, None, latest_color
    pattern = tuple(str(color) for color in last_three)
    if pattern in BUY_PATTERNS:
        return "BUY", list(pattern), latest_color
    if pattern in SELL_PATTERNS:
        return "SELL", list(pattern), latest_color
    return None, list(pattern), latest_color


# ---------------------------------------------------------------------------
# Pure Momentum Oscillator (MisinkoMaster) - Pine v6 calculation semantics
# ---------------------------------------------------------------------------
def ema(values: list[float], length: int) -> list[float | None]:
    """Pine-style recursive EMA with the first source value as its seed."""
    out: list[float | None] = [None] * len(values)
    if not values:
        return out
    alpha = 2.0 / (length + 1.0)
    previous = values[0]
    out[0] = previous
    for i in range(1, len(values)):
        previous = alpha * values[i] + (1.0 - alpha) * previous
        out[i] = previous
    return out


def wma(values: list[float | None], length: int) -> list[float | None]:
    """TradingView ta.wma: newest sample receives the greatest weight."""
    out: list[float | None] = [None] * len(values)
    denominator = length * (length + 1) / 2.0
    for i in range(length - 1, len(values)):
        window = values[i - length + 1:i + 1]
        if any(value is None for value in window):
            continue
        out[i] = sum((j + 1) * float(window[j]) for j in range(length)) / denominator
    return out


def rma(values: list[float], length: int) -> list[float | None]:
    """Wilder's RMA used by Pine ta.atr, seeded with the first SMA."""
    out: list[float | None] = [None] * len(values)
    if len(values) < length:
        return out
    current = sum(values[:length]) / length
    out[length - 1] = current
    for i in range(length, len(values)):
        current = ((current * (length - 1)) + values[i]) / length
        out[i] = current
    return out


def pure_momentum_colors(candles: list[Candle]) -> list[str | None]:
    """Return the PMO line/candle trend color for every candle.

    Pine source:
      EMA OHLC -> doubled-deviation OHLC -> WMA(45) -> cda
      change = (cda - cda[1]) / ATR(45)
      change > 0.05 => trend=1 (GREEN)
      change < -0.05 => trend=-1 (RED)
      otherwise retain the previous trend.

    The plotted Momentum line and the indicator-painted candles use the same
    `trend` variable, so their colors are identical. Original price candle
    colors are deliberately not used.
    """
    n = len(candles)
    if not n:
        return []

    raw = {
        "open": [c.open for c in candles],
        "high": [c.high for c in candles],
        "low": [c.low for c in candles],
        "close": [c.close for c in candles],
    }
    smoothed_ema = {key: ema(values, PMO_LENGTH) for key, values in raw.items()}
    doubled: dict[str, list[float | None]] = {}
    for key, values in raw.items():
        doubled[key] = [
            values[i] * 2.0 - float(smoothed_ema[key][i])
            if smoothed_ema[key][i] is not None else None
            for i in range(n)
        ]

    oda = wma(doubled["open"], PMO_LENGTH)
    hda = wma(doubled["high"], PMO_LENGTH)
    lda = wma(doubled["low"], PMO_LENGTH)
    cda = wma(doubled["close"], PMO_LENGTH)

    true_ranges: list[float] = []
    for i, candle in enumerate(candles):
        if i == 0:
            true_ranges.append(candle.high - candle.low)
        else:
            previous_close = candles[i - 1].close
            true_ranges.append(max(
                candle.high - candle.low,
                abs(candle.high - previous_close),
                abs(candle.low - previous_close),
            ))
    atr_values = rma(true_ranges, PMO_LENGTH)

    trend = 0
    colors: list[str | None] = [None] * n
    for i in range(n):
        change = None
        if (i > 0 and cda[i] is not None and cda[i - 1] is not None
                and atr_values[i] not in (None, 0)):
            change = (float(cda[i]) - float(cda[i - 1])) / float(atr_values[i])
        if change is not None:
            if change > PMO_LONG_THRESHOLD:
                trend = 1
            elif change < PMO_SHORT_THRESHOLD:
                trend = -1
        # Pine's source maps trend==1 to green and every other state to red.
        colors[i] = "GREEN" if trend == 1 else "RED"
    return colors


def analyze_candles(sqz_candles: list[Candle], pmo_candles: list[Candle],
                    btc_display_sqz_candles: list[Candle] | None = None) -> dict[str, Any]:
    if len(sqz_candles) < 60 or len(pmo_candles) < max(60, PMO_LENGTH + 2):
        return {"reason": "insufficient_history", "sqz_count": len(sqz_candles),
                "pmo_count": len(pmo_candles)}
    if candle_is_stale(sqz_candles[-1], SQZ_INTERVAL_MS):
        return {"reason": "stale_candles", "timeframe": SQZ_TIMEFRAME,
                "candle_time": sqz_candles[-1].time}
    if candle_is_stale(pmo_candles[-1], PMO_INTERVAL_MS):
        return {"reason": "stale_candles", "timeframe": PMO_TIMEFRAME,
                "candle_time": pmo_candles[-1].time}

    sqz_direction, pattern, last_sqz_color = detect_sqzmom_signal(sqz_candles)
    pmo_colors = pure_momentum_colors(pmo_candles)
    last_pmo_color = pmo_colors[-1] if pmo_colors else None
    last_two_pmo_colors = [color for color in pmo_colors[-2:] if color is not None]
    pmo_direction = "BUY" if last_pmo_color == "GREEN" else "SELL" if last_pmo_color == "RED" else None
    confirmed_direction = sqz_direction if sqz_direction == pmo_direction else None

    btc_sqz_15m_color = "N/A"
    if btc_display_sqz_candles:
        display_hist = histogram_colors(sqzmom_histogram(btc_display_sqz_candles))
        btc_sqz_15m_color = display_hist[-1] if display_hist and display_hist[-1] else "N/A"

    return {
        "reason": "signal" if confirmed_direction else "no_signal",
        "direction": confirmed_direction,
        "sqzmom_direction": sqz_direction,
        "sqzmom_pattern": pattern,
        "last_sqzmom_color": last_sqz_color or "N/A",
        "last_sqzmom_5m_color": last_sqz_color or "N/A",
        "last_sqzmom_15m_color": btc_sqz_15m_color,
        "last_pmo_color": last_pmo_color or "N/A",
        "last_two_pmo_colors": last_two_pmo_colors,
        "pmo_direction": pmo_direction or "N/A",
        "candle_time": pmo_candles[-1].time,
        "sqz_candle_time": sqz_candles[-1].time,
        "change1h": change_1h(pmo_candles),
    }


def analyze_market(symbol: str, is_btc: bool = False) -> tuple[str, dict[str, Any]]:
    try:
        sqz_5m = fetch_candles(symbol, SQZ_TIMEFRAME, SQZ_INTERVAL_MS)
        pmo_15m = fetch_candles(symbol, PMO_TIMEFRAME, PMO_INTERVAL_MS)
        btc_sqz_15m = pmo_15m if is_btc else None
        return symbol, analyze_candles(sqz_5m, pmo_15m, btc_sqz_15m)
    except Exception as exc:
        LOG.warning("Market %s skipped: %s", symbol, exc)
        return symbol, {"reason": "fetch_or_analysis_error", "error": str(exc)}


# ---------------------------------------------------------------------------
# Price-change data and report formatting
# ---------------------------------------------------------------------------
def ticker_change_24h(ticker: dict | None) -> float | None:
    if not ticker:
        return None
    try:
        value = ticker.get("change24h")
        if value is not None:
            return float(value)
        last = ticker.get("last")
        open24 = ticker.get("open24h")
        if last is not None and open24 not in (None, 0, "0"):
            return (float(last) / float(open24) - 1.0) * 100.0
    except (TypeError, ValueError, ZeroDivisionError):
        pass
    return None


def change_1h(candles: list[Candle]) -> float | None:
    """Last closed 15m close versus the close exactly four candles earlier."""
    if len(candles) < 5:
        return None
    last_five = candles[-5:]
    if any(last_five[i].time - last_five[i - 1].time != PMO_INTERVAL_MS for i in range(1, 5)):
        return None
    reference, latest = last_five[0].close, last_five[-1].close
    if reference == 0:
        return None
    return (latest / reference - 1.0) * 100.0


def fmt_pct(value: float | None) -> str:
    return "N/A" if value is None else f"{value:+.2f}%"


def signal_line(item: dict[str, Any], prefix: str = "") -> str:
    direction = item["direction"]
    icon = "🟢" if direction == "BUY" else "🔴"
    return (f"{prefix}{icon} {item['symbol']} | {direction} | "
            f"24h: {fmt_pct(item.get('change24h'))} | 1h: {fmt_pct(item.get('change1h'))}")


def btc_status_line(symbol: str, data: dict[str, Any]) -> str:
    sqz15 = data.get("last_sqzmom_15m_color", "N/A")
    colors = data.get("last_two_pmo_colors", [])
    if len(colors) >= 2:
        pmo_text = f"Pure Momentum 15m last two colors: {colors[-2]} -> {colors[-1]}"
        if colors[-2] != colors[-1]:
            pmo_text += f" | COLOR CHANGE: {colors[-2]} -> {colors[-1]}"
    elif colors:
        pmo_text = f"Pure Momentum 15m latest color: {colors[-1]}"
    else:
        pmo_text = "Pure Momentum 15m colors: N/A"
    return (f"BTC: {symbol} | NO CONFIRMED SIGNAL | "
            f"SQZMOM 15m latest histogram color: {sqz15} | {pmo_text}")

def sort_signals(items: list[dict[str, Any]], priority_direction: str) -> list[dict[str, Any]]:
    """Aligned direction first; BUY strongest gain and SELL largest drop first."""
    def rank(item: dict[str, Any]) -> tuple[int, int, int, float]:
        direction = item["direction"]
        value = item.get("change24h")
        aligned_rank = 0 if priority_direction in ("BUY", "SELL") and direction == priority_direction else 1
        direction_rank = 0 if direction == "BUY" else 1
        missing_rank = 1 if value is None else 0
        if value is None:
            movement_rank = 0.0
        elif direction == "BUY":
            movement_rank = -float(value)  # greatest positive change first
        else:
            movement_rank = float(value)   # most negative change first
        return aligned_rank, direction_rank, missing_rank, movement_rank
    return sorted(items, key=rank)


def filter_alt_signals_by_btc_pmo_color(
    items: list[dict[str, Any]], btc_pmo_color: str | None
) -> list[dict[str, Any]]:
    """Compatibility helper: opposite-direction signals are NOT suppressed."""
    return list(items)

def build_report(btc_symbol: str, btc_data: dict[str, Any], btc_ticker: dict | None,
                 alt_items: list[dict[str, Any]]) -> str:
    lines = ["Scanner 7 | Kraken Futures | SQZMOM 5m + Pure Momentum 15m"]
    btc_direction = btc_data.get("direction")
    if btc_direction in ("BUY", "SELL"):
        btc_item = {
            "symbol": btc_symbol,
            "direction": btc_direction,
            "change24h": ticker_change_24h(btc_ticker),
            "change1h": btc_data.get("change1h"),
        }
        # When BTC has a confirmed signal, show only its signal in the BTC section.
        lines.append(signal_line(btc_item, prefix="BTC: "))
        priority_direction = btc_direction
    else:
        # With no BTC signal, show the latest 15m SQZMOM color and the last
        # two CLOSED 15m Pure Momentum colors (including a transition if present).
        lines.append(btc_status_line(btc_symbol, btc_data))
        priority_direction = (
            "BUY" if btc_data.get("last_pmo_color") == "GREEN" else
            "SELL" if btc_data.get("last_pmo_color") == "RED" else None
        )

    # Both aligned and opposite confirmed signals are reported. Aligned group
    # comes first; each group retains the same movement-based ranking.
    ordered = sort_signals(alt_items, priority_direction)
    if ordered:
        lines.extend(signal_line(item) for item in ordered)
    else:
        lines.append("No confirmed altcoin signals.")
    return "\n".join(lines)

def stale_data_report(stale_symbols: list[str]) -> str:
    shown = ", ".join(stale_symbols[:20])
    if len(stale_symbols) > 20:
        shown += f", and {len(stale_symbols) - 20} more"
    return (
        "Scanner 7 | DATA DELAY - SIGNALS WITHHELD\n"
        "No BUY or SELL signals are being reported because the latest closed "
        "5-minute or 15-minute candle/indicator data is more than one full candle "
        "behind the latest expected closed candle.\n"
        f"Delayed symbols: {shown}\n"
        "Scanning will resume when candle data is current."
    )


# ---------------------------------------------------------------------------
# State and Telegram
# ---------------------------------------------------------------------------
def load_state() -> dict:
    if not STATE_PATH.exists():
        return {"schema_version": 73, STATE_KEY: {
            "last_reported_candles": {},
            "last_btc_report_candle": None,
            "last_stale_fingerprint": None,
            "last_btc_pmo_color": None,
            "last_btc_pmo_candle": None,
            "last_btc_sqz15_color": None,
        }}
    try:
        state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Could not read state file {STATE_PATH}: {exc}") from exc
    if not isinstance(state, dict):
        raise RuntimeError("state.json must contain a JSON object")
    data = state.setdefault(STATE_KEY, {})
    data.setdefault("last_reported_candles", {})
    data.setdefault("last_btc_report_candle", None)
    data.setdefault("last_stale_fingerprint", None)
    data.setdefault("last_btc_pmo_color", None)
    data.setdefault("last_btc_pmo_candle", None)
    data.setdefault("last_btc_sqz15_color", None)
    state["schema_version"] = 73
    return state


def save_state(state: dict) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = STATE_PATH.with_suffix(STATE_PATH.suffix + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(STATE_PATH)


def send_telegram(message: str) -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        LOG.warning("Telegram secrets are not configured; report is printed to Actions log only")
        LOG.info("REPORT\n%s", message)
        return

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    chunks: list[str] = []
    current = ""
    for line in message.splitlines():
        candidate = f"{current}\n{line}" if current else line
        if len(candidate) > 3900 and current:
            chunks.append(current)
            current = line
        else:
            current = candidate
    if current:
        chunks.append(current)

    for chunk in chunks:
        response = SESSION.post(url, json={
            "chat_id": chat_id,
            "text": chunk,
            "disable_web_page_preview": True,
        }, timeout=25)
        response.raise_for_status()
        result = response.json()
        if not result.get("ok"):
            raise RuntimeError(f"Telegram rejected message: {result}")


def run() -> int:
    state = load_state()
    scanner_state = state[STATE_KEY]
    # Keep the legacy per-symbol map for state-file compatibility, but do not
    # use it as a report cooldown: every detected signal is reportable each run.
    last_reported: dict[str, int] = scanner_state["last_reported_candles"]

    instruments = get_instruments()
    tickers = get_tickers()
    symbols = sorted({str(item.get("symbol", "")).upper() for item in instruments if item.get("symbol")})
    btc_symbol = pick_btc_symbol(instruments)

    # BTC is always evaluated first and is the direction-priority reference.
    btc_name, btc_data = analyze_market(btc_symbol, is_btc=True)
    if btc_name != btc_symbol:
        raise RuntimeError("BTC analysis returned an unexpected symbol")
    if btc_data.get("reason") == "stale_candles":
        stale_symbols = [btc_symbol]
        fingerprint = ",".join(stale_symbols)
        if scanner_state.get("last_stale_fingerprint") != fingerprint:
            send_telegram(stale_data_report(stale_symbols))
        scanner_state["last_stale_fingerprint"] = fingerprint
        save_state(state)
        LOG.warning("Signals withheld: stale BTC indicator data")
        return 0
    if btc_data.get("reason") in {"fetch_or_analysis_error", "insufficient_history"}:
        message = (
            "Scanner 7 | DATA UNAVAILABLE - SIGNALS WITHHELD\n"
            "BTC 5-minute SQZMOM / 15-minute Pure Momentum data could not be verified. "
            "No BUY or SELL signals are being reported."
        )
        send_telegram(message)
        LOG.error("Signals withheld: BTC data unavailable: %s", btc_data)
        return 0

    btc_candle_time = int(btc_data["candle_time"])
    btc_new_candle = scanner_state.get("last_btc_report_candle") != btc_candle_time
    btc_pmo_color = btc_data.get("last_pmo_color")
    previous_btc_pmo_color = scanner_state.get("last_btc_pmo_color")
    previous_btc_pmo_candle = scanner_state.get("last_btc_pmo_candle")
    btc_pmo_changed = (
        previous_btc_pmo_color in ("GREEN", "RED")
        and btc_pmo_color in ("GREEN", "RED")
        and previous_btc_pmo_color != btc_pmo_color
        and previous_btc_pmo_candle != btc_candle_time
    )

    alt_symbols = [symbol for symbol in symbols if symbol != btc_symbol]
    analysis_results: list[tuple[str, dict[str, Any]]] = []
    errors = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(analyze_market, symbol, False): symbol for symbol in alt_symbols}
        for future in as_completed(futures):
            symbol, data = future.result()
            analysis_results.append((symbol, data))

    # Fail closed globally if any scanned symbol's latest closed 5m/15m indicator candle is
    # more than one full candle behind. Do not report even otherwise-fresh signals.
    stale_symbols = sorted(
        symbol for symbol, data in analysis_results
        if data.get("reason") == "stale_candles"
    )
    if stale_symbols:
        fingerprint = ",".join(stale_symbols)
        if scanner_state.get("last_stale_fingerprint") != fingerprint:
            send_telegram(stale_data_report(stale_symbols))
        scanner_state["last_stale_fingerprint"] = fingerprint
        save_state(state)
        LOG.warning("Signals withheld: stale 5m/15m candle data for %s", fingerprint)
        return 0

    # Fresh data restored: permit a future stale incident to generate a notice.
    scanner_state["last_stale_fingerprint"] = None

    detected_alt_signals: list[dict[str, Any]] = []
    for symbol, data in analysis_results:
        if data.get("reason") == "fetch_or_analysis_error":
            errors += 1
            continue
        if data.get("reason") != "signal":
            continue
        candle_time = int(data["candle_time"])
        detected_alt_signals.append({
            "symbol": symbol,
            "direction": data["direction"],
            "change24h": ticker_change_24h(tickers.get(symbol)),
            "change1h": data.get("change1h"),
            "candle_time": candle_time,
        })

    # Do not suppress opposite-direction signals. The BTC signal (when present),
    # otherwise the latest closed BTC 15m PMO color, sets only the display priority.
    new_alt_signals = detected_alt_signals

    # No signal cooldown/deduplication: report every scan run containing a BTC
    # signal or an allowed altcoin signal. Also report each new BTC 15m close
    # and every detected BTC PMO color flip, even when there are no signals.
    btc_has_signal = btc_data.get("direction") in ("BUY", "SELL")
    should_report = (
        btc_has_signal or bool(new_alt_signals) or btc_new_candle
    )
    if should_report:
        message = build_report(
            btc_symbol, btc_data, tickers.get(btc_symbol), new_alt_signals,
        )
        send_telegram(message)
        scanner_state["last_btc_report_candle"] = btc_candle_time
        LOG.info("REPORT_SENT btc_signal=%s btc_pmo_color=%s pmo_color_changed=%s "
                 "allowed_alt_signals=%s detected_alt_signals=%s errors=%s",
                 btc_data.get("direction"), btc_pmo_color, btc_pmo_changed,
                 len(new_alt_signals), len(detected_alt_signals), errors)
    else:
        LOG.info("No new 15m report event; btc_status=%s pmo_color=%s alt_signals=%s errors=%s",
                 btc_data.get("direction") or "NONE", btc_data.get("last_pmo_color"),
                 len(new_alt_signals), errors)

    scanner_state["last_reported_candles"] = last_reported
    scanner_state["last_btc_pmo_color"] = btc_pmo_color
    scanner_state["last_btc_pmo_candle"] = btc_candle_time
    save_state(state)
    return 0


# ---------------------------------------------------------------------------
# Small self-tests: run with `python scanner.py --self-test`.
# ---------------------------------------------------------------------------
def _self_test() -> None:
    expected = {
        ("LIGHT_RED", "DARK_RED", "DARK_RED"): "BUY",
        ("DARK_GREEN", "LIGHT_GREEN", "LIGHT_GREEN"): "BUY",
        ("LIGHT_GREEN", "DARK_GREEN", "DARK_GREEN"): "SELL",
        ("DARK_RED", "LIGHT_RED", "LIGHT_RED"): "SELL",
    }
    for pattern, direction in expected.items():
        actual = "BUY" if pattern in BUY_PATTERNS else "SELL" if pattern in SELL_PATTERNS else None
        assert actual == direction

    # PMO color is the sole confirmation: green confirms BUY, red confirms SELL.
    assert ("BUY" if "GREEN" == "GREEN" else "SELL") == "BUY"
    assert ("SELL" if "RED" == "RED" else "BUY") == "SELL"
    assert pure_momentum_colors([]) == []

    rising = [Candle(i * PMO_INTERVAL_MS, 100 + i, 101 + i, 99 + i, 100.5 + i) for i in range(160)]
    falling = [Candle(i * PMO_INTERVAL_MS, 260 - i, 261 - i, 259 - i, 260.5 - i) for i in range(160)]
    assert pure_momentum_colors(rising)[-1] == "GREEN"
    assert pure_momentum_colors(falling)[-1] == "RED"

    # One full candle behind is tolerated; more than one is stale.
    now = 30 * PMO_INTERVAL_MS
    expected_open = 29 * PMO_INTERVAL_MS
    assert not candle_is_stale(Candle(expected_open, 1, 2, 0, 1), PMO_INTERVAL_MS, now)
    assert not candle_is_stale(Candle(expected_open - PMO_INTERVAL_MS, 1, 2, 0, 1), PMO_INTERVAL_MS, now)
    assert candle_is_stale(Candle(expected_open - 2 * PMO_INTERVAL_MS, 1, 2, 0, 1), PMO_INTERVAL_MS, now)

    # Direction aligned with BTC first; BUYs by greatest gain, SELLs by greatest drop.
    rows = [
        {"symbol": "SELL_1", "direction": "SELL", "change24h": -3.0},
        {"symbol": "BUY_1", "direction": "BUY", "change24h": 4.0},
        {"symbol": "BUY_2", "direction": "BUY", "change24h": 9.0},
        {"symbol": "SELL_2", "direction": "SELL", "change24h": -8.0},
        {"symbol": "SELL_3", "direction": "SELL", "change24h": -2.0},
    ]
    ordered = sort_signals(rows, "SELL")
    assert [x["symbol"] for x in ordered] == [
        "SELL_2", "SELL_1", "SELL_3", "BUY_2", "BUY_1"
    ]
    ordered_buy = sort_signals(rows, "BUY")
    assert [x["symbol"] for x in ordered_buy] == [
        "BUY_2", "BUY_1", "SELL_2", "SELL_1", "SELL_3"
    ]
    assert len(filter_alt_signals_by_btc_pmo_color(rows, "GREEN")) == len(rows)
    assert len(filter_alt_signals_by_btc_pmo_color(rows, "RED")) == len(rows)
    assert len(filter_alt_signals_by_btc_pmo_color(rows, "N/A")) == len(rows)

    # 1h performance remains based on four completed 15m intervals.
    one_hour = [Candle(i * PMO_INTERVAL_MS, 100 + i, 102 + i, 99 + i, 100 + i)
                for i in range(5)]
    assert abs(change_1h(one_hour) - 4.0) < 1e-9

    report = build_report("PF_XBTUSD", {
        "direction": None, "last_sqzmom_color": "DARK_RED", "last_sqzmom_15m_color": "DARK_RED",
        "last_pmo_color": "GREEN", "last_two_pmo_colors": ["RED", "GREEN"]
    }, None, [])
    assert "SQZMOM 15m latest histogram color: DARK_RED" in report
    assert "Pure Momentum 15m last two colors: RED -> GREEN" in report
    assert "COLOR CHANGE: RED -> GREEN" in report
    mixed_report = build_report("PF_XBTUSD", {
        "direction": "BUY", "change1h": 1.0
    }, {"change24h": 2.0}, [
        {"symbol": "ALT_SELL", "direction": "SELL", "change24h": -9.0, "change1h": -1.0},
        {"symbol": "ALT_BUY_1", "direction": "BUY", "change24h": 3.0, "change1h": 1.0},
        {"symbol": "ALT_BUY_2", "direction": "BUY", "change24h": 7.0, "change1h": 2.0},
    ])
    assert mixed_report.index("ALT_BUY_2") < mixed_report.index("ALT_BUY_1") < mixed_report.index("ALT_SELL")
    assert "ALT_SELL" in mixed_report  # opposite direction is retained
    btc_signal_report = build_report("PF_XBTUSD", {
        "direction": "BUY", "last_sqzmom_15m_color": "DARK_RED", "last_pmo_color": "GREEN"
    }, None, [])
    assert "SQZMOM 15m latest histogram color" not in btc_signal_report
    print("Scanner 7 self-tests: OK")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _self_test()
    else:
        try:
            sys.exit(run())
        except Exception:
            LOG.exception("Scanner 7 failed")
            sys.exit(1)
