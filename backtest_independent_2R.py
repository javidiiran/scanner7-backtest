#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""2R audited runner for the independent BTC backtest.

Keep this file beside the original backtest_independent.py. The original file is
left untouched; this runner loads its engine, applies audited execution/reporting
corrections in memory, and executes it with a 2:1 reward/risk ratio.
"""
from __future__ import annotations
from pathlib import Path

BASE_FILE = Path(__file__).with_name("backtest_independent.py")


def replace_once(source: str, old: str, new: str, label: str) -> str:
    count = source.count(old)
    if count != 1:
        raise RuntimeError(
            f"Cannot safely apply patch '{label}': expected exactly one match, found {count}. "
            "The base engine may have changed; no backtest was run."
        )
    return source.replace(old, new, 1)


def build_source(source: str) -> str:
    patches = [
        ("reward/risk", "REWARD_RISK = 4.0", "REWARD_RISK = 2.0"),
        ("report filename", 'OUTPUT = Path("independent_backtest_report.xlsx")',
         'OUTPUT = Path("independent_backtest_report_2R.xlsx")'),
        ("funding timestamp alignment",
         '        if start_ms < ts <= end_ms:\n            out.append((ts, rate))',
         '        if ts % SQZ_MS != 0:\n'
         '            raise RuntimeError(f"Funding settlement {stamp} is not aligned to a 5m candle boundary; refusing ambiguous timing")\n'
         '        if start_ms < ts <= end_ms:\n            out.append((ts, rate))'),
        ("exit timestamp basis",
         'def close_position(pos, raw_exit, timestamp, reason, free_balance):',
         'def close_position(pos, raw_exit, timestamp, reason, free_balance, time_basis="INTRABAR_TIME_UNKNOWN_RECORDED_AT_CANDLE_CLOSE"):'),
        ("liquidation fee for labeled liquidation cases",
         '    liq_fee = abs(fill * pos.qty) * LIQUIDATION_FEE_RATE if reason == "LIQUIDATION" else 0.0',
         '    liq_fee = abs(fill * pos.qty) * LIQUIDATION_FEE_RATE if reason.startswith("LIQUIDATION") else 0.0'),
        ("exit time metadata",
         '        "exit_time_utc": datetime.fromtimestamp(timestamp / 1000, timezone.utc).isoformat(),',
         '        "exit_time_utc": datetime.fromtimestamp(timestamp / 1000, timezone.utc).isoformat(),\n'
         '        "exit_time_basis": time_basis,'),
        ("existing-position ambiguity classification",
         '            if gap_liq:\n'
         '                reason, raw = "LIQUIDATION_GAP", mark_candle.open\n'
         '            elif gap_stop:\n'
         '                reason, raw = "STOP_GAP", candle.open\n'
         '            elif gap_target:\n'
         '                reason, raw = "TAKE_PROFIT_GAP", candle.open\n'
         '            elif liq_hit:\n'
         '                reason, raw = "LIQUIDATION", pos.liq_price\n'
         '            elif stop_hit:\n'
         '                reason, raw = "STOP_LOSS", pos.stop\n'
         '            elif target_hit:\n'
         '                reason, raw = "TAKE_PROFIT", pos.target',
         '            if gap_liq:\n'
         '                reason = "LIQUIDATION_GAP_AMBIGUOUS_STOP" if gap_stop else "LIQUIDATION_GAP"\n'
         '                if gap_target:\n                    reason += "_AMBIGUOUS_TAKE_PROFIT"\n'
         '                raw = mark_candle.open\n'
         '            elif gap_stop:\n                reason, raw = "STOP_GAP", candle.open\n'
         '            elif gap_target:\n                reason, raw = "TAKE_PROFIT_GAP", candle.open\n'
         '            elif liq_hit and stop_hit:\n                reason, raw = "LIQUIDATION_AMBIGUOUS_STOP", pos.liq_price\n'
         '            elif liq_hit and target_hit:\n                reason, raw = "LIQUIDATION_AMBIGUOUS_TAKE_PROFIT", pos.liq_price\n'
         '            elif stop_hit and target_hit:\n                reason, raw = "STOP_LOSS_AMBIGUOUS_TAKE_PROFIT", pos.stop\n'
         '            elif liq_hit:\n                reason, raw = "LIQUIDATION", pos.liq_price\n'
         '            elif stop_hit:\n                reason, raw = "STOP_LOSS", pos.stop\n'
         '            elif target_hit:\n                reason, raw = "TAKE_PROFIT", pos.target'),
        ("existing-position exit timestamp semantics",
         '            if reason:\n'
         '                free_balance, trade = close_position(pos, raw, now, "LIQUIDATION" if reason.startswith("LIQUIDATION") else reason, free_balance)\n'
         '                ledger.append(trade)',
         '            if reason:\n'
         '                is_gap = "_GAP" in reason\n'
         '                exit_ts = now if is_gap else now + SQZ_MS\n'
         '                basis = "CANDLE_OPEN_GAP" if is_gap else "INTRABAR_TIME_UNKNOWN_RECORDED_AT_CANDLE_CLOSE"\n'
         '                free_balance, trade = close_position(pos, raw, exit_ts, reason, free_balance, basis)\n'
         '                ledger.append(trade)'),
        ("next-candle entry continuity guard",
         '        event = events.get(i)\n'
         '        if event and now >= test_start_ms:\n'
         '            if not event["confirmed"]:',
         '        event = events.get(i)\n'
         '        if event and now >= test_start_ms:\n'
         '            if now != event["signal_time"]:\n'
         '                signal_log.append({**event, "entry_bar_time": now, "status": "SKIPPED_NONCONTIGUOUS_ENTRY_BAR",\n'
         '                                   "reason": "NEXT_5M_OPEN_DOES_NOT_MATCH_SIGNAL_CLOSE"})\n'
         '            elif not event["confirmed"]:'),
        ("entry-candle ambiguity classification",
         '                    if liq_hit:\n'
         '                        reason, raw = "LIQUIDATION", pos.liq_price\n'
         '                    elif stop_hit:\n'
         '                        reason, raw = "STOP_LOSS_ENTRY_BAR", pos.stop\n'
         '                    elif target_hit:\n'
         '                        reason, raw = "TAKE_PROFIT_ENTRY_BAR", pos.target\n'
         '                    if reason:\n'
         '                        free_balance, trade = close_position(pos, raw, now, "LIQUIDATION" if reason == "LIQUIDATION" else reason, free_balance)\n'
         '                        ledger.append(trade)\n'
         '                        open_positions.remove(pos)',
         '                    if liq_hit and stop_hit:\n'
         '                        reason, raw = "LIQUIDATION_AMBIGUOUS_STOP_ENTRY_BAR", pos.liq_price\n'
         '                    elif liq_hit and target_hit:\n'
         '                        reason, raw = "LIQUIDATION_AMBIGUOUS_TAKE_PROFIT_ENTRY_BAR", pos.liq_price\n'
         '                    elif stop_hit and target_hit:\n'
         '                        reason, raw = "STOP_LOSS_AMBIGUOUS_TAKE_PROFIT_ENTRY_BAR", pos.stop\n'
         '                    elif liq_hit:\n                        reason, raw = "LIQUIDATION_ENTRY_BAR", pos.liq_price\n'
         '                    elif stop_hit:\n                        reason, raw = "STOP_LOSS_ENTRY_BAR", pos.stop\n'
         '                    elif target_hit:\n                        reason, raw = "TAKE_PROFIT_ENTRY_BAR", pos.target\n'
         '                    if reason:\n'
         '                        free_balance, trade = close_position(pos, raw, now + SQZ_MS, reason, free_balance,\n'
         '                                                             "INTRABAR_TIME_UNKNOWN_RECORDED_AT_CANDLE_CLOSE")\n'
         '                        ledger.append(trade)\n'
         '                        open_positions.remove(pos)'),
        ("summary ambiguity counters",
         '        "signals_skipped": sum(x["status"].startswith("SKIPPED") for x in signal_log),',
         '        "signals_skipped": sum(x["status"].startswith("SKIPPED") for x in signal_log),\n'
         '        "ambiguous_liquidation_stop_exits": sum("LIQUIDATION" in str(t["exit_reason"]) and "AMBIGUOUS_STOP" in str(t["exit_reason"]) for t in ledger),\n'
         '        "ambiguous_liquidation_take_profit_exits": sum("LIQUIDATION" in str(t["exit_reason"]) and "AMBIGUOUS_TAKE_PROFIT" in str(t["exit_reason"]) for t in ledger),\n'
         '        "ambiguous_stop_take_profit_exits": sum("STOP_LOSS_AMBIGUOUS_TAKE_PROFIT" in str(t["exit_reason"]) for t in ledger),'),
        ("assumption entry timing",
         '        ("Entry timing", "Next 5m candle open after the 5m signal candle has closed and confirmation is available"),',
         '        ("Entry timing", "Signal time is the 5m candle close boundary; entry is the immediately following 5m candle open at that same boundary. Non-contiguous next bars are skipped and logged."),'),
        ("assumption reward/risk",
         '        ("Take profit", "4:1 reward-to-risk based on entry-to-stop distance"),',
         '        ("Take profit", "2:1 reward-to-risk based on entry-to-stop distance"),'),
        ("assumption same-candle policy",
         '        ("Same candle ambiguity", "Mark-price liquidation has priority; otherwise stop-loss is prioritized over take-profit when both trade-price levels are touched in one 5m candle"),',
         '        ("Same candle ambiguity", "5m OHLC cannot determine intrabar event order. Conservative priority is applied but explicitly labeled: liquidation before SL/TP if mark-price liquidation and another exit are touched; SL before TP if both trade-price levels are touched. Summary counts ambiguous cases; this is a stress assumption, not observed exchange execution."),\n'
         '        ("Exit timestamps", "Gap exits use candle-open time. Intrabar exits are timestamped at candle close only as a reporting boundary; exact trigger time is unknown from 5m OHLC."),'),
    ]
    for label, old, new in patches:
        source = replace_once(source, old, new, label)
    return source


def main() -> None:
    if not BASE_FILE.is_file():
        raise FileNotFoundError(
            f"Missing {BASE_FILE.name}. Keep this 2R runner beside the original backtest_independent.py."
        )
    original = BASE_FILE.read_text(encoding="utf-8")
    revised = build_source(original)
    namespace = {
        "__name__": "__main__",
        "__file__": str(BASE_FILE),
        "__package__": None,
        "__cached__": None,
    }
    exec(compile(revised, str(BASE_FILE) + " [2R audited]", "exec"), namespace)


if __name__ == "__main__":
    main()
