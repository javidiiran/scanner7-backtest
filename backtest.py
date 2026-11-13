#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Scanner 7 historical backtest: Kraken Futures BTC perpetual (PF_XBTUSD),
30d test + 10d warmup, 5m SQZMOM + 15m Pure Momentum.
Single BTC-only isolated-margin strategy. Every confirmed signal opens a new
position using 10% of currently free balance at 50x leverage. Fixed 4:1 reward/risk.
"""
from __future__ import annotations
import math, time, logging
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
import requests
import pandas as pd
from openpyxl.utils import get_column_letter

BASE = "https://futures.kraken.com"
CHARTS = f"{BASE}/api/charts/v1"
INSTRUMENTS = f"{BASE}/derivatives/api/v3/instruments"
TICKERS = f"{BASE}/derivatives/api/v3/tickers"
DAYS, WARMUP_DAYS = 30, 10
BTC_SYMBOL = "PF_XBTUSD"
INITIAL_EQUITY, LEVERAGE = 100.0, 50.0
FEE_RATE, SLIPPAGE_RATE = 0.0005, 0.0002
MAINTENANCE_MARGIN_RATE, LIQUIDATION_FEE_RATE = 0.005, 0.005
SQZ_MS, PMO_MS = 300_000, 900_000
KC_LENGTH, PMO_LENGTH = 20, 45
PMO_UP, PMO_DOWN = 0.05, -0.05
SL_BUFFER, MARGIN_FRACTION, REWARD_RISK = 0.001, 0.10, 4.0
OUTPUT = Path("backtest_report.xlsx")
LOG = logging.getLogger("scanner7_backtest")
SESSION = requests.Session()
SESSION.headers.update({"User-Agent":"Scanner7-Backtest/3.0", "Accept":"application/json"})

@dataclass(frozen=True)
class Candle:
    time:int; open:float; high:float; low:float; close:float; volume:float=0.0
@dataclass(frozen=True)
class Scenario:
    name:str; margin_fraction:float; buffer:float; exit_type:str; rr:float|None=None

def num(v):
    try:
        x=float(v); return x if math.isfinite(x) else None
    except (TypeError,ValueError): return None

def api_get(url, params=None, attempts=5):
    err=None
    for k in range(attempts):
        try:
            r=SESSION.get(url,params=params,timeout=35)
            if r.status_code in (429,500,502,503,504): raise RuntimeError(f"HTTP {r.status_code}")
            r.raise_for_status(); d=r.json()
            if not isinstance(d,dict) or d.get("result")=="error": raise RuntimeError(f"Bad API payload: {d}")
            return d
        except Exception as e:
            err=e
            if k+1<attempts: time.sleep(1.5*(k+1))
    raise RuntimeError(f"Kraken API failed: {err}")

def get_market_metadata():
    """Return exactly Kraken's BTC perpetual; never fall back to another market."""
    ins=api_get(INSTRUMENTS).get("instruments",[])
    ticks=api_get(TICKERS).get("tickers",[])
    instrument=next((x for x in ins if str(x.get("symbol","" )).upper()==BTC_SYMBOL),None)
    if instrument is None:
        raise RuntimeError(f"Required BTC perpetual {BTC_SYMBOL} is absent from Kraken instrument metadata")
    typ=str(instrument.get("type","")).lower()
    if typ not in ("futures_vanilla","flexible_futures"):
        raise RuntimeError(f"{BTC_SYMBOL} is not a supported perpetual futures type (type={typ!r})")
    if instrument.get("tradeable") is False or instrument.get("suspended") is True:
        raise RuntimeError(f"{BTC_SYMBOL} is not currently tradeable")
    ticker=next((x for x in ticks if str(x.get("symbol","" )).upper()==BTC_SYMBOL),None)
    if ticker is None: raise RuntimeError(f"Ticker for {BTC_SYMBOL} is unavailable")
    last=num(ticker.get("last"))
    if last is None or last<=0: raise RuntimeError(f"Invalid last price for {BTC_SYMBOL}")
    return {"symbol":BTC_SYMBOL,"type":typ,"last_price":last,
            "contract_size":instrument.get("contractSize"),"tick_size":instrument.get("tickSize"),
            "quote_currency":instrument.get("quoteCurrency"),"underlying":instrument.get("underlying"),
            "market_selection":"BTC perpetual only (PF_XBTUSD)"}

def fetch_candles(symbol,resolution,interval_ms,start_ms,end_ms,tick_type="trade"):
    url=f"{CHARTS}/{tick_type}/{symbol}/{resolution}"; got={}; errors=[]
    cursor=start_ms//1000; end_s=end_ms//1000; window=3*86400
    while cursor<=end_s:
        stop=min(cursor+window,end_s)
        try:
            payload=api_get(url,{"from":cursor,"to":stop}); arr=payload.get("candles")
            if not isinstance(arr,list): raise RuntimeError("Missing candles array")
            for r in arr:
                try:
                    ts=int(float(r["time"])); ts=ts*1000 if ts<100_000_000_000 else ts
                    c=Candle(ts,float(r["open"]),float(r["high"]),float(r["low"]),float(r["close"]),float(r.get("volume",0) or 0))
                    if not all(math.isfinite(v) and v>0 for v in (c.open,c.high,c.low,c.close)): raise ValueError("invalid OHLC")
                    if c.high<max(c.open,c.close,c.low) or c.low>min(c.open,c.close,c.high): raise ValueError("invalid range")
                    if c.time+interval_ms<=end_ms: got[c.time]=c
                except (KeyError,TypeError,ValueError,OverflowError): errors.append(f"malformed {resolution} candle in {cursor}-{stop}")
        except Exception as e: errors.append(f"{resolution} {cursor}-{stop}: {e}")
        cursor=stop+1
    return [got[k] for k in sorted(got)],errors

def fetch_funding_rates(symbol,start_ms,end_ms):
    """Fetch Kraken Futures historical funding settlements for this contract."""
    url=f"{BASE}/derivatives/api/v4/historicalfundingrates"
    payload=api_get(url,{"symbol":symbol})
    rows=payload.get("rates")
    if not isinstance(rows,list):
        raise RuntimeError("Kraken historical funding response is missing its rates array")
    out=[]
    for row in rows:
        rate=num(row.get("relativeFundingRate")); stamp=row.get("timestamp")
        if rate is None or not stamp: continue
        try:
            ts=int(datetime.fromisoformat(str(stamp).replace("Z","+00:00")).timestamp()*1000)
        except Exception:
            continue
        if start_ms < ts <= end_ms: out.append((ts,rate))
    if not out:
        raise RuntimeError(f"No historical funding records returned for {symbol} in the requested period")
    return sorted(set(out))


def sma(a,n):
    out=[None]*len(a)
    for i in range(n-1,len(a)):
        w=a[i-n+1:i+1]
        if all(x is not None for x in w): out[i]=sum(w)/n
    return out

def extrema(a,n,high=True):
    out=[None]*len(a); f=max if high else min
    for i in range(n-1,len(a)): out[i]=f(a[i-n+1:i+1])
    return out

def linreg(a,n):
    out=[None]*len(a); xm=(n-1)/2; den=sum((j-xm)**2 for j in range(n))
    for i in range(n-1,len(a)):
        w=a[i-n+1:i+1]
        if any(x is None for x in w): continue
        ym=sum(w)/n; slope=sum((j-xm)*(w[j]-ym) for j in range(n))/den
        out[i]=ym+slope*(n-1-xm)
    return out

def sqz_colors(cs):
    h=extrema([c.high for c in cs],KC_LENGTH,True); l=extrema([c.low for c in cs],KC_LENGTH,False)
    av=sma([c.close for c in cs],KC_LENGTH); src=[None]*len(cs)
    for i,c in enumerate(cs):
        if h[i] is not None and l[i] is not None and av[i] is not None: src[i]=c.close-((h[i]+l[i])/2+av[i])/2
    vals=linreg(src,KC_LENGTH); colors=[]
    for i,v in enumerate(vals):
        if v is None: colors.append(None); continue
        p=vals[i-1] if i and vals[i-1] is not None else 0.0
        colors.append("LIGHT_GREEN" if v>0 and v>p else "DARK_GREEN" if v>0 else "LIGHT_RED" if v<p else "DARK_RED")
    return colors
BUY_PATTERNS={("LIGHT_RED","DARK_RED","DARK_RED"),("DARK_GREEN","LIGHT_GREEN","LIGHT_GREEN")}
SELL_PATTERNS={("LIGHT_GREEN","DARK_GREEN","DARK_GREEN"),("DARK_RED","LIGHT_RED","LIGHT_RED")}
def sqz_dir(p): return "BUY" if tuple(p) in BUY_PATTERNS else "SELL" if tuple(p) in SELL_PATTERNS else None

def ema(a,n):
    if not a:return []
    out=[a[0]]; alpha=2/(n+1)
    for x in a[1:]: out.append(alpha*x+(1-alpha)*out[-1])
    return out
def wma(a,n):
    out=[None]*len(a); den=n*(n+1)/2
    for i in range(n-1,len(a)):
        w=a[i-n+1:i+1]
        if all(x is not None for x in w):out[i]=sum((j+1)*w[j] for j in range(n))/den
    return out
def rma(a,n):
    out=[None]*len(a)
    if len(a)<n:return out
    out[n-1]=sum(a[:n])/n
    for i in range(n,len(a)):out[i]=(out[i-1]*(n-1)+a[i])/n
    return out
def pmo_colors(cs):
    n=len(cs); raw={k:[getattr(c,k) for c in cs] for k in ("open","high","low","close")}
    dbl={k:[v[i]*2-e[i] for i in range(n)] for k,v in raw.items() for e in [ema(v,PMO_LENGTH)]}
    cda=wma(dbl["close"],PMO_LENGTH); tr=[]
    for i,c in enumerate(cs):tr.append(c.high-c.low if i==0 else max(c.high-c.low,abs(c.high-cs[i-1].close),abs(c.low-cs[i-1].close)))
    atr=rma(tr,PMO_LENGTH); trend=0; out=[]
    for i in range(n):
        change=(cda[i]-cda[i-1])/atr[i] if i and cda[i] is not None and cda[i-1] is not None and atr[i] else None
        if change is not None:
            if change>PMO_UP:trend=1
            elif change<PMO_DOWN:trend=-1
        out.append("GREEN" if trend==1 else "RED")
    return out

def build_events(c5,c15):
    a=sqz_colors(c5); b=pmo_colors(c15); events={}; j=-1
    for i in range(3,len(c5)):
        t=c5[i].time+SQZ_MS  # signal is known only after this 5m candle closes
        while j+1<len(c15) and c15[j+1].time+PMO_MS<=t:j+=1
        if j<0:continue
        d=sqz_dir(a[i-3:i]); confirm="BUY" if b[j]=="GREEN" else "SELL"
        if d==confirm:events[t]={"direction":d,"sqz_pattern":tuple(a[i-3:i]),"pmo_index":j}
    return events,a,b

def initial_stop(c15,t,d,buff):
    prior=[c for c in c15 if c.time+PMO_MS<=t]
    if len(prior)<10:return None
    w=prior[-10:]
    return min(c.low for c in w)*(1-buff) if d=="BUY" else max(c.high for c in w)*(1+buff)

def pivots(cs,left=2,right=2):
    out=[]
    for i in range(left,len(cs)-right):
        w=cs[i-left:i+right+1]; p=cs[i]
        lo=all(p.low<x.low for j,x in enumerate(w) if j!=left); hi=all(p.high>x.high for j,x in enumerate(w) if j!=left)
        if lo or hi:out.append((cs[i+right].time+PMO_MS,p.low if lo else None,p.high if hi else None))
    return out

def adverse_entry(p,d):return p*(1+SLIPPAGE_RATE if d=="BUY" else 1-SLIPPAGE_RATE)
def adverse_exit(p,d):return p*(1-SLIPPAGE_RATE if d=="BUY" else 1+SLIPPAGE_RATE)
def gap_fill(trigger,op,d,is_stop=True):
    if d=="BUY":return min(trigger,op) if is_stop else max(trigger,op)
    return max(trigger,op) if is_stop else min(trigger,op)
def liquidation_price(entry,qty,margin,d):
    notional=entry*qty; reserve=notional*(MAINTENANCE_MARGIN_RATE+LIQUIDATION_FEE_RATE)
    move=max(0.0,margin-reserve)/qty
    return max(1e-12,entry-move) if d=="BUY" else entry+move
def mark_equity(equity,pos,price):
    if not pos:return equity
    return equity+(price-pos["entry"])*pos["qty"]*(1 if pos["direction"]=="BUY" else -1)

def run_scenario(symbol,c5,c15,events,colors5,colors15,test_start,mark5,funding_rates):
    """One account, isolated positions; each signal allocates 10% of free cash."""
    balance=INITIAL_EQUITY; peak=INITIAL_EQUITY; maxdd=0.0
    positions=[]; trades=[]; signal_audit=[]; skipped=0; signals=0; entries=0
    mark_by_time={c.time:c for c in mark5}; funding_i=0
    skipped_no_stop=0; skipped_bad_stop=0; skipped_no_funds=0
    for i,c in enumerate(c5):
        now=c.time
        # Apply settlements to positions that were open immediately before settlement.
        while funding_i < len(funding_rates) and funding_rates[funding_i][0] <= now:
            fts,frate=funding_rates[funding_i]
            mark_ts=((fts-1)//SQZ_MS)*SQZ_MS
            mc=mark_by_time.get(mark_ts)
            if mc is None:
                raise RuntimeError(f"Missing 5m mark candle for funding settlement at {fts}")
            for fp in positions:
                if fp["entry_time"] < fts:
                    # Positive relative funding: longs pay, shorts receive.
                    payment=-(1 if fp["direction"]=="BUY" else -1)*fp["qty"]*mc.close*frate
                    balance += payment
                    fp["funding_pnl"] += payment
            funding_i += 1
        # Process every existing isolated position independently.
        still_open=[]
        for pos in positions:
            d=pos['direction']; why=None; raw=None; trigger='stop'
            mc=mark_by_time.get(now)
            if mc is None: raise RuntimeError(f'Missing 5m mark candle at {now}')
            liq_hit=mc.low<=pos['liq_price'] if d=='BUY' else mc.high>=pos['liq_price']
            stop_gap=(d=='BUY' and c.open<=pos['stop']) or (d=='SELL' and c.open>=pos['stop'])
            if liq_hit:
                # Conservative tie-break: mark-price liquidation takes priority
                # whenever liquidation and another exit are touched in one candle.
                why='LIQUIDATION'; raw=pos['liq_price']; trigger='liq'
            elif stop_gap:
                why='STOP_GAP'; raw=c.open
            else:
                stop_hit=c.low<=pos['stop'] if d=='BUY' else c.high>=pos['stop']
                target_hit=c.high>=pos['target'] if d=='BUY' else c.low<=pos['target']
                if stop_hit:
                    why='STOP_LOSS'; raw=pos['stop']
                elif target_hit:
                    why='TAKE_PROFIT'; raw=pos['target']; trigger='target'
            if why is None:
                still_open.append(pos); continue
            if why in ('STOP_LOSS','TAKE_PROFIT'):
                raw=gap_fill(raw,c.open,d,is_stop=(trigger=='stop'))
            fill=adverse_exit(raw,d); qty=pos['qty']
            gross=(fill-pos['entry'])*qty*(1 if d=='BUY' else -1)
            exit_fee=abs(fill*qty)*FEE_RATE
            liq_fee=abs(fill*qty)*LIQUIDATION_FEE_RATE if why=='LIQUIDATION' else 0.0
            balance += gross-exit_fee-liq_fee
            net=gross-pos['entry_fee']-exit_fee-liq_fee+pos.get('funding_pnl',0.0)
            trades.append({
                'symbol':symbol,'scenario':'SINGLE_10PCT_50X_RR4_SL10x15m',
                'direction':d,'signal_time':pos['signal_time'],'entry_time':pos['entry_time'],
                'exit_time':now,'entry_price':pos['entry'],'exit_price':fill,
                'initial_stop':pos['stop'],'target':pos['target'],'margin':pos['margin'],
                'leverage':LEVERAGE,'notional':pos['notional'],'qty_base_assumed':qty,
                'gross_pnl':gross,'entry_fee':pos['entry_fee'],'exit_fee':exit_fee,
                'liquidation_fee':liq_fee,'funding_pnl':pos.get('funding_pnl',0.0),'net_pnl':net,'equity_after':balance,
                'exit_reason':why})
        positions=still_open

        # Account mark-to-market and drawdown include all concurrently open positions.
        unrealized=sum((c.close-p['entry'])*p['qty']*(1 if p['direction']=='BUY' else -1) for p in positions)
        mtm=balance+unrealized
        peak=max(peak,mtm)
        if peak>0: maxdd=max(maxdd,(peak-mtm)/peak)

        if now<test_start: continue
        ev=events.get(now)
        if not ev: continue
        signals += 1
        d=ev['direction']; stop=initial_stop(c15,now,d,SL_BUFFER)
        audit={'signal_time':now-SQZ_MS,'execution_time':now,'direction':d,'status':'SKIPPED',
               'entry_price':adverse_entry(c.open,d),'stop_price':stop,'target_price':None,
               'margin_reserved':0.0,'reason':''}
        if stop is None:
            skipped+=1; skipped_no_stop+=1; audit['reason']='NOT_ENOUGH_10_CLOSED_15M_CANDLES'; signal_audit.append(audit); continue
        entry=adverse_entry(c.open,d)
        if (d=='BUY' and stop>=entry) or (d=='SELL' and stop<=entry):
            skipped+=1; skipped_bad_stop+=1; audit['reason']='STOP_ON_WRONG_SIDE_OF_ENTRY'; signal_audit.append(audit); continue
        locked=sum(p['margin'] for p in positions)
        free=max(0.0,balance-locked)
        margin=free*MARGIN_FRACTION
        if margin<=0 or balance<=0:
            skipped+=1; skipped_no_funds+=1; audit['reason']='NO_FREE_BALANCE'; signal_audit.append(audit); continue
        notional=margin*LEVERAGE; qty=notional/entry
        entry_fee=notional*FEE_RATE
        # Charge fee immediately; collateral itself remains in wallet balance
        # but is reserved and excluded from the next position's free balance.
        balance-=entry_fee
        risk=abs(entry-stop)
        target=entry+risk*REWARD_RISK if d=='BUY' else entry-risk*REWARD_RISK
        audit.update({'status':'ENTERED','entry_price':entry,'stop_price':stop,'target_price':target,
                      'margin_reserved':margin,'reason':'OPENED_ISOLATED_50X'}); signal_audit.append(audit)
        liq=liquidation_price(entry,qty,margin,d)
        pos={'direction':d,'signal_time':now-SQZ_MS,'entry_time':now,'entry':entry,
             'stop':stop,'target':target,'margin':margin,'notional':notional,
             'qty':qty,'entry_fee':entry_fee,'liq_price':liq,'funding_pnl':0.0}
        entries+=1
        # Entry candle is evaluated too; signal was confirmed on the prior close.
        stop_hit=c.low<=stop if d=='BUY' else c.high>=stop
        target_hit=c.high>=target if d=='BUY' else c.low<=target
        mc=mark_by_time.get(now)
        if mc is None: raise RuntimeError(f'Missing 5m mark candle at {now}')
        liq_hit=mc.low<=liq if d=='BUY' else mc.high>=liq
        why=None; raw=None
        if liq_hit:
            why='LIQUIDATION'; raw=liq
        elif stop_hit: why='STOP_LOSS_ENTRY_CANDLE'; raw=stop
        elif target_hit: why='TAKE_PROFIT_ENTRY_CANDLE'; raw=target
        if why:
            fill=adverse_exit(raw,d); gross=(fill-entry)*qty*(1 if d=='BUY' else -1)
            exit_fee=abs(fill*qty)*FEE_RATE
            liq_fee=abs(fill*qty)*LIQUIDATION_FEE_RATE if why=='LIQUIDATION' else 0.0
            balance+=gross-exit_fee-liq_fee
            trades.append({'symbol':symbol,'scenario':'SINGLE_10PCT_50X_RR4_SL10x15m',
                'direction':d,'signal_time':now-SQZ_MS,'entry_time':now,'exit_time':now,
                'entry_price':entry,'exit_price':fill,'initial_stop':stop,'target':target,
                'margin':margin,'leverage':LEVERAGE,'notional':notional,'qty_base_assumed':qty,
                'gross_pnl':gross,'entry_fee':entry_fee,'exit_fee':exit_fee,
                'liquidation_fee':liq_fee,'funding_pnl':0.0,'net_pnl':gross-entry_fee-exit_fee-liq_fee,
                'equity_after':balance,'exit_reason':why})
        else:
            positions.append(pos)

    # Close remaining positions at final available close for a transparent MTM report.
    if positions and c5:
        last=c5[-1]
        for pos in positions:
            d=pos['direction']; fill=adverse_exit(last.close,d)
            gross=(fill-pos['entry'])*pos['qty']*(1 if d=='BUY' else -1)
            exit_fee=abs(fill*pos['qty'])*FEE_RATE
            balance+=gross-exit_fee
            trades.append({'symbol':symbol,'scenario':'SINGLE_10PCT_50X_RR4_SL10x15m',
                'direction':d,'signal_time':pos['signal_time'],'entry_time':pos['entry_time'],
                'exit_time':last.time+SQZ_MS,'entry_price':pos['entry'],'exit_price':fill,
                'initial_stop':pos['stop'],'target':pos['target'],'margin':pos['margin'],
                'leverage':LEVERAGE,'notional':pos['notional'],'qty_base_assumed':pos['qty'],
                'gross_pnl':gross,'entry_fee':pos['entry_fee'],'exit_fee':exit_fee,
                'liquidation_fee':0.0,'funding_pnl':pos.get('funding_pnl',0.0),'net_pnl':gross-pos['entry_fee']-exit_fee+pos.get('funding_pnl',0.0),
                'equity_after':balance,'exit_reason':'END_OF_DATA_MARK_TO_MARKET'})
    wins=sum(t['net_pnl']>0 for t in trades); losses=sum(t['net_pnl']<0 for t in trades); n=len(trades)
    return ({'symbol':symbol,'scenario':'SINGLE_10PCT_50X_RR4_SL10x15m',
        'margin_mode':'ISOLATED','initial_equity':INITIAL_EQUITY,'final_equity':balance,
        'net_profit':balance-INITIAL_EQUITY,'return_pct':(balance/INITIAL_EQUITY-1)*100,
        'leverage':LEVERAGE,'margin_fraction_of_free_balance':MARGIN_FRACTION,
        'stop_buffer_pct':SL_BUFFER*100,'stop_lookback_15m_candles':10,
        'reward_risk':REWARD_RISK,'signals_detected':signals,'entries_opened':entries,
        'trades_closed_or_marked':n,'wins':wins,'losses':losses,
        'win_rate_pct':wins/n*100 if n else 0,'max_drawdown_pct':maxdd*100,
        'skipped_signals':skipped,'skipped_no_10_candle_stop':skipped_no_stop,
        'skipped_invalid_stop':skipped_bad_stop,'skipped_insufficient_free_balance':skipped_no_funds},trades,signal_audit)

def quality(symbol,tf,cs,interval,errors,start_ms,end_ms):
    gaps=[]
    for a,b in zip(cs,cs[1:]):
        if b.time-a.time!=interval:gaps.append({"symbol":symbol,"timeframe":tf,"previous_time":a.time,"current_time":b.time,"gap_ms":b.time-a.time,"missing_intervals":max(0,(b.time-a.time)//interval-1)})
    test_gaps=[g for g in gaps if g["current_time"]>=start_ms and g["previous_time"]<end_ms]
    return {"symbol":symbol,"timeframe":tf,"candles":len(cs),"gaps":len(gaps),"test_window_gaps":len(test_gaps),"api_errors":len(errors),"error_details":" | ".join(errors[:15]),"valid_for_results":not errors and not test_gaps},gaps


def main():
    logging.basicConfig(level=logging.INFO,format="%(asctime)s %(levelname)s %(message)s")
    now=datetime.now(timezone.utc);test_start=now-timedelta(days=DAYS);warm=test_start-timedelta(days=WARMUP_DAYS)
    start_ms=int(warm.timestamp()*1000);end_ms=int(now.timestamp()*1000);test_ms=int(test_start.timestamp()*1000)
    meta=get_market_metadata()
    summaries=[];ledgers=[];signal_rows=[];qualities=[];gaps=[];selected_rows=[meta];excluded=[]
    s=meta["symbol"];LOG.info("BTC-only market: %s | single strategy",s)
    c5,e5=fetch_candles(s,"5m",SQZ_MS,start_ms,end_ms);c15,e15=fetch_candles(s,"15m",PMO_MS,start_ms,end_ms)
    mark5,em=fetch_candles(s,"5m",SQZ_MS,start_ms,end_ms,tick_type="mark")
    funding_rates=fetch_funding_rates(s,start_ms,end_ms)
    q5,g5=quality(s,"5m",c5,SQZ_MS,e5,test_ms,end_ms);q15,g15=quality(s,"15m",c15,PMO_MS,e15,test_ms,end_ms)
    qm,gm=quality(s,"5m MARK",mark5,SQZ_MS,em,test_ms,end_ms)
    qualities += [q5,q15,qm];gaps += g5+g15+gm
    if len(c5)<200 or len(c15)<100 or len(mark5)<200 or not q5["valid_for_results"] or not q15["valid_for_results"] or not qm["valid_for_results"]:
        excluded.append({"symbol":s,"reason":"insufficient history, API errors, or missing test-window candles; no results produced"})
        LOG.error("%s excluded: data quality failed",s)
    else:
        events,colors5,colors15=build_events(c5,c15)
        summary,trades,signals=run_scenario(s,c5,c15,events,colors5,colors15,test_ms,mark5,funding_rates);summaries.append(summary);ledgers.extend(trades);signal_rows.extend(signals)
    assumptions=[
        ("Market selection","BTC perpetual only: PF_XBTUSD"),
        ("Strategy count","One strategy only; no 54-scenario batch"),
        ("Margin mode","ISOLATED; each position has its own allocated collateral"),
        ("Initial equity",f"${INITIAL_EQUITY:.2f}"),
        ("Position sizing","Each new signal reserves 10% of currently free balance; multiple positions may remain open"),
        ("Leverage",f"{LEVERAGE:.0f}x"),
        ("Reward/risk",f"{REWARD_RISK:.0f}:1 fixed take-profit relative to initial stop distance"),
        ("Stop loss","BUY: 0.1% below the lowest low of the 10 latest CLOSED 15m candles; SELL: 0.1% above their highest high"),
        ("Signal","Existing Scanner 7 logic unchanged: SQZMOM LazyBear 5m + latest closed Pure Momentum 15m confirmation"),
        ("Signal execution","Entry at the first 5m open after the signal candle closes; no lookahead"),
        ("Multiple entries","Every valid new signal may open a new isolated position; no one-position-only restriction"),
        ("Exit","Close at stop loss or 4R take-profit; isolated liquidation is also modeled if it occurs first"),
        ("Entry/exit fee",f"Estimated taker fee {FEE_RATE*100:.4f}% per side"),
        ("Slippage",f"Adverse estimate {SLIPPAGE_RATE*100:.4f}% per side"),
        ("Maintenance margin",f"Illustrative estimate {MAINTENANCE_MARGIN_RATE*100:.2f}% of notional; actual Kraken tier/region may differ"),
        ("Liquidation fee",f"Illustrative estimate {LIQUIDATION_FEE_RATE*100:.2f}% of notional"),
        ("Liquidation trigger","5m historical MARK-price OHLC; trade-price candles are used for stop/target execution"),
        ("Funding","Historical Kraken relativeFundingRate applied at settlement timestamps; positive rates debit longs and credit shorts"),
        ("Data integrity","Results withheld if API errors or test-window candle gaps are detected"),
        ("Limitations","Funding uses the preceding 5m mark close at settlement; maintenance tiers, liquidation fees and matching engine remain estimates, not an exact Kraken replay")]
    with pd.ExcelWriter(OUTPUT,engine="openpyxl") as w:
        sheets={"Summary":summaries,"Signal Log":signal_rows,"Trade Ledger":ledgers,"Data Quality":qualities,"Data Gaps":gaps,"Selected Contracts":selected_rows,"Assumptions":[{"item":a,"value":b} for a,b in assumptions],"Excluded Symbols":excluded}
        for name,data in sheets.items():pd.DataFrame(data).to_excel(w,sheet_name=name,index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                letter=get_column_letter(col[0].column);width=max((len(str(c.value or "")) for c in col[:5000]),default=10)
                ws.column_dimensions[letter].width=min(width+2,48)
    LOG.info("Done: %s | market=%s | single strategy | summaries=%d | trades=%d",OUTPUT.resolve(),BTC_SYMBOL,len(summaries),len(ledgers))

if __name__=="__main__":main()
