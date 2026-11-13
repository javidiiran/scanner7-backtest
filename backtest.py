#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Scanner 7 historical backtest: Kraken Futures BTC perpetual (PF_XBTUSD),
30d test + 10d warmup, 5m SQZMOM + 15m Pure Momentum.
54 independent isolated-margin scenarios, each starting at $100. BTC BUY and
SELL signals are both included; there is no BTC-direction filter. Closed
candles only; execution costs and liquidation parameters are explicit estimates.
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
INITIAL_EQUITY, LEVERAGE = 100.0, 10.0
FEE_RATE, SLIPPAGE_RATE = 0.0005, 0.0002
MAINTENANCE_MARGIN_RATE, LIQUIDATION_FEE_RATE = 0.05, 0.005
FUNDING_RATE_PER_8H = 0.0
SQZ_MS, PMO_MS = 300_000, 900_000
KC_LENGTH, PMO_LENGTH = 20, 45
PMO_UP, PMO_DOWN = 0.05, -0.05
SL_BUFFERS, MARGIN_FRACTIONS = (0.0, 0.001, 0.002), (0.01, 0.10)
RR_VALUES = (1, 2, 3, 4, 8, 10)
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

def fetch_candles(symbol,resolution,interval_ms,start_ms,end_ms):
    url=f"{CHARTS}/trade/{symbol}/{resolution}"; got={}; errors=[]
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
        t=c5[i].time
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

def run_scenario(symbol,c5,c15,events,colors5,colors15,scenario,test_start):
    equity=INITIAL_EQUITY; peak=equity; maxdd=0.0; pos=None; trades=[]; skipped=0; piv=pivots(c15); pi=0
    for i,c in enumerate(c5):
        now=c.time
        while pi<len(piv) and piv[pi][0]<=now:
            _,pl,ph=piv[pi]; pi+=1
            if pos and scenario.exit_type=="STRUCTURE_TRAIL":
                if pos["direction"]=="BUY" and pl is not None and pl>pos["stop"]:pos["stop"]=pl
                if pos["direction"]=="SELL" and ph is not None and ph<pos["stop"]:pos["stop"]=ph
        if pos:
            d=pos["direction"]; why=None; raw=None; trigger_type="stop"
            if (d=="BUY" and c.open<=pos["stop"]) or (d=="SELL" and c.open>=pos["stop"]):why="STOP_GAP";raw=c.open
            else:
                if scenario.exit_type=="OPPOSITE_PMO":
                    known=[j for j,x in enumerate(c15) if x.time+PMO_MS<=now]
                    if known:
                        col=colors15[known[-1]]
                        if (d=="BUY" and col=="RED") or (d=="SELL" and col=="GREEN"):why="OPPOSITE_PMO";raw=c.open
                elif scenario.exit_type=="OPPOSITE_SQZ" and i>=3:
                    rev=sqz_dir(colors5[i-3:i])
                    if rev and rev!=d:why="OPPOSITE_SQZ";raw=c.open
                stop_hit=c.low<=pos["stop"] if d=="BUY" else c.high>=pos["stop"]
                target_hit=pos["target"] is not None and (c.high>=pos["target"] if d=="BUY" else c.low<=pos["target"])
                liq_hit=c.low<=pos["liq_price"] if d=="BUY" else c.high>=pos["liq_price"]
                if why is None:
                    if liq_hit and (not stop_hit or (pos["liq_price"]>pos["stop"] if d=="BUY" else pos["liq_price"]<pos["stop"])):why="LIQUIDATION";raw=pos["liq_price"];trigger_type="liq"
                    elif stop_hit:why="STOP";raw=pos["stop"]
                    elif target_hit:why="TAKE_PROFIT";raw=pos["target"];trigger_type="target"
            if why:
                if why in ("STOP","TAKE_PROFIT"):raw=gap_fill(raw,c.open,d,is_stop=(trigger_type=="stop"))
                fill=adverse_exit(raw,d); qty=pos["qty"]; gross=(fill-pos["entry"])*qty*(1 if d=="BUY" else -1)
                xf=abs(fill*qty)*FEE_RATE; lf=abs(fill*qty)*LIQUIDATION_FEE_RATE if why=="LIQUIDATION" else 0.0
                net=gross-xf-lf; equity+=net
                trades.append({"symbol":symbol,"scenario":scenario.name,"direction":d,"entry_time":pos["entry_time"],"exit_time":now,"entry_price":pos["entry"],"exit_price":fill,"initial_stop":pos["initial_stop"],"final_stop":pos["stop"],"liq_price":pos["liq_price"],"target":pos["target"],"margin":pos["margin"],"notional":pos["notional"],"qty_base_assumed":qty,"gross_pnl":gross,"entry_fee":pos["entry_fee"],"exit_fee":xf,"liquidation_fee":lf,"net_pnl":net-pos["entry_fee"],"equity_after":equity,"exit_reason":why})
                pos=None
        mtm=mark_equity(equity,pos,c.close); peak=max(peak,mtm)
        if peak>0:maxdd=max(maxdd,(peak-mtm)/peak)
        if pos is not None or now<test_start:continue
        ev=events.get(now)
        if not ev:continue
        d=ev["direction"]; stop=initial_stop(c15,now,d,scenario.buffer)
        if stop is None:skipped+=1;continue
        entry=adverse_entry(c.open,d)
        if (d=="BUY" and stop>=entry) or (d=="SELL" and stop<=entry):skipped+=1;continue
        margin=equity*scenario.margin_fraction; notional=margin*LEVERAGE
        if margin<=0 or notional<=0:skipped+=1;continue
        qty=notional/entry; fee=notional*FEE_RATE; equity-=fee; risk=abs(entry-stop)
        target=(entry+risk*scenario.rr if d=="BUY" else entry-risk*scenario.rr) if scenario.exit_type=="FIXED" else None
        liq=liquidation_price(entry,qty,margin,d)
        pos={"direction":d,"entry_time":now,"entry":entry,"initial_stop":stop,"stop":stop,"target":target,"margin":margin,"notional":notional,"qty":qty,"entry_fee":fee,"liq_price":liq}
        stop_hit=c.low<=stop if d=="BUY" else c.high>=stop
        liq_hit=c.low<=liq if d=="BUY" else c.high>=liq
        target_hit=target is not None and (c.high>=target if d=="BUY" else c.low<=target)
        if liq_hit and (not stop_hit or (liq>stop if d=="BUY" else liq<stop)):why="LIQUIDATION";raw=liq
        elif stop_hit:why="STOP_ENTRY_CANDLE";raw=stop
        elif target_hit:why="TAKE_PROFIT_ENTRY_CANDLE";raw=target
        else:why=None
        if why:
            fill=adverse_exit(raw,d); gross=(fill-entry)*qty*(1 if d=="BUY" else -1); xf=abs(fill*qty)*FEE_RATE
            lf=abs(fill*qty)*LIQUIDATION_FEE_RATE if why=="LIQUIDATION" else 0.0; net=gross-xf-lf; equity+=net
            trades.append({"symbol":symbol,"scenario":scenario.name,"direction":d,"entry_time":now,"exit_time":now,"entry_price":entry,"exit_price":fill,"initial_stop":stop,"final_stop":stop,"liq_price":liq,"target":target,"margin":margin,"notional":notional,"qty_base_assumed":qty,"gross_pnl":gross,"entry_fee":fee,"exit_fee":xf,"liquidation_fee":lf,"net_pnl":net-fee,"equity_after":equity,"exit_reason":why});pos=None
    if pos and c5:
        last=c5[-1];d=pos["direction"];fill=adverse_exit(last.close,d);gross=(fill-pos["entry"])*pos["qty"]*(1 if d=="BUY" else -1);xf=abs(fill*pos["qty"])*FEE_RATE;net=gross-xf;equity+=net
        trades.append({"symbol":symbol,"scenario":scenario.name,"direction":d,"entry_time":pos["entry_time"],"exit_time":last.time+SQZ_MS,"entry_price":pos["entry"],"exit_price":fill,"initial_stop":pos["initial_stop"],"final_stop":pos["stop"],"liq_price":pos["liq_price"],"target":pos["target"],"margin":pos["margin"],"notional":pos["notional"],"qty_base_assumed":pos["qty"],"gross_pnl":gross,"entry_fee":pos["entry_fee"],"exit_fee":xf,"liquidation_fee":0.0,"net_pnl":net-pos["entry_fee"],"equity_after":equity,"exit_reason":"END_OF_DATA_MARK_TO_MARKET"})
    wins=sum(t["net_pnl"]>0 for t in trades);losses=sum(t["net_pnl"]<0 for t in trades);n=len(trades)
    return {"symbol":symbol,"scenario":scenario.name,"margin_mode":"ISOLATED","exit_type":scenario.exit_type,"margin_fraction":scenario.margin_fraction,"sl_buffer":scenario.buffer,"rr":scenario.rr,"initial_equity":INITIAL_EQUITY,"final_equity":equity,"net_profit":equity-INITIAL_EQUITY,"return_pct":(equity/INITIAL_EQUITY-1)*100,"trades":n,"wins":wins,"losses":losses,"win_rate_pct":wins/n*100 if n else 0,"max_drawdown_pct":maxdd*100,"skipped_entries":skipped},trades

def quality(symbol,tf,cs,interval,errors,start_ms,end_ms):
    gaps=[]
    for a,b in zip(cs,cs[1:]):
        if b.time-a.time!=interval:gaps.append({"symbol":symbol,"timeframe":tf,"previous_time":a.time,"current_time":b.time,"gap_ms":b.time-a.time,"missing_intervals":max(0,(b.time-a.time)//interval-1)})
    test_gaps=[g for g in gaps if g["current_time"]>=start_ms and g["previous_time"]<end_ms]
    return {"symbol":symbol,"timeframe":tf,"candles":len(cs),"gaps":len(gaps),"test_window_gaps":len(test_gaps),"api_errors":len(errors),"error_details":" | ".join(errors[:15]),"valid_for_results":not errors and not test_gaps},gaps

def make_scenarios():
    out=[]
    for m in MARGIN_FRACTIONS:
        for b in SL_BUFFERS:
            for rr in RR_VALUES:out.append(Scenario(f"FIXED_M{m}_B{b}_RR{rr}",m,b,"FIXED",float(rr)))
    for m in MARGIN_FRACTIONS:
        for b in SL_BUFFERS:
            for ex in ("OPPOSITE_PMO","OPPOSITE_SQZ","STRUCTURE_TRAIL"):out.append(Scenario(f"{ex}_M{m}_B{b}",m,b,ex))
    assert len(out)==54
    return out

def main():
    logging.basicConfig(level=logging.INFO,format="%(asctime)s %(levelname)s %(message)s")
    now=datetime.now(timezone.utc);test_start=now-timedelta(days=DAYS);warm=test_start-timedelta(days=WARMUP_DAYS)
    start_ms=int(warm.timestamp()*1000);end_ms=int(now.timestamp()*1000);test_ms=int(test_start.timestamp()*1000)
    meta=get_market_metadata();scenarios=make_scenarios();assert len(scenarios)==54
    summaries=[];ledgers=[];qualities=[];gaps=[];selected_rows=[meta];excluded=[]
    s=meta["symbol"];LOG.info("BTC-only market: %s | scenarios=%d",s,len(scenarios))
    c5,e5=fetch_candles(s,"5m",SQZ_MS,start_ms,end_ms);c15,e15=fetch_candles(s,"15m",PMO_MS,start_ms,end_ms)
    q5,g5=quality(s,"5m",c5,SQZ_MS,e5,test_ms,end_ms);q15,g15=quality(s,"15m",c15,PMO_MS,e15,test_ms,end_ms)
    qualities += [q5,q15];gaps += g5+g15
    if len(c5)<200 or len(c15)<100 or not q5["valid_for_results"] or not q15["valid_for_results"]:
        excluded.append({"symbol":s,"reason":"insufficient history, API errors, or missing test-window candles; all 54 scenarios excluded"})
        LOG.error("%s excluded: data quality failed",s)
    else:
        events,colors5,colors15=build_events(c5,c15)
        for sc in scenarios:
            summary,trades=run_scenario(s,c5,c15,events,colors5,colors15,sc,test_ms);summaries.append(summary);ledgers.extend(trades)
    assumptions=[
        ("Market selection","BTC perpetual only: PF_XBTUSD (Kraken uses XBT for Bitcoin)"),
        ("Scenario count","54 total for BTC: 36 fixed TP + 18 floating exits"),
        ("Margin mode","ISOLATED only; no Cross mode"),
        ("Initial equity","$100 independently for each of the 54 scenarios"),
        ("Position sizing","Margin = 1% or 10% of current scenario equity; notional = margin × 10x"),
        ("Leverage","10x target; no per-contract leverage tier verification available"),
        ("Entry/exit fee",f"Taker estimate {FEE_RATE*100:.4f}% per side"),("Slippage",f"Adverse estimate {SLIPPAGE_RATE*100:.4f}% per side"),
        ("Maintenance margin",f"Generic estimate {MAINTENANCE_MARGIN_RATE*100:.2f}% of notional"),("Liquidation fee",f"Illustrative estimate {LIQUIDATION_FEE_RATE*100:.2f}% of notional"),
        ("Liquidation model","Approximate isolated collateral threshold; not Kraken's contract-specific liquidation engine"),
        ("Funding","0 assumed; historical funding not modeled"),
        ("Contract quantity","Base quantity = notional / price; assumes 1:1 linear multiplier; validate against Kraken instrument metadata"),
        ("Signal","SQZMOM LazyBear 5m patterns + latest closed Pure Momentum 15m confirmation"),
        ("BTC direction filter","None; all valid BTC BUY and SELL signals included"),
        ("Entry","First 5m open after both conditions are confirmed"),
        ("Initial stop","10 most recent closed 15m candles before entry, plus configured buffer"),
        ("Pivot trail","Confirmed 2-left / 2-right pivots; usable only after right candle closes"),
        ("Same-candle SL/TP","Stop-first; if liquidation is nearer to entry than stop, liquidation is modeled first"),
        ("Gap stop","If candle opens through stop, use open rather than stale stop price"),("Entry candle","Intrabar stop, target, and liquidation checks are applied"),
        ("Drawdown","Mark-to-market equity sampled at each 5m close"),("Open final position","Closed at final available 5m close"),
        ("Data integrity","BTC excluded from results if API errors or test-window candle gaps are detected"),
        ("Limitations","Fees, MMR, liquidation fee, funding, contract multiplier and leverage tiers are estimates; not an exact exchange simulator")]
    with pd.ExcelWriter(OUTPUT,engine="openpyxl") as w:
        sheets={"Summary":summaries,"Trade Ledger":ledgers,"Data Quality":qualities,"Data Gaps":gaps,"Selected Contracts":selected_rows,"Assumptions":[{"item":a,"value":b} for a,b in assumptions],"Excluded Symbols":excluded}
        for name,data in sheets.items():pd.DataFrame(data).to_excel(w,sheet_name=name,index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                letter=get_column_letter(col[0].column);width=max((len(str(c.value or "")) for c in col[:5000]),default=10)
                ws.column_dimensions[letter].width=min(width+2,48)
    LOG.info("Done: %s | market=%s | scenarios=%d | summaries=%d | trades=%d",OUTPUT.resolve(),BTC_SYMBOL,len(scenarios),len(summaries),len(ledgers))

if __name__=="__main__":main()
