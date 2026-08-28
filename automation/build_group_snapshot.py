# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RS_FILE = ROOT / "rs_latest.xlsx"
THEME_FILE = ROOT / "stock_theme_map.csv"
HISTORY_FILE = ROOT / "group_history.csv"

COLS = ["日期","類型","族群","族群強度","強勢股家數","占強勢股%","平均RS","RS90以上","成交金額占比%"]

def stock_code(v):
    s=str(v).replace(".0","").strip()
    return s

def split_values(v):
    if v is None or pd.isna(v): return []
    s=str(v).strip()
    if not s: return []
    return [x.strip() for x in re.split(r"[｜|/、;,，]+",s) if x.strip()]

def latest_date():
    daily=pd.read_excel(RS_FILE,sheet_name="每日強勢股數量")
    if "日期" not in daily.columns or daily.empty:
        return pd.Timestamp.today().normalize()
    d=pd.to_datetime(daily["日期"],errors="coerce").dropna()
    return d.max().normalize() if not d.empty else pd.Timestamp.today().normalize()

def summary(strong, group_col, group_type):
    x=strong.copy()
    x["_groups"]=x[group_col].map(split_values)
    x=x.explode("_groups")
    x=x[x["_groups"].fillna("").astype(str).str.strip()!=""]
    if x.empty: return pd.DataFrame(columns=COLS)
    x["RS_num"]=pd.to_numeric(x.get("RS"),errors="coerce")
    x["成交_num"]=pd.to_numeric(x.get("成交金額"),errors="coerce").fillna(0)
    total_codes=max(1,strong["代號"].nunique())
    total_amt=max(0.0,float(pd.to_numeric(strong.get("成交金額"),errors="coerce").fillna(0).sum()))
    date=latest_date().strftime("%Y-%m-%d")
    rows=[]
    for g,d in x.groupby("_groups"):
        n=int(d["代號"].nunique())
        share=n/total_codes*100
        avg=float(d["RS_num"].mean()) if d["RS_num"].notna().any() else 0.0
        rs90=int((d["RS_num"]>=90).sum())
        rs90_ratio=rs90/max(1,n)*100
        amt=float(d["成交_num"].sum())
        amt_share=amt/total_amt*100 if total_amt>0 else 0.0
        concentration=min(100.0,share*4.0)
        liquidity=min(100.0,amt_share*4.0)
        score=max(0.0,min(100.0,avg*0.35+concentration*0.30+rs90_ratio*0.20+liquidity*0.15))
        rows.append({
            "日期":date,"類型":group_type,"族群":str(g),
            "族群強度":round(score,1),"強勢股家數":n,"占強勢股%":round(share,1),
            "平均RS":round(avg,1),"RS90以上":rs90,"成交金額占比%":round(amt_share,1)
        })
    return pd.DataFrame(rows,columns=COLS)

def main():
    if not RS_FILE.exists() or not THEME_FILE.exists():
        print("group snapshot skipped: rs_latest.xlsx or stock_theme_map.csv missing")
        return
    strong=pd.read_excel(RS_FILE,sheet_name="最新強勢股名單")
    theme=pd.read_csv(THEME_FILE,dtype=str,encoding="utf-8-sig")
    if strong.empty or "代號" not in strong.columns or theme.empty or "代號" not in theme.columns:
        print("group snapshot skipped: required columns missing")
        return
    strong["代號"]=strong["代號"].map(stock_code)
    theme["代號"]=theme["代號"].map(stock_code)
    for c in ["題材族群","供應鏈"]:
        if c not in theme.columns: theme[c]=""
    x=strong.merge(theme[["代號","題材族群","供應鏈"]],on="代號",how="left")
    parts=[
        summary(x,"題材族群","題材族群"),
        summary(x,"供應鏈","供應鏈"),
    ]
    cur=pd.concat([p for p in parts if p is not None and not p.empty],ignore_index=True) if parts else pd.DataFrame(columns=COLS)
    if cur.empty:
        print("group snapshot skipped: no mapped themes in current strong list")
        return
    try:
        hist=pd.read_csv(HISTORY_FILE,encoding="utf-8-sig")
    except Exception:
        hist=pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in hist.columns: hist[c]=pd.NA
    # 同一日期、類型、族群採最新重算結果，避免手動重跑造成重複。
    key=set(zip(cur["日期"].astype(str),cur["類型"].astype(str),cur["族群"].astype(str)))
    if not hist.empty:
        mask=[(str(a),str(b),str(c)) not in key for a,b,c in zip(hist["日期"],hist["類型"],hist["族群"])]
        hist=hist.loc[mask, COLS]
    out=pd.concat([hist,cur],ignore_index=True)
    out["日期"]=pd.to_datetime(out["日期"],errors="coerce")
    out=out.dropna(subset=["日期"]).sort_values(["日期","類型","族群"])
    out["日期"]=out["日期"].dt.strftime("%Y-%m-%d")
    out.to_csv(HISTORY_FILE,index=False,encoding="utf-8-sig")
    print(f"group snapshot updated: {len(cur)} rows for {cur['日期'].iloc[0]} | history={len(out)}")

if __name__=="__main__":
    main()
