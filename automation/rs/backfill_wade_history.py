# -*- coding: utf-8 -*-
"""一次性補齊 rs_latest.xlsx 的歷史 Wade / 盤面原始欄位。

資料來源：2018-01-02～2026-08-21 的 wade-history-results 歷史回測產物。
原則：
1. 既有非空值優先，不覆蓋目前正式資料。
2. 只用 seed 補既有歷史中的空值；若某日期整列不存在，才以 seed 建立原始列。
3. 補完原始欄位後，呼叫目前 repo 內 rs_breadth.py 的 _enrich_daily_counts / write_excel
   重新依「現行正式公式」計算 Wade 分數、狀態、警示與 Excel 顯示，不直接照抄舊版衍生分數。
4. 完成後覆蓋 repo 根目錄 rs_latest.xlsx；未來 daily-update 會從該檔讀回歷史 raw 欄位並延續保存。
"""
from __future__ import annotations

import datetime as dt
import importlib.util
import json
import shutil
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
RS_FILE = ROOT / "rs_latest.xlsx"
SEED_FILE = HERE / "wade_raw_overlay_seed_2018_20260821.csv.gz"
REPORT_FILE = ROOT / "wade_history_backfill_report.json"
RB_PATH = HERE / "rs_breadth.py"

RAW_COLS = [
    "eligible_count", "rs85_count", "strong_count",
    "advance_count", "decline_count", "flat_count",
    "limit_up_count", "limit_down_count", "new_high_count", "new_low_count",
    "advance_ratio", "twse_advance_ratio", "tpex_advance_ratio",
    "total_amount", "amount_ratio20", "mean_return", "twse_mean_return", "tpex_mean_return",
    "leader_retention_ratio", "new_leader_count",
    "twii_close", "twii_ret", "twoii_close", "twoii_ret",
]

# 真正需要 Wade 顯示的原始欄位；櫃買指數 twoii 在舊歷史 source 本來就不完整，因此不列入成功門檻。
REQUIRED_OVERLAY = [
    "advance_count", "decline_count", "flat_count",
    "limit_up_count", "limit_down_count", "new_high_count", "new_low_count",
    "advance_ratio", "twse_advance_ratio", "tpex_advance_ratio",
    "total_amount", "mean_return", "twse_mean_return", "tpex_mean_return",
    "leader_retention_ratio", "new_leader_count", "twii_close", "twii_ret",
]


def load_rb():
    spec = importlib.util.spec_from_file_location("rs_breadth_live", RB_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入 {RB_PATH}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def norm_date(s):
    return pd.to_datetime(s, errors="coerce").dt.normalize()


def missing_cells(df: pd.DataFrame, cols: list[str], cutoff: pd.Timestamp | None = None) -> int:
    z = df.copy()
    if cutoff is not None:
        z = z[z["date"] < cutoff]
    use = [c for c in cols if c in z.columns]
    if not use:
        return len(z) * len(cols)
    return int(z[use].isna().sum().sum())


def load_latest_list() -> pd.DataFrame:
    try:
        x = pd.read_excel(RS_FILE, sheet_name="最新強勢股名單")
    except Exception:
        return pd.DataFrame(columns=["code","name","market","close","rs","ma50","ma200","high250","low250","amount"])
    wanted = ["代號","名稱","市場","收盤","RS","MA50","MA200","52週高點","52週低點","成交金額"]
    for c in wanted:
        if c not in x.columns:
            x[c] = np.nan
    x = x[wanted].copy()
    x.columns = ["code","name","market","close","rs","ma50","ma200","high250","low250","amount"]
    return x


def load_recovery():
    try:
        return pd.read_excel(RS_FILE, sheet_name="被遺忘資金候選")
    except Exception:
        return None


def main():
    if not RS_FILE.exists():
        raise SystemExit(f"找不到 {RS_FILE}")
    if not SEED_FILE.exists():
        raise SystemExit(f"找不到 {SEED_FILE}")

    rb = load_rb()
    cfg = rb.cfg_get()
    current = rb._load_published_raw_history()
    if current.empty:
        raise SystemExit("目前 rs_latest.xlsx 無法讀取『每日強勢股數量』歷史。")

    current = current.copy()
    current["date"] = norm_date(current["date"])
    current = current.dropna(subset=["date"]).sort_values("date").drop_duplicates("date", keep="last")

    seed = pd.read_csv(SEED_FILE, compression="gzip")
    seed["date"] = norm_date(seed["date"])
    seed = seed.dropna(subset=["date"]).sort_values("date").drop_duplicates("date", keep="last")

    # 只補到 seed 結束日；seed 之後完全由目前正式 Excel / 每日更新負責。
    seed_end = seed["date"].max()
    cutoff = pd.Timestamp("2026-03-03")
    before_missing = missing_cells(current, REQUIRED_OVERLAY, cutoff=cutoff)

    cur = current.set_index("date")
    sd = seed.set_index("date")

    # 若 seed 有目前 Excel 完全沒有的日期，先加入；核心欄位只在目前缺值時才補。
    all_index = cur.index.union(sd.index).sort_values()
    cur = cur.reindex(all_index)
    for c in RAW_COLS:
        if c not in cur.columns:
            cur[c] = np.nan
        if c not in sd.columns:
            continue
        seed_vals = pd.to_numeric(sd[c], errors="coerce").reindex(all_index)
        old_vals = pd.to_numeric(cur[c], errors="coerce")
        # current 非空優先；seed 只補空值。
        cur[c] = old_vals.where(old_vals.notna(), seed_vals)

    raw = cur.reset_index().rename(columns={"index":"date"})
    raw = raw[(raw["date"] <= max(current["date"].max(), seed_end))].copy()
    raw = raw.sort_values("date").drop_duplicates("date", keep="last")

    after_missing_raw = missing_cells(raw, REQUIRED_OVERLAY, cutoff=cutoff)

    start = dt.datetime.strptime(cfg["general"].get("start_date", "2018-01-01"), "%Y-%m-%d").date()
    end = raw["date"].max().date()
    d, m, extreme_df = rb._enrich_daily_counts(raw, start, end, cfg)
    if d.empty:
        raise SystemExit("回填後無法重新計算每日統計。")

    # 只把 seed 當原始資料補洞；所有 Wade 衍生欄位使用「目前 repo 的正式公式」重算。
    latest = load_latest_list()
    recovery = load_recovery()
    out = rb.write_excel(d, m, latest, extreme_df, cfg, start, d["date"].max().date(), recovery)

    backup = RS_FILE.with_name("rs_latest_before_wade_backfill.xlsx")
    if not backup.exists():
        shutil.copy2(RS_FILE, backup)
    shutil.copy2(out, RS_FILE)

    pre = d[d["date"] < cutoff].copy()
    wade_missing = int(pd.to_numeric(pre.get("wade_score"), errors="coerce").isna().sum()) if "wade_score" in pre.columns else len(pre)
    display_missing = {}
    for c in ["advance_count","decline_count","new_high_count","new_low_count","wade_score"]:
        if c in pre.columns:
            display_missing[c] = int(pd.to_numeric(pre[c], errors="coerce").isna().sum())

    report = {
        "status": "success",
        "seed_period": [seed["date"].min().date().isoformat(), seed_end.date().isoformat()],
        "current_period_before": [current["date"].min().date().isoformat(), current["date"].max().date().isoformat()],
        "output_period": [d["date"].min().date().isoformat(), d["date"].max().date().isoformat()],
        "rows": int(len(d)),
        "cutoff_checked": cutoff.date().isoformat(),
        "required_overlay_missing_cells_before": before_missing,
        "required_overlay_missing_cells_after_raw_fill": after_missing_raw,
        "pre_cutoff_wade_score_missing_rows_after_recalc": wade_missing,
        "pre_cutoff_display_missing_rows_after_recalc": display_missing,
        "note": "既有非空值優先；seed只補空值。Wade衍生欄位使用目前rs_breadth.py重新計算。",
    }
    REPORT_FILE.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2), flush=True)
    print(f"已覆蓋：{RS_FILE}", flush=True)
    print(f"備份：{backup}", flush=True)


if __name__ == "__main__":
    main()
