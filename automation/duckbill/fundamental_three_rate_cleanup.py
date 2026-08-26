# -*- coding: utf-8 -*-
"""
全市場三率剩餘例外清理 v2.1
===========================

目的：
- v2 批次補洞後，全市場已由 1,478 個核心 stock-period 缺口降到 38。
- 本程式只處理這種「最後少量例外」，不再對全市場暴力重抓。

流程：
1. 先做 3 輪季度批次重試（MOPS市場彙總表），處理暫時 timeout。
2. 對仍缺少的少量 stock-period，才使用 MOPS 官方單一公司歷史頁逐筆補。
3. 最後不把所有非 yes/no 股票混叫「未知」，而是依既有驗收函式分成：
   - 可計算三率
   - 三率不適用
   - 上市時間不足/不適用
   - 比較期歷史不足
   - 缺資料
4. 產出 JSON + CSV 例外清單，讓剩餘公司都有明確原因。

原則：
- 只使用官方 MOPS/TWSE/TPEx 財報資料作正式三率。
- 解析失敗不寫成 NODATA。
- 不把資料不足當成「未三升」。
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import sqlite3
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))

import stock_duckbill as sd  # noqa: E402
import fundamental_batch_income_backfill as batch  # noqa: E402


def _target_date(s: Optional[str]) -> dt.date:
    if s:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    now = dt.datetime.now(dt.timezone(dt.timedelta(hours=8)))
    day = now.date() if now.hour >= 16 else now.date() - dt.timedelta(days=1)
    while day.weekday() >= 5:
        day -= dt.timedelta(days=1)
    return day


def _all_market_ids(conn: sqlite3.Connection, target: dt.date) -> Dict[str, str]:
    rows = conn.execute(
        "SELECT stock_id,name FROM prices WHERE date=? ORDER BY stock_id",
        (target.isoformat(),),
    ).fetchall()
    out = {}
    for sid, name in rows:
        sid = str(sid)
        name = str(name or "")
        if sd.is_common_stock(sid, name):
            out[sid] = name
    return out


def _validation(conn, target: dt.date, ids: Dict[str, str]) -> pd.DataFrame:
    df = sd._income_validation_df(conn, target, list(ids))
    if df is None or df.empty:
        return pd.DataFrame(columns=["stock_id", "status", "reason", "required"])
    df = df.copy()
    df["stock_id"] = df["stock_id"].astype(str)
    df["name"] = df["stock_id"].map(ids).fillna("")
    return df[["stock_id", "name", "status", "reason", "required"]]


def _status_counts(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {}
    vc = df["status"].fillna("未分類").value_counts()
    return {str(k): int(v) for k, v in vc.items()}


def _batch_retry_round(
    conn: sqlite3.Connection,
    target: dt.date,
    ids: Dict[str, str],
    round_no: int,
) -> dict:
    missing, _ = batch.collect_missing_core(conn, target, ids)
    before = sum(len(v) for v in missing.values())
    logs = []
    if before == 0:
        return {"round": round_no, "before": 0, "after": 0, "filled": 0, "logs": []}

    print(f"[批次重試 round {round_no}] 核心缺口 {before}", flush=True)

    for (year, quarter), sids in sorted(missing.items()):
        remaining = set(sids)
        for typek, market_label in batch.MARKETS:
            if not remaining:
                break
            try:
                found, msg = batch.fetch_market_quarter(
                    year, quarter, typek, market_label, remaining
                )
                added = batch.upsert_records(conn, target, found)
                remaining.difference_update(found.keys())
                logs.append({
                    "period": f"{year}Q{quarter}",
                    "TYPEK": typek,
                    "found": len(found),
                    "added": added,
                    "remaining": len(remaining),
                    "status": "OK",
                    "message": msg,
                })
                print(
                    f"  {year}Q{quarter} {typek}: 找到{len(found)} 寫入{added} 剩{len(remaining)}",
                    flush=True,
                )
            except Exception as e:
                logs.append({
                    "period": f"{year}Q{quarter}",
                    "TYPEK": typek,
                    "found": 0,
                    "added": 0,
                    "remaining": len(remaining),
                    "status": "ERROR",
                    "message": f"{type(e).__name__}: {str(e)[:240]}",
                })
                print(
                    f"  {year}Q{quarter} {typek}: {type(e).__name__}: {str(e)[:140]}",
                    flush=True,
                )
            time.sleep(1.0)

    after_map, _ = batch.collect_missing_core(conn, target, ids)
    after = sum(len(v) for v in after_map.values())
    return {
        "round": round_no,
        "before": before,
        "after": after,
        "filled": max(0, before - after),
        "logs": logs,
    }


def _single_company_official_fallback(
    conn: sqlite3.Connection,
    target: dt.date,
    ids: Dict[str, str],
    tries_per_item: int = 2,
) -> dict:
    """只對批次重試後仍缺的少量 stock-period 做官方逐檔補抓。"""
    missing, _ = batch.collect_missing_core(conn, target, ids)
    tasks = [
        (sid, y, q)
        for (y, q), sids in sorted(missing.items())
        for sid in sorted(sids)
    ]
    added = official_nodata = failed = 0
    details = []

    print(f"[官方單檔最後補洞] 待補 {len(tasks)} 個 stock-period", flush=True)

    for idx, (sid, y, q) in enumerate(tasks, 1):
        rec = None
        last_err = None
        for attempt in range(1, max(1, tries_per_item) + 1):
            try:
                rec = sd._fetch_mops_income_history_one(sid, y, q)
                last_err = None
                break
            except Exception as e:
                last_err = f"{type(e).__name__}: {str(e)[:220]}"
                time.sleep(1.5 * attempt)

        if rec:
            try:
                sd._upsert_income_history_record(conn, rec)
                added += 1
                status = "ADDED"
                msg = "官方單檔頁補入"
            except Exception as e:
                failed += 1
                status = "DB_ERROR"
                msg = f"{type(e).__name__}: {str(e)[:220]}"
        elif last_err is None:
            # 官方明確查無；只記錄到 cleanup report，不塞 NODATA 快取，
            # 避免未來市場身分/資料修正後被永久跳過。
            official_nodata += 1
            status = "OFFICIAL_NO_DATA"
            msg = "官方明確查無"
        else:
            failed += 1
            status = "ERROR"
            msg = last_err

        details.append({
            "stock_id": sid,
            "name": ids.get(sid, ""),
            "period": f"{y}Q{q}",
            "status": status,
            "message": msg,
        })
        if idx == 1 or idx % 10 == 0 or idx == len(tasks):
            print(
                f"  [{idx}/{len(tasks)}] 新增={added} 官方查無={official_nodata} 失敗={failed}",
                flush=True,
            )
        time.sleep(0.6)

    after_map, _ = batch.collect_missing_core(conn, target, ids)
    return {
        "tasks": len(tasks),
        "added": added,
        "official_no_data": official_nodata,
        "failed": failed,
        "remaining_stock_period": sum(len(v) for v in after_map.values()),
        "details": details,
    }


def _write_exception_csv(df: pd.DataFrame, path: Path) -> None:
    if df is None:
        return
    keep = df.loc[df["status"] != "可計算三率"].copy()
    keep.to_csv(path, index=False, encoding="utf-8-sig")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date")
    ap.add_argument("--batch-retries", type=int, default=3)
    ap.add_argument("--single-tries", type=int, default=2)
    args = ap.parse_args()

    target = _target_date(args.date)
    conn = sqlite3.connect(sd.DB_PATH)
    sd.init_db(conn)

    ids = _all_market_ids(conn, target)
    if not ids:
        raise SystemExit(f"{target} stock.db 找不到全市場普通股")

    before_df = _validation(conn, target, ids)
    missing_before, _ = batch.collect_missing_core(conn, target, ids)
    core_before = sum(len(v) for v in missing_before.values())

    print("=" * 78)
    print(f"[三率例外清理 v2.1] target={target} 普通股={len(ids)}")
    print("清理前驗收：", json.dumps(_status_counts(before_df), ensure_ascii=False))
    print(f"清理前核心 stock-period 缺口={core_before}")
    print("=" * 78, flush=True)

    rounds = []
    last_after = core_before
    for r in range(1, max(1, args.batch_retries) + 1):
        rr = _batch_retry_round(conn, target, ids, r)
        rounds.append(rr)
        if rr["after"] == 0:
            break
        # 若本輪沒有進展，仍允許下一輪重試一次不同網路狀態。
        last_after = rr["after"]

    single = _single_company_official_fallback(
        conn, target, ids, tries_per_item=max(1, args.single_tries)
    )

    after_df = _validation(conn, target, ids)
    missing_after, _ = batch.collect_missing_core(conn, target, ids)
    core_after = sum(len(v) for v in missing_after.values())

    status_after = _status_counts(after_df)
    true_unknown_statuses = {"缺資料", "比較期歷史不足"}
    true_unresolved = int(
        after_df["status"].isin(true_unknown_statuses).sum()
    ) if not after_df.empty else 0

    exception_csv = ROOT / "fundamental_three_rate_exceptions.csv"
    _write_exception_csv(after_df, exception_csv)

    unresolved_rows = (
        after_df.loc[after_df["status"].isin(true_unknown_statuses)]
        .to_dict("records")
        if not after_df.empty else []
    )

    report = {
        "version": "three-rate-cleanup-v2.1",
        "target_date": target.isoformat(),
        "all_market_stock_count": len(ids),
        "core_missing_before": core_before,
        "core_missing_after": core_after,
        "core_filled": max(0, core_before - core_after),
        "validation_before": _status_counts(before_df),
        "validation_after": status_after,
        "batch_retry_rounds": rounds,
        "single_company_official_fallback": single,
        "true_unresolved_count": true_unresolved,
        "true_unresolved": unresolved_rows,
        "classification_note": {
            "可計算三率": "正式可判定三率三升/未三升",
            "三率不適用": "金融/保險/金控/異業等報表結構不同",
            "上市時間不足/不適用": "沒有足夠可比季度，不應列成資料抓取失敗",
            "比較期歷史不足": "最新季有，但較舊比較期官方資料仍缺",
            "缺資料": "依法應有最新季或核心欄位仍缺，需後續追蹤",
        },
        "official_only_for_three_rate": True,
    }

    report_path = ROOT / "fundamental_three_rate_cleanup_status.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print("\n" + "=" * 78)
    print("[清理完成]")
    print(f"核心缺口：{core_before} → {core_after}")
    print("清理後驗收：", json.dumps(status_after, ensure_ascii=False))
    print(f"真正仍需追蹤：{true_unresolved} 檔")
    print(f"JSON：{report_path}")
    print(f"CSV：{exception_csv}")
    print("=" * 78, flush=True)

    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
