# -*- coding: utf-8 -*-
"""
全市場三率「季度批次」歷史補洞器 v2.0
========================================

解決 v1 問題：
- v1 對 1,600+ 個缺少季度逐檔呼叫 MOPS 單一公司頁，GitHub Actions 大量出現
  "No tables found" / 版型解析失敗。
- v2 改用 MOPS「綜合損益彙總表」ajax_t163sb04，
  一次查一個 市場 × 年度 × 季別，取得整批公司資料。

策略：
1. 先從 SQLite 找出「今天判斷三率真正缺少」的 (stock, year, quarter)。
2. 依季度分組。
3. 每一季依序抓：
   sii  上市
   otc  上櫃
   rotc 興櫃
   pub  公開發行
   找到股票後就不再對該股票繼續往後市場抓。
4. 只把目前普通股母體、且確實缺少的季度寫回 SQLite。
5. 不把解析失敗/查不到誤寫成 NODATA，保留下次重試。
6. 寫入後由既有 fundamental_full_market_fix.py 重新驗收並產生 duck_latest.xlsx。

三率定義仍固定：
毛利率、營業利益率、稅後淨利率；
最新單季 vs 前一單季，三者全部上升才是「三率三升」。
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import math
import re
import sqlite3
import sys
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))

import stock_duckbill as sd  # noqa: E402


MOPS_HOSTS = [
    "https://mops.twse.com.tw",
    "https://mopsov.twse.com.tw",
]
MARKETS = [
    ("sii", "TWSE/上市"),
    ("otc", "TPEx/上櫃"),
    ("rotc", "興櫃"),
    ("pub", "公開發行"),
]


def _target_date(s: Optional[str]) -> dt.date:
    if s:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    now = dt.datetime.now(dt.timezone(dt.timedelta(hours=8)))
    day = now.date() if now.hour >= 16 else now.date() - dt.timedelta(days=1)
    while day.weekday() >= 5:
        day -= dt.timedelta(days=1)
    return day


def _num(v):
    try:
        if v is None or pd.isna(v):
            return None
        s = str(v).strip().replace(",", "").replace("－", "-")
        s = re.sub(r"[\s*#=]", "", s)
        if s in {"", "-", "--", "---", "nan", "None", "N/A"}:
            return None
        x = float(s)
        return x if math.isfinite(x) else None
    except Exception:
        return sd.clean_num(v)


def _all_market_ids(conn: sqlite3.Connection, target: dt.date) -> Dict[str, str]:
    rows = conn.execute(
        "SELECT stock_id,name FROM prices WHERE date=? ORDER BY stock_id",
        (target.isoformat(),),
    ).fetchall()
    out: Dict[str, str] = {}
    for sid, name in rows:
        sid = str(sid)
        name = str(name or "")
        if sd.is_common_stock(sid, name):
            out[sid] = name
    return out


def _existing_periods(conn: sqlite3.Connection, sid: str, target: dt.date) -> Set[Tuple[int, int]]:
    rows = conn.execute(
        """SELECT year,quarter
           FROM fundamental_income_history
           WHERE stock_id=? AND availability_date IS NOT NULL AND availability_date<=?""",
        (sid, target.isoformat()),
    ).fetchall()
    return {(int(y), int(q)) for y, q in rows}


def collect_missing_core(
    conn: sqlite3.Connection,
    target: dt.date,
    ids: Dict[str, str],
) -> Tuple[Dict[Tuple[int, int], Set[str]], int]:
    by_period: Dict[Tuple[int, int], Set[str]] = {}
    skipped_na = 0

    for sid in ids:
        if sd._is_three_rate_na_stock(conn, sid, target):
            skipped_na += 1
            continue
        req = sd._three_rate_required_periods_for_stock(conn, sid, target)
        have = _existing_periods(conn, sid, target)
        for y, q in req:
            if (int(y), int(q)) in have:
                continue
            by_period.setdefault((int(y), int(q)), set()).add(sid)

    return by_period, skipped_na


def _unique_headers(values: List[str]) -> List[str]:
    out = []
    seen = {}
    for i, raw in enumerate(values):
        h = str(raw or "").strip()
        if not h or h.lower() == "nan":
            h = f"_col_{i}"
        n = seen.get(h, 0)
        seen[h] = n + 1
        out.append(h if n == 0 else f"{h}__{n+1}")
    return out


def _prepare_table(raw: pd.DataFrame) -> Optional[pd.DataFrame]:
    """把 MOPS 不同年代/產業表格整理成「第一列為欄名」的 DataFrame。"""
    if raw is None or raw.empty:
        return None

    # 先試 pandas 已推斷出的欄名。
    try:
        d = sd._flatten_columns(raw)
        if sd._df_col(d, ["公司代號", "公司代碼", "股票代號"]):
            return d
    except Exception:
        pass

    # 再從前 10 列找真正表頭。
    x = raw.copy()
    for idx in list(x.index)[:10]:
        vals = [
            "" if pd.isna(v) else str(v).replace("\n", "").strip()
            for v in x.loc[idx].tolist()
        ]
        if any(("公司代號" in v) or ("公司代碼" in v) or ("股票代號" in v) for v in vals):
            d = x.loc[x.index[x.index.get_loc(idx) + 1:]].copy()
            d.columns = _unique_headers(vals)
            return d.reset_index(drop=True)
    return None


def _find_col(df: pd.DataFrame, aliases: Iterable[str]) -> Optional[str]:
    try:
        c = sd._df_col(df, aliases)
        if c:
            return c
    except Exception:
        pass

    def nk(x):
        return re.sub(r"[\s（）()％%_\-:/]", "", str(x)).lower()

    norms = [(c, nk(c)) for c in df.columns]
    for a in aliases:
        na = nk(a)
        for c, nc in norms:
            if na and (na == nc or na in nc or nc in na):
                return c
    return None


def _read_mops_tables(html: str) -> List[pd.DataFrame]:
    """用兩種模式解析，避免不同年份 HTML 表頭結構不同。"""
    tables: List[pd.DataFrame] = []
    errors = []
    for kwargs in ({}, {"header": None}):
        try:
            got = pd.read_html(io.StringIO(html), **kwargs)
            tables.extend(got)
            if got:
                break
        except Exception as e:
            errors.append(f"{type(e).__name__}:{e}")
    if not tables:
        raise RuntimeError("MOPS季度彙總頁沒有可解析表格：" + " | ".join(errors[:2]))
    return tables


def parse_bulk_income_html(
    html: str,
    year: int,
    quarter: int,
    market_label: str,
    source: str,
    wanted: Set[str],
) -> Dict[str, dict]:
    out: Dict[str, dict] = {}
    tables = _read_mops_tables(html)

    for raw in tables:
        df = _prepare_table(raw)
        if df is None or df.empty:
            continue

        c_code = _find_col(df, ["公司代號", "公司代碼", "股票代號", "證券代號"])
        c_name = _find_col(df, ["公司名稱", "名稱"])
        if not c_code:
            continue

        c_rev = _find_col(df, [
            "營業收入", "營業收入合計", "營業收入淨額", "營業收益", "收益合計", "收入合計"
        ])
        c_gross = _find_col(df, [
            "營業毛利（毛損）淨額", "營業毛利(毛損)淨額",
            "營業毛利（毛損）", "營業毛利(毛損)", "營業毛利", "營業毛利淨額"
        ])
        c_op = _find_col(df, [
            "營業利益（損失）", "營業利益(損失)",
            "營業利益（損失）淨額", "營業利益(損失)淨額",
            "營業利益", "營業淨利", "營業利益淨額", "營業損益"
        ])
        c_net = _find_col(df, [
            "本期淨利（淨損）", "本期淨利(淨損)", "本期淨利",
            "本期稅後淨利（淨損）", "本期稅後淨利(淨損)", "本期稅後淨利",
            "歸屬於母公司業主之淨利（損）", "歸屬於母公司業主之淨利(損)",
            "歸屬母公司業主淨利（損）", "歸屬母公司業主淨利(損)"
        ])
        c_eps = _find_col(df, [
            "基本每股盈餘（元）", "基本每股盈餘(元)", "基本每股盈餘", "每股盈餘"
        ])

        # 一般業三率至少需要 revenue/gross/op/net 四個欄位。
        if not all([c_rev, c_gross, c_op, c_net]):
            continue

        for _, rr in df.iterrows():
            code = str(rr.get(c_code, "")).strip().replace(".0", "")
            code = re.sub(r"\s+", "", code)
            if not re.fullmatch(r"\d{4}", code):
                continue
            if wanted and code not in wanted:
                continue

            name = str(rr.get(c_name, "") if c_name else "").strip()
            rev = _num(rr.get(c_rev))
            gross = _num(rr.get(c_gross))
            op = _num(rr.get(c_op))
            net = _num(rr.get(c_net))
            eps = _num(rr.get(c_eps)) if c_eps else None

            if None in (rev, gross, op, net):
                continue
            if rev == 0:
                continue

            out[code] = {
                "stock_id": code,
                "name": name,
                "market": market_label,
                "year": int(year),
                "quarter": int(quarter),
                "revenue_ytd": rev,
                "gross_profit_ytd": gross,
                "operating_income_ytd": op,
                "net_income_ytd": net,
                "eps_ytd": eps,
                "source": source,
            }
    return out


def fetch_market_quarter(
    year: int,
    quarter: int,
    typek: str,
    market_label: str,
    wanted: Set[str],
) -> Tuple[Dict[str, dict], str]:
    roc = year - 1911
    payload = {
        "encodeURIComponent": "1",
        "step": "1",
        "firstin": "1",
        "off": "1",
        "isQuery": "Y",
        "TYPEK": typek,
        "year": str(roc),
        "season": f"{quarter:02d}",
    }
    errors = []

    for host in MOPS_HOSTS:
        url = f"{host}/mops/web/ajax_t163sb04"
        try:
            # 先暖 session/cookie；失敗不阻塞 POST。
            try:
                sd.SESSION.get(
                    f"{host}/mops/web/t163sb04",
                    timeout=12,
                    headers={"User-Agent": sd.UA, "Connection": "close"},
                )
            except Exception:
                pass

            r = sd.SESSION.post(
                url,
                data=payload,
                timeout=50,
                headers={
                    "User-Agent": sd.UA,
                    "Accept": "text/html,application/xhtml+xml,*/*",
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Referer": f"{host}/mops/web/t163sb04",
                    "Connection": "close",
                },
            )
            r.raise_for_status()

            raw = r.content
            html = None
            for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
                try:
                    html = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if html is None:
                html = raw.decode("utf-8", errors="replace")

            compact = re.sub(r"\s+", "", html)
            if any(x in compact for x in ["查無資料", "無符合條件", "沒有符合條件"]):
                return {}, f"{host}:{typek}:{year}Q{quarter}:官方查無"

            source = f"MOPS季度批次[t163sb04/{typek}/{year}Q{quarter}]"
            rows = parse_bulk_income_html(
                html, year, quarter, market_label, source, wanted
            )
            return rows, f"{host}:{typek}:{year}Q{quarter}:解析{len(rows)}檔"
        except Exception as e:
            errors.append(f"{host}:{typek}:{type(e).__name__}:{str(e)[:160]}")
            time.sleep(1.0)

    raise RuntimeError("；".join(errors))


def upsert_records(
    conn: sqlite3.Connection,
    target: dt.date,
    records: Dict[str, dict],
) -> int:
    n = 0
    for sid, rec in records.items():
        name = str(rec.get("name") or sd._stock_name_asof(conn, sid, target))
        source = str(rec.get("source") or "MOPS季度批次")
        avail = sd._income_available_for_identity(
            sid, name, int(rec["year"]), int(rec["quarter"]), source
        )
        if avail > target:
            # point-in-time 保守原則：當天依法還不能使用就不寫成可得資料。
            continue
        sd._upsert_income_record_with_basis(
            conn, rec, avail, "mops_quarterly_market_batch_conservative_deadline", source
        )
        n += 1
    if n:
        conn.commit()
    return n


def quanta_check(conn: sqlite3.Connection, target: dt.date, allow_equal: bool) -> dict:
    sid = "2382"
    hist = sd._latest_income_history_asof(conn, target, [sid])
    if hist is None or hist.empty:
        return {"stock_id": sid, "status": "無最新財報"}
    snap = hist.iloc[-1]
    result = sd._three_rates_for_stock(conn, snap, allow_equal)
    if not result:
        periods = [
            f"{int(y)}Q{int(q)}"
            for y, q in conn.execute(
                """SELECT year,quarter FROM fundamental_income_history
                   WHERE stock_id=? AND availability_date<=?
                   ORDER BY year,quarter""",
                (sid, target.isoformat()),
            ).fetchall()
        ]
        return {
            "stock_id": sid,
            "status": "仍無法計算",
            "history_periods": periods,
        }
    return {
        "stock_id": sid,
        "status": "三率三升" if bool(result.get("三率三升")) else "未三升",
        "period": result.get("三率資料期"),
        "gross_current_pct": result.get("毛利率本季%"),
        "gross_prev_pct": result.get("毛利率上季%"),
        "operating_current_pct": result.get("營業利益率本季%"),
        "operating_prev_pct": result.get("營業利益率上季%"),
        "net_current_pct": result.get("稅後純益率本季%"),
        "net_prev_pct": result.get("稅後純益率上季%"),
        "source": result.get("三率來源"),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date")
    args = ap.parse_args()

    target = _target_date(args.date)
    cfg = sd.load_config()
    allow_equal = cfg.getboolean("general", "three_rates_allow_equal", fallback=True)

    conn = sqlite3.connect(sd.DB_PATH)
    sd.init_db(conn)
    ids = _all_market_ids(conn, target)
    if not ids:
        raise SystemExit(f"{target} stock.db 找不到普通股母體")

    before, skipped_na = collect_missing_core(conn, target, ids)
    before_pairs = sum(len(v) for v in before.values())

    print("=" * 72)
    print(f"[季度批次三率補洞 v2] target={target} 普通股={len(ids)}")
    print(f"三率不適用={skipped_na}；缺少核心 stock-period={before_pairs}")
    print("缺少季度：", ", ".join(
        f"{y}Q{q}:{len(sids)}檔" for (y, q), sids in sorted(before.items())
    ) or "無")
    print("=" * 72, flush=True)

    fetch_log = []
    total_parsed = 0
    total_upserted = 0

    for (year, quarter), missing_sids in sorted(before.items()):
        remaining = set(missing_sids)
        print(f"\n[{year}Q{quarter}] 待補 {len(remaining)} 檔", flush=True)

        for typek, market_label in MARKETS:
            if not remaining:
                break
            try:
                found, msg = fetch_market_quarter(
                    year, quarter, typek, market_label, remaining
                )
                total_parsed += len(found)
                added = upsert_records(conn, target, found)
                total_upserted += added
                remaining.difference_update(found.keys())
                fetch_log.append({
                    "period": f"{year}Q{quarter}",
                    "TYPEK": typek,
                    "status": "OK",
                    "parsed": len(found),
                    "upserted": added,
                    "remaining": len(remaining),
                    "message": msg,
                })
                print(
                    f"  {typek}: 找到 {len(found)}，寫入 {added}，剩 {len(remaining)}",
                    flush=True,
                )
            except Exception as e:
                fetch_log.append({
                    "period": f"{year}Q{quarter}",
                    "TYPEK": typek,
                    "status": "ERROR",
                    "parsed": 0,
                    "upserted": 0,
                    "remaining": len(remaining),
                    "message": f"{type(e).__name__}: {e}",
                })
                print(f"  {typek}: 失敗 {type(e).__name__}: {e}", flush=True)
            time.sleep(0.7)

    # 寫入後重新做缺口統計。
    after, skipped_na_after = collect_missing_core(conn, target, ids)
    after_pairs = sum(len(v) for v in after.values())

    q2382 = quanta_check(conn, target, allow_equal)

    report = {
        "version": "quarterly-batch-v2",
        "target_date": target.isoformat(),
        "all_market_stock_count": len(ids),
        "three_rate_na": skipped_na_after,
        "missing_stock_period_before": before_pairs,
        "missing_stock_period_after": after_pairs,
        "filled_stock_period": max(0, before_pairs - after_pairs),
        "parsed_records": total_parsed,
        "upserted_records": total_upserted,
        "remaining_by_period": {
            f"{y}Q{q}": len(sids) for (y, q), sids in sorted(after.items())
        },
        "fetch_log": fetch_log,
        "quanta_2382": q2382,
        "method": {
            "endpoint": "MOPS ajax_t163sb04",
            "markets": ["sii", "otc", "rotc", "pub"],
            "requests_are_market_quarter_batch": True,
            "point_in_time": "conservative legal availability deadline",
            "unknown_is_false": False,
        },
    }

    report_path = ROOT / "fundamental_batch_income_backfill_status.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print("\n" + "=" * 72)
    print("[批次補洞完成]")
    print(f"缺口：{before_pairs} → {after_pairs}；補齊 {before_pairs-after_pairs}")
    print("2382 廣達：")
    print(json.dumps(q2382, ensure_ascii=False, indent=2))
    print(f"報告：{report_path}")
    print("=" * 72, flush=True)

    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
