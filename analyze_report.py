# -*- coding: utf-8 -*-
"""
analyze_report.py
- trade_logs/YYYYMMDD.jsonl ¡æ report_out/YYYYMMDD/{tables,charts,summary.xml}
"""


import os
import json
import argparse
import datetime as dt
from typing import Any, Dict, List, Optional


import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import xml.etree.ElementTree as ET


DEFAULT_LOG_DIR = "trade_logs"
DEFAULT_OUT_DIR = "report_out"


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--date", default=None, help="YYYYMMDD (default=today)")
    p.add_argument("--log-dir", default=DEFAULT_LOG_DIR)
    p.add_argument("--out-dir", default=DEFAULT_OUT_DIR)
    p.add_argument("--bucket-min", type=int, default=30)
    return p.parse_args()


def yyyymmdd_today() -> str:
    return dt.datetime.now().strftime("%Y%m%d")


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def read_jsonl(path: str) -> List[Dict[str, Any]]:
    rows = []
    if not os.path.exists(path):
        return rows
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except:
                continue
    return rows


def to_dt(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def floor_bucket(ts: pd.Series, minutes: int) -> pd.Series:
    return ts.dt.floor(f"{minutes}min")


def save_csv(df: pd.DataFrame, path: str):
    df.to_csv(path, index=False, encoding="utf-8-sig")


def plot_time_series(df: pd.DataFrame, x: str, y: str, hue: Optional[str], title: str, out_path: str):
    plt.figure()
    if df.empty:
        plt.title(title + " (no data)")
        plt.savefig(out_path, dpi=150, bbox_inches="tight")
        plt.close()
        return


    if hue and hue in df.columns:
        for k, g in df.groupby(hue):
            plt.plot(g[x], g[y], marker="o", label=str(k))
        plt.legend()
    else:
        plt.plot(df[x], df[y], marker="o")


    plt.title(title)
    plt.xlabel(x)
    plt.ylabel(y)


    if pd.api.types.is_datetime64_any_dtype(df[x]):
        ax = plt.gca()
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
        plt.xticks(rotation=45)


    plt.grid(True, alpha=0.3)
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()


def df_to_xml_table(parent: ET.Element, name: str, df: pd.DataFrame, max_rows: int = 300):
    t = ET.SubElement(parent, "table", {"name": name})
    cols = list(df.columns)
    head = ET.SubElement(t, "header")
    for c in cols:
        ET.SubElement(head, "col").text = str(c)


    body = ET.SubElement(t, "rows")
    sliced = df.head(max_rows)
    for _, row in sliced.iterrows():
        r = ET.SubElement(body, "row")
        for c in cols:
            v = row.get(c)
            if pd.isna(v):
                v = ""
            ET.SubElement(r, "cell").text = str(v)


def main():
    args = parse_args()
    day = args.date or yyyymmdd_today()


    log_path = os.path.join(args.log_dir, f"{day}.jsonl")
    out_base = os.path.join(args.out_dir, day)
    charts_dir = os.path.join(out_base, "charts")
    tables_dir = os.path.join(out_base, "tables")
    ensure_dir(charts_dir)
    ensure_dir(tables_dir)


    events = read_jsonl(log_path)
    if not events:
        print(f"[WARN] no events: {log_path}")
        root = ET.Element("report", {"date": day})
        ET.SubElement(root, "note").text = "no events"
        ET.ElementTree(root).write(os.path.join(out_base, "summary.xml"), encoding="utf-8", xml_declaration=True)
        return


    df = pd.DataFrame(events)
    if "ts" not in df.columns:
        df["ts"] = None
    df["ts_dt"] = to_dt(df["ts"])
    df["bucket"] = floor_bucket(df["ts_dt"], args.bucket_min)


    sell = df[df["event"] == "sell_fill"].copy()
    buy = df[df["event"] == "buy_fill"].copy()


    for c in ["cond", "sell_reason", "code", "name"]:
        if c not in sell.columns:
            sell[c] = ""
    for c in ["realized_pnl", "realized_roi", "account_remain_sec", "account_remain_min", "sell_qty", "sell_price", "avg_price"]:
        if c not in sell.columns:
            sell[c] = 0


    sell_reason_bucket = (
        sell.groupby(["bucket", "sell_reason"], dropna=False)
        .agg(
            sells=("code", "count"),
            pnl_sum=("realized_pnl", "sum"),
            roi_mean=("realized_roi", "mean"),
            remain_min_mean=("account_remain_min", "mean"),
        )
        .reset_index()
        .sort_values(["bucket", "sell_reason"])
    )
    save_csv(sell_reason_bucket, os.path.join(tables_dir, "sell_reason_by_bucket.csv"))


    if not sell_reason_bucket.empty:
        plot_time_series(
            sell_reason_bucket.sort_values("bucket"),
            x="bucket", y="pnl_sum", hue="sell_reason",
            title="Sell PnL Sum by 30m Bucket (by sell_reason)",
            out_path=os.path.join(charts_dir, "sell_pnl_sum_by_reason.png"),
        )
        plot_time_series(
            sell_reason_bucket.sort_values("bucket"),
            x="bucket", y="roi_mean", hue="sell_reason",
            title="Sell ROI Mean by 30m Bucket (by sell_reason)",
            out_path=os.path.join(charts_dir, "sell_roi_mean_by_reason.png"),
        )


    sell_cond_bucket = (
        sell.groupby(["bucket", "cond"], dropna=False)
        .agg(
            sells=("code", "count"),
            pnl_sum=("realized_pnl", "sum"),
            roi_mean=("realized_roi", "mean"),
        )
        .reset_index()
        .sort_values(["bucket", "cond"])
    )
    save_csv(sell_cond_bucket, os.path.join(tables_dir, "sell_by_bucket_cond.csv"))


    if not sell_cond_bucket.empty:
        plot_time_series(
            sell_cond_bucket.sort_values("bucket"),
            x="bucket", y="pnl_sum", hue="cond",
            title="Realized PnL Sum by 30m Bucket (by cond)",
            out_path=os.path.join(charts_dir, "sell_pnl_sum_by_cond.png"),
        )
        plot_time_series(
            sell_cond_bucket.sort_values("bucket"),
            x="bucket", y="roi_mean", hue="cond",
            title="Realized ROI Mean by 30m Bucket (by cond)",
            out_path=os.path.join(charts_dir, "sell_roi_mean_by_cond.png"),
        )


    mark = df[df["event"] == "mark_30m"].copy()
    mark_items = []
    if not mark.empty:
        for _, r in mark.iterrows():
            b = r.get("bucket")
            items = r.get("items")
            if isinstance(items, list):
                for it in items:
                    it2 = dict(it)
                    it2["bucket"] = b
                    mark_items.append(it2)
    mark_df = pd.DataFrame(mark_items)
    if not mark_df.empty:
        mark_bucket_cond = (
            mark_df.groupby(["bucket", "cond"], dropna=False)
            .agg(symbols=("code", "nunique"), mark_roi_mean=("mark_roi", "mean"), mark_pnl_sum=("mark_pnl", "sum"))
            .reset_index()
            .sort_values(["bucket", "cond"])
        )
        save_csv(mark_bucket_cond, os.path.join(tables_dir, "mark_30m_by_bucket_cond.csv"))
        plot_time_series(
            mark_bucket_cond.sort_values("bucket"),
            x="bucket", y="mark_roi_mean", hue="cond",
            title="Mark ROI Mean by 30m Bucket (by cond)",
            out_path=os.path.join(charts_dir, "mark_roi_mean_by_cond.png"),
        )
    else:
        mark_bucket_cond = pd.DataFrame(columns=["bucket", "cond", "symbols", "mark_roi_mean", "mark_pnl_sum"])


    snap = df[df["event"] == "condition_snapshot_30m"].copy()
    snap_rows = []
    if not snap.empty:
        for _, r in snap.iterrows():
            b = r.get("bucket")
            conds = r.get("conditions")
            if isinstance(conds, dict):
                for cond_name, v in conds.items():
                    row = {"bucket": b, "cond": cond_name}
                    if isinstance(v, dict):
                        row.update(v)
                    snap_rows.append(row)
    snap_df = pd.DataFrame(snap_rows)
    if not snap_df.empty:
        snap_bucket_cond = (
            snap_df.groupby(["bucket", "cond"], dropna=False)
            .agg(
                current_count=("current_count", "max"),
                exited_today_count=("exited_today_count", "max"),
                dup_exited_today_count=("dup_exited_today_count", "max"),
                exit_repeat_total=("exit_repeat_total", "max"),
            )
            .reset_index()
            .sort_values(["bucket", "cond"])
        )
        save_csv(snap_bucket_cond, os.path.join(tables_dir, "condition_snapshot_by_bucket_cond.csv"))
        plot_time_series(
            snap_bucket_cond.sort_values("bucket"),
            x="bucket", y="current_count", hue="cond",
            title="Condition Current Count by 30m Bucket",
            out_path=os.path.join(charts_dir, "cond_current_count.png"),
        )
    else:
        snap_bucket_cond = pd.DataFrame(columns=["bucket", "cond", "current_count", "exited_today_count", "dup_exited_today_count", "exit_repeat_total"])


    root = ET.Element("report", {"date": day, "bucket_min": str(args.bucket_min)})
    meta = ET.SubElement(root, "meta")
    ET.SubElement(meta, "log_path").text = log_path
    ET.SubElement(meta, "generated_at").text = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ET.SubElement(meta, "event_count").text = str(len(df))


    summ = ET.SubElement(root, "summary")
    ET.SubElement(summ, "sell_count").text = str(len(sell))
    ET.SubElement(summ, "buy_count").text = str(len(buy))
    ET.SubElement(summ, "sell_pnl_sum").text = str(int(sell["realized_pnl"].sum())) if not sell.empty else "0"
    ET.SubElement(summ, "sell_roi_mean").text = str(float(sell["realized_roi"].mean())) if not sell.empty else "0"


    tables = ET.SubElement(root, "tables")
    df_to_xml_table(tables, "sell_reason_by_bucket", sell_reason_bucket)
    df_to_xml_table(tables, "sell_by_bucket_cond", sell_cond_bucket)
    df_to_xml_table(tables, "mark_30m_by_bucket_cond", mark_bucket_cond)
    df_to_xml_table(tables, "condition_snapshot_by_bucket_cond", snap_bucket_cond)


    xml_path = os.path.join(out_base, "summary.xml")
    ET.ElementTree(root).write(xml_path, encoding="utf-8", xml_declaration=True)


    print(f"[OK] done. out={out_base}")


if __name__ == "__main__":
    main()
