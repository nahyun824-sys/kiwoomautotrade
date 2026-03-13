# -*- coding: utf-8 -*-
"""Report helpers (file output)."""

import os
from typing import List, Dict, Any


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def write_daily_trade_report(report_dir: str, target_date, trades: List[Dict[str, Any]]) -> str:
    """Write a simple daily trade report txt file and return the filename."""
    ensure_dir(report_dir)
    filename = os.path.join(report_dir, f"trade_report_{target_date.strftime('%Y%m%d')}.txt")

    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"=== {target_date.strftime('%Y-%m-%d')} 매매 리포트 ===\n\n")
        for t in trades:
            f.write(f"{t['code']} {t['name']} profit={t['profit']} profit%={t['profit_rate']:.2f}%\n")

    return filename
