# -*- coding: utf-8 -*-
"""
trade_report_logger.py (B안)
- JSONL 파일 I/O 뿐 아니라,
    ✅ 이벤트 payload 표준화/계산(잔존시간, sell_price 확정, roi/pnl 계산, snapshot/mark payload 생성)
    ✅ cond 추론 로직(보유메타/인텐트/조건 편입 상태 기반)
    ✅ 부분체결 누적/평균체결가 계산(매도/매수)


Kiwoom 쪽은:
- “이런 이벤트가 발생했다”만 알려주고,
- recorder가 JSONL에 쓸 내용을 만들어서 기록한다.
"""


import os
import json
import datetime
from typing import Any, Dict, Optional, Set, Callable, DefaultDict
from collections import defaultdict




def _today_yyyymmdd() -> str:
    return datetime.datetime.now().strftime("%Y%m%d")




def _now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")




def _remain_to_close_1530() -> Dict[str, Any]:
    now_dt = datetime.datetime.now()
    close_dt = now_dt.replace(hour=15, minute=30, second=0, microsecond=0)
    remain_sec = max(0, int((close_dt - now_dt).total_seconds()))
    remain_min = float(remain_sec) / 60.0
    return {"account_remain_sec": remain_sec, "account_remain_min": remain_min}




class JsonlWriter:
    """파일 I/O만 담당(안전하게 append)."""
    def __init__(self, base_dir: str = "trade_logs", filename: Optional[str] = None):
        self.base_dir = base_dir
        self.filename = filename
        os.makedirs(self.base_dir, exist_ok=True)


    def _path(self) -> str:
        if self.filename:
            return os.path.join(self.base_dir, self.filename)
        return os.path.join(self.base_dir, f"{_today_yyyymmdd()}.jsonl")


    def write(self, event: str, payload: Optional[Dict[str, Any]] = None):
        rec = {"ts": _now_iso(), "event": str(event)}
        if payload:
            if isinstance(payload, dict):
                rec.update(payload)
            else:
                rec["payload"] = payload


        try:
            with open(self._path(), "a", encoding="utf-8") as f:
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        except Exception:
            # 기록 실패가 자동매매를 죽이면 안 됨
            pass




class ReportRecorder:
    """
    ✅ Kiwoom에서 최대한 뺀 버전
    - cond 추론 + sell_reason 표준화 + roi/pnl 계산 + 잔존시간 계산 + snapshot/mark payload 생성
    - 부분체결 누적: on_exec()로 누적하고, 보유수량 0 되는 순간 flush하여 sell_fill 기록
    """


    def __init__(
        self,
        writer: JsonlWriter,


        # 참조용(kiwoom 상태 dict를 그대로 넘겨서 recorder가 추론/계산에 사용)
        buy_cond_names: Set[str],
        sell_cond_names: Set[str],


        cond_current_members_ref: Dict[str, Set[str]],
        buy_meta_by_code_ref: Dict[str, Dict[str, Any]],
        pending_buy_intent_ref: Dict[str, Dict[str, Any]],
        pending_sell_intent_ref: Dict[str, Dict[str, Any]],
    ):
        self.w = writer


        self.buy_conds = set(buy_cond_names or set())
        self.sell_conds = set(sell_cond_names or set())


        self.cond_current_members = cond_current_members_ref
        self.buy_meta_by_code = buy_meta_by_code_ref
        self.pending_buy_intent = pending_buy_intent_ref
        self.pending_sell_intent = pending_sell_intent_ref


        # 부분체결 누적(매도/매수)
        self._sell_exec_agg: DefaultDict[str, Dict[str, int]] = defaultdict(lambda: {"qty": 0, "value": 0, "last_price": 0})
        self._buy_exec_agg: DefaultDict[str, Dict[str, int]] = defaultdict(lambda: {"qty": 0, "value": 0, "last_price": 0})


    # -------------------------
    # 공통: cond 추론
    # -------------------------
    def infer_cond(self, code: str) -> str:
        code = str(code or "").strip()
        if not code:
            return "UNKNOWN"


        meta = self.buy_meta_by_code.get(code, {}) or {}
        if meta.get("cond"):
            return str(meta["cond"])


        pi = self.pending_sell_intent.get(code, {}) or {}
        if pi.get("cond"):
            return str(pi["cond"])


        bi = self.pending_buy_intent.get(code, {}) or {}
        if bi.get("cond"):
            return str(bi["cond"])


        # 현재 조건 편입 상태 기반 추론
        for c in sorted(self.sell_conds):
            if code in (self.cond_current_members.get(c, set()) or set()):
                return c
        for c in sorted(self.buy_conds):
            if code in (self.cond_current_members.get(c, set()) or set()):
                return c


        return "UNKNOWN"


    # -------------------------
    # chejan exec 누적
    # -------------------------
    def on_exec(self, code: str, side: str, exec_qty: int, exec_price: int):
        code = str(code or "").strip()
        if not code:
            return
        if exec_qty <= 0 or exec_price <= 0:
            return


        side = str(side or "").upper() # "BUY" / "SELL"
        if side == "SELL":
            agg = self._sell_exec_agg[code]
        else:
            agg = self._buy_exec_agg[code]


        agg["qty"] += int(exec_qty)
        agg["value"] += int(exec_qty) * int(exec_price)
        agg["last_price"] = int(exec_price)


    # -------------------------
    # buy_fill 기록
    # -------------------------
    def record_buy_fill(self, code: str, name: str, qty: int, avg_price: int):
        code = str(code or "").strip()
        if not code or qty <= 0 or avg_price <= 0:
            return


        cond = self.infer_cond(code)


        # 보유메타 저장(추후 cond 추론 강화)
        self.buy_meta_by_code[code] = {"cond": cond, "buy_price": int(avg_price)}


        self.w.write("buy_fill", {
            "code": code,
            "name": str(name or "").strip(),
            "cond": cond,
            "buy_price": int(avg_price),
            "buy_qty": int(qty),
        })


    # -------------------------
    # sell intent 기록(사유/조건)
    # -------------------------
    def set_sell_intent(self, code: str, cond: str, sell_reason: str):
        code = str(code or "").strip()
        if not code:
            return
        self.pending_sell_intent[code] = {
            "cond": str(cond or ""),
            "sell_reason": str(sell_reason or "UNKNOWN"),
            "ts": _now_iso(),
        }


    # -------------------------
    # sell_fill 기록(보유수량 0 되는 순간 호출)
    # - sell_price는 부분체결 누적에서 평균 계산
    # - 없으면 price_fallback_getter(현재가/대체값)로 보정
    # -------------------------
    def record_sell_fill_closed(
        self,
        code: str,
        name: str,
        old_qty: int,
        old_avg: int,
        price_fallback_getter: Optional[Callable[[str], int]] = None,
    ):
        code = str(code or "").strip()
        if not code or old_qty <= 0:
            return


        # sell_price 결정
        agg = self._sell_exec_agg.get(code, {"qty": 0, "value": 0, "last_price": 0})
        sell_px = 0
        if int(agg.get("qty", 0)) > 0 and int(agg.get("value", 0)) > 0:
            sell_px = int(round(int(agg["value"]) / int(agg["qty"])))


        if sell_px <= 0 and price_fallback_getter:
            try:
                sell_px = int(price_fallback_getter(code) or 0)
            except:
                sell_px = 0


        if sell_px <= 0:
            sell_px = int(old_avg or 0)


        # agg flush
        try:
            self._sell_exec_agg.pop(code, None)
        except:
            pass


        # cond / reason
        intent = self.pending_sell_intent.get(code, {}) or {}
        cond = str(intent.get("cond") or "") or self.infer_cond(code)
        sell_reason = str(intent.get("sell_reason") or "") or "UNKNOWN"


        pnl = int((sell_px - int(old_avg or 0)) * int(old_qty)) if int(old_avg or 0) > 0 else 0
        roi = float((sell_px - int(old_avg or 0)) / int(old_avg) * 100.0) if int(old_avg or 0) > 0 else 0.0


        payload = {
            "code": code,
            "name": str(name or "").strip(),
            "cond": cond,
            "sell_reason": sell_reason,
            "sell_price": int(sell_px),
            "sell_qty": int(old_qty),
            "avg_price": int(old_avg or 0),
            "realized_pnl": int(pnl),
            "realized_roi": float(roi),
        }
        payload.update(_remain_to_close_1530())


        self.w.write("sell_fill", payload)


    # -------------------------
    # 30분 snapshot payload 생성/기록
    # -------------------------
    def record_condition_snapshot_30m(
        self,
        cond_current_members: Dict[str, Set[str]],
        cond_exited_today: Dict[str, Set[str]],
        cond_exit_repeat: Dict[str, Dict[str, int]],
        all_conditions: Set[str],
    ):
        conditions_payload: Dict[str, Any] = {}
        for cond in sorted(all_conditions):
            cur_set = cond_current_members.get(cond, set()) or set()
            exited_set = cond_exited_today.get(cond, set()) or set()
            rep_map = cond_exit_repeat.get(cond, {}) or {}


            dup_cnt = sum(1 for _c, n in rep_map.items() if int(n) >= 2)
            rep_total = sum(int(n) for n in rep_map.values())


            conditions_payload[cond] = {
                "current_count": int(len(cur_set)),
                "exited_today_count": int(len(exited_set)),
                "dup_exited_today_count": int(dup_cnt),
                "exit_repeat_total": int(rep_total),
            }


        self.w.write("condition_snapshot_30m", {"conditions": conditions_payload})


    # -------------------------
    # 30분 mark payload 생성/기록
    # -------------------------
    def record_mark_30m(
        self,
        holdings_qty: Dict[str, int],
        holdings_avg: Dict[str, int],
        price_getter: Callable[[str], int],
    ):
        items = []
        for code in sorted(list(holdings_qty.keys())):
            qty = int(holdings_qty.get(code, 0) or 0)
            if qty <= 0:
                continue


            avg = int(holdings_avg.get(code, 0) or 0)
            try:
                price = int(price_getter(code) or 0)
            except:
                price = 0
            if price <= 0:
                continue


            cond = self.infer_cond(code)


            roi = 0.0
            pnl = 0
            if avg > 0:
                roi = float((price - avg) / avg * 100.0)
                pnl = int((price - avg) * qty)


            items.append({
                "code": code,
                "cond": cond,
                "qty": qty,
                "avg_price": avg,
                "mark_price": int(price),
                "mark_roi": float(roi),
                "mark_pnl": int(pnl),
            })


        self.w.write("mark_30m", {"items": items})
