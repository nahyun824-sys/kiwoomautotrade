# -*- coding: utf-8 -*-
"""
Kiwoom OpenAPI+ 자동매매 통합본 (단일 파일) - B안(리포트 로깅 구조) 통합본

B안 핵심
- kiwoomautotrade(본 파일)는 "이벤트 발생"과 "주문/리스크 관리"만 수행
- ReportRecorder가 이벤트 payload 표준화 + 계산(가능한 범위) + JSONL 기록
- (오프라인 분석/그래프/XML 생성은 별도 analyze_report에서 수행하는 전제. 본 파일은 로그만 남김)

✅ 기존 핵심 패치 유지(중요)
1) opw00018(잔고) / opt10075(미체결) 연속조회(prev_next=2) 구현
2) 로컬 방어막(reserved_exposure): 잔고 누락/지연에도 MAX_POSITION_PER_CODE 초과매수 차단
3) 숫자 파서 강화(_safe_int): <a href="tel:...">..</a>, 콤마, 공백 등 처리

⚠️ 이번 수정(최소침습)
- 들여쓰기/스코프 깨져서 Kiwoom 메서드들이 전역/중첩 함수로 튀어나간 문제 복구
- stoploss_base_qty 누락 초기화 추가
"""

import sys
import os
import time
import json
import datetime
import re
from collections import deque, defaultdict
from typing import Dict, Any, Optional, Set, Tuple

from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer


# =========================
# 사용자 설정
# =========================
BUY_COND_NAMES = {"w3", "x2"}
SELL_COND_NAMES = {"w", "w3", "x2"}

TARGET_BUY_AMOUNT = 150000
MAX_POSITION_PER_CODE = 150000
ALLOW_ADD_BUY = False

SELL_DELAY_SEC = 5.0
REBUY_COOLDOWN_SEC = 600.0

STOPLOSS_TIERS = [(-3.0, 0.50), (-5.0, 0.50), (-7.0, 1.00)]
AUTO_SELL_INTERVAL_SEC = 60

TRAILING_STOP_PCT = 10.0
TRAILING_ENABLED = True

BALANCE_COOLDOWN_SEC = 3.0
SYNC_INTERVAL_SEC = 15.0
PRICE_REQ_INTERVAL = 0.25
PRICE_RETRY_MAX = 3
PRICE_RETRY_SLEEP = 0.8

ORPHAN_CHECK_INTERVAL_SEC = 10
ORPHAN_GRACE_SEC = 60

LOG_LEVEL = "INFO"  # "DEBUG"/"INFO"/"WARN"/"ERROR"
LOG_TO_FILE = True
LOG_FILE_PATH = "kiwoomautotrade.log"
LOG_THROTTLE_SEC = 2.0
LOG_PRINT_CODELIST_MAX = 8

# === B안: JSONL 이벤트 로깅 ===
REPORT_ENABLED = True
REPORT_DIR = "reports"
REPORT_PREFIX = "trade_events"  # trade_events_YYYYMMDD.jsonl
REPORT_FLUSH_EVERY_N = 1  # 1이면 매 이벤트 flush (안전). 성능 필요시 10~50 권장.

PASSWD = ""
PASSWD_MEDIA = "00"

SCREEN_TR_PRICE = "2000"
SCREEN_TR_BAL = "2100"
SCREEN_TR_UNFILLED = "2200"
SCREEN_COND_BASE = 5000

SOLD_TODAY_PERSIST = True
SOLD_TODAY_FILE = "sold_today.json"

BOUGHT_TODAY_PERSIST = True
BOUGHT_TODAY_FILE = "bought_today.json"


def _today_yyyymmdd() -> str:
    return datetime.datetime.now().strftime("%Y%m%d")


def _now_hms() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")


def _now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def _strip_html(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"<[^>]*>", "", s)
    return s.strip()


def _safe_int(x, default=0) -> int:
    try:
        s = _strip_html(x)
        if not s:
            return default
        s = s.replace(",", "").strip()
        m = re.search(r"[-+]?\d+", s)
        if not m:
            return default
        return int(m.group(0))
    except Exception:
        return default


def _norm_code(code: str) -> str:
    if code is None:
        return ""
    code = str(code).strip()
    if not code:
        return ""
    if code.startswith("A") and len(code) == 7:
        code = code[1:]
    return code


def _digits_only(s: str) -> str:
    s = _strip_html(s)
    m = re.findall(r"\d+", s)
    return "".join(m) if m else ""


# =========================
# B안: ReportRecorder (JSONL)
# =========================
class ReportRecorder:
    """
    - 이벤트를 JSONL로 남김(오프라인 분석용 원본)
    - payload는 가능한 한 표준화
    """
    def __init__(
        self,
        enabled: bool = True,
        report_dir: str = "reports",
        prefix: str = "trade_events",
        flush_every_n: int = 1,
        logger=None
    ):
        self.enabled = bool(enabled)
        self.report_dir = report_dir
        self.prefix = prefix
        self.flush_every_n = max(1, int(flush_every_n))
        self._fp = None
        self._day = ""
        self._count_since_flush = 0
        self._logger = logger  # callable(level, msg) optional

        if self.enabled:
            try:
                os.makedirs(self.report_dir, exist_ok=True)
            except Exception:
                self.enabled = False

    def _path_for_day(self, day: str) -> str:
        return os.path.join(self.report_dir, f"{self.prefix}_{day}.jsonl")

    def _ensure_open(self):
        if not self.enabled:
            return
        day = _today_yyyymmdd()
        if self._fp is None or self._day != day:
            try:
                if self._fp:
                    try:
                        self._fp.flush()
                        self._fp.close()
                    except Exception:
                        pass
                self._day = day
                self._fp = open(self._path_for_day(day), "a", encoding="utf-8")
                self._count_since_flush = 0
            except Exception as e:
                self.enabled = False
                if self._logger:
                    self._logger("WARN", f"[REPORT] open failed: {e}")

    def emit(self, event: str, payload: Dict[str, Any]):
        if not self.enabled:
            return
        self._ensure_open()
        if not self.enabled or not self._fp:
            return
        try:
            rec = {
                "ts": _now_iso(),
                "day": _today_yyyymmdd(),
                "event": str(event),
                "payload": payload or {},
            }
            self._fp.write(json.dumps(rec, ensure_ascii=False) + "\n")
            self._count_since_flush += 1
            if self._count_since_flush >= self.flush_every_n:
                self._fp.flush()
                self._count_since_flush = 0
        except Exception as e:
            if self._logger:
                self._logger("WARN", f"[REPORT] write failed: {e}")

    def close(self):
        try:
            if self._fp:
                self._fp.flush()
                self._fp.close()
        except Exception:
            pass
        self._fp = None


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        self.OnEventConnect.connect(self._on_event_connect)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)
        self.OnReceiveMsg.connect(self._on_receive_msg)

        self.OnReceiveConditionVer.connect(self._on_receive_condition_ver)
        self.OnReceiveTrCondition.connect(self._on_receive_tr_condition)
        self.OnReceiveRealCondition.connect(self._on_receive_real_condition)

        self.OnReceiveChejanData.connect(self._on_receive_chejan_data)

        self.login_loop = QEventLoop()
        self.tr_loop = QEventLoop()
        self.cond_loop = QEventLoop()

        # TR serial gate
        self._tr_busy = False
        self._tr_busy_name = ""
        self._tr_deadline = 0.0
        self._tr_timeout_timer = QTimer()
        self._tr_timeout_timer.timeout.connect(self._tr_timeout_tick)

        self.account = ""
        self.server_gubun = ""

        # conditions
        self.cond_name_to_idx: Dict[str, int] = {}
        self.subscribed_conds: Set[str] = set()
        self.sell_cond_members: Dict[str, Set[str]] = defaultdict(set)

        # holdings / orders
        self.holdings_qty: Dict[str, int] = defaultdict(int)
        self.holdings_avg: Dict[str, int] = defaultdict(int)
        self.holdings_name: Dict[str, str] = defaultdict(str)

        self.unfilled_orders: Dict[str, Dict[str, Any]] = {}
        self.pending_codes: Set[str] = set()  # any live unfilled for that code

        # ✅ 로컬 방어막: 잔고 누락 시에도 "추정 노출"로 한도 강제
        self.reserved_exposure: Dict[str, int] = defaultdict(int)

        # buy queue
        self.buy_queue = deque()
        self.buy_queue_set: Set[str] = set()

        # delayed sells
        self.delayed_sells: Dict[Tuple[str, str], float] = {}

        # sold/bought today
        self.sold_today: Set[str] = set()
        self.bought_today: Set[str] = set()

        self.rebuy_block_until: Dict[str, float] = {}
        self.rebuy_block_pending: Set[str] = set()

        # orphan
        self.orphan_first_seen: Dict[str, float] = {}

        # risk
        self.peak_price: Dict[str, int] = {}
        self.stoploss_stage: Dict[str, int] = defaultdict(int)
        self.stoploss_base_qty: Dict[str, int] = defaultdict(int)  # ✅ 누락 초기화
        self._risk_running = False

        # caching / price
        self.code_name_cache: Dict[str, str] = {}
        self._price_resp: Dict[str, int] = {}
        self._last_price_req_ts = 0.0
        self._last_balance_req_ts = 0.0
        self._last_price_req_code = ""

        # ✅ price request queue (TR busy 시 손절/리스크 누락 방지용)
        self._price_cache: Dict[str, Tuple[int, float]] = {}   # code -> (price, ts)
        self.price_req_queue = deque()
        self.price_req_set: Set[str] = set()
        self.price_queue_timer = QTimer()
        self.price_queue_timer.timeout.connect(self._process_price_queue_tick)

        # change signatures
        self._last_balance_sig = None
        self._last_unfilled_sig = None

        # ✅ 연속조회용 임시 버퍼
        self._bal_tmp_qty = defaultdict(int)
        self._bal_tmp_avg = defaultdict(int)
        self._bal_tmp_name = defaultdict(str)
        self._bal_is_accumulating = False

        self._unf_tmp = {}
        self._unf_is_accumulating = False

        self._init_logger()

        # B안 Recorder
        self.report = ReportRecorder(
            enabled=REPORT_ENABLED,
            report_dir=REPORT_DIR,
            prefix=REPORT_PREFIX,
            flush_every_n=REPORT_FLUSH_EVERY_N,
            logger=self._log,
        )

        self._log("INFO", "======================================================================")
        self._log("INFO", "[BOOT-CONFIG] 실제 실행중 설정값 확인 (이게 로그랑 다르면 '다른 파일 실행중'임)")
        self._log("INFO", f"[BOOT-CONFIG] __file__={__file__}")
        self._log("INFO", f"[BOOT-CONFIG] BUY_COND_NAMES={BUY_COND_NAMES}")
        self._log("INFO", f"[BOOT-CONFIG] SELL_COND_NAMES={SELL_COND_NAMES}")
        self._log("INFO", f"[BOOT-CONFIG] TARGET_BUY_AMOUNT={TARGET_BUY_AMOUNT}")
        self._log("INFO", f"[BOOT-CONFIG] MAX_POSITION_PER_CODE={MAX_POSITION_PER_CODE}")
        self._log("INFO", f"[BOOT-CONFIG] ALLOW_ADD_BUY={ALLOW_ADD_BUY}")
        self._log("INFO", f"[BOOT-CONFIG] SELL_DELAY_SEC={SELL_DELAY_SEC}")
        self._log("INFO", f"[BOOT-CONFIG] STOPLOSS_TIERS={STOPLOSS_TIERS} / AUTO_SELL_INTERVAL_SEC={AUTO_SELL_INTERVAL_SEC}")
        self._log("INFO", f"[BOOT-CONFIG] TRAILING_ENABLED={TRAILING_ENABLED} / TRAILING_STOP_PCT={TRAILING_STOP_PCT}")
        self._log("INFO", f"[BOOT-CONFIG] REPORT_ENABLED={REPORT_ENABLED} dir={REPORT_DIR} prefix={REPORT_PREFIX} flushN={REPORT_FLUSH_EVERY_N}")
        self._log("INFO", f"[BOOT-CONFIG] PASSWD={'(EMPTY)' if PASSWD=='' else '(SET)'} / PASSWD_MEDIA={PASSWD_MEDIA}")
        self._log("INFO", f"[BOOT-CONFIG] BOUGHT_TODAY_PERSIST={BOUGHT_TODAY_PERSIST} file={BOUGHT_TODAY_FILE}")
        self._log("INFO", "======================================================================")

        self.report.emit("BOOT_CONFIG", {
            "file": __file__,
            "BUY_COND_NAMES": sorted(list(BUY_COND_NAMES)),
            "SELL_COND_NAMES": sorted(list(SELL_COND_NAMES)),
            "TARGET_BUY_AMOUNT": TARGET_BUY_AMOUNT,
            "MAX_POSITION_PER_CODE": MAX_POSITION_PER_CODE,
            "ALLOW_ADD_BUY": ALLOW_ADD_BUY,
            "SELL_DELAY_SEC": SELL_DELAY_SEC,
            "STOPLOSS_TIERS": STOPLOSS_TIERS,
            "AUTO_SELL_INTERVAL_SEC": AUTO_SELL_INTERVAL_SEC,
            "TRAILING_ENABLED": TRAILING_ENABLED,
            "TRAILING_STOP_PCT": TRAILING_STOP_PCT,
        })

        if SOLD_TODAY_PERSIST:
            self._load_sold_today()
        if BOUGHT_TODAY_PERSIST:
            self._load_bought_today()

        # timers
        self.buy_timer = QTimer()
        self.buy_timer.timeout.connect(self._process_buy_queue_tick)

        self.delay_sell_timer = QTimer()
        self.delay_sell_timer.timeout.connect(self._process_delayed_sells_tick)

        self.orphan_timer = QTimer()
        self.orphan_timer.timeout.connect(self._orphan_sweeper_tick)

        self.sync_timer = QTimer()
        self.sync_timer.timeout.connect(self._periodic_sync_tick)

        self.risk_timer = QTimer()
        self.risk_timer.timeout.connect(self._risk_monitor_tick)

    # -----------------------
    # Logger
    # -----------------------
    def _init_logger(self):
        self._log_level_map = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}
        self._log_level = self._log_level_map.get(str(LOG_LEVEL).upper(), 20)
        self._log_last_ts = {}
        self._log_fp = None
        if LOG_TO_FILE:
            try:
                self._log_fp = open(LOG_FILE_PATH, "a", encoding="utf-8")
            except Exception as e:
                print(f"[LOGGER] 파일 로그 오픈 실패: {e}")

    def _log(self, level: str, msg: str, key: str = None, throttle: float = None):
        lv = self._log_level_map.get(level, 20)
        if lv < self._log_level:
            return
        now = time.time()
        if key:
            th = LOG_THROTTLE_SEC if throttle is None else throttle
            last = self._log_last_ts.get(key, 0.0)
            if th and (now - last) < th:
                return
            self._log_last_ts[key] = now
        line = f"[{_now_hms()}] [{level}] {msg}"
        print(line)
        if self._log_fp:
            try:
                self._log_fp.write(line + "\n")
                self._log_fp.flush()
            except Exception:
                pass

    # -----------------------
    # Util
    # -----------------------
    def _get_code_name(self, code: str) -> str:
        code = _norm_code(code)
        if not code:
            return ""
        nm = str(self.holdings_name.get(code, "")).strip()
        if nm:
            self.code_name_cache[code] = nm
            return nm
        nm = str(self.code_name_cache.get(code, "")).strip()
        if nm:
            return nm
        try:
            nm = str(self.dynamicCall("GetMasterCodeName(QString)", code)).strip()
        except Exception:
            nm = ""
        if nm:
            self.code_name_cache[code] = nm
        return nm

    def _code_tag(self, code: str) -> str:
        code = _norm_code(code)
        if not code:
            return ""
        nm = self._get_code_name(code)
        return f"{code}({nm})" if nm else code

    def _fmt_codelist(self, codes):
        codes = list(codes) if codes else []
        n = len(codes)
        maxn = int(LOG_PRINT_CODELIST_MAX)
        if maxn <= 0:
            return f"{n} codes"
        head = ", ".join([self._code_tag(c) for c in codes[:maxn]])
        more = "" if n <= maxn else f" ... (+{n - maxn})"
        return f"{n} codes: {head}{more}"

    # -----------------------
    # Price queue helpers (TR busy 시 다음 tick 처리)
    # -----------------------
    def _enqueue_price_request(self, code: str, why: str = ""):
        code = _norm_code(code)
        if not code:
            return
        if code in self.price_req_set:
            return
        self.price_req_queue.append(code)
        self.price_req_set.add(code)
        self._log(
            "DEBUG",
            f"[PRICE-QUEUE] add: {self._code_tag(code)} why={why} qlen={len(self.price_req_queue)}",
            key=f"PRICEQ_ADD_{code}",
            throttle=0.5,
        )
        self.report.emit("PRICE_QUEUE_ADD", {"code": code, "why": why, "queue_len": len(self.price_req_queue)})

    def _request_price_tr_now(self, code: str) -> Optional[int]:
        """TR을 즉시 호출해서 현재가를 받아옴(큐잉/바쁨 체크 없이)."""
        code = _norm_code(code)
        if not code:
            return None

        now = time.time()
        gap = now - self._last_price_req_ts
        if gap < PRICE_REQ_INTERVAL:
            time.sleep(PRICE_REQ_INTERVAL - gap)
        self._last_price_req_ts = time.time()

        self._price_resp.pop(code, None)
        self._last_price_req_code = code
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        for attempt in range(1, PRICE_RETRY_MAX + 1):
            ok = self._tr_request("opt10001_req", "opt10001", SCREEN_TR_PRICE, prev_next=0, timeout_sec=6.0)
            if ok:
                price = int(self._price_resp.get(code, 0) or 0)
                if price > 0:
                    self._price_cache[code] = (price, time.time())
                    self.report.emit("PRICE_OK", {"code": code, "price": price, "via": "TR_NOW"})
                    return price
            else:
                self._log(
                    "WARN",
                    f"[PRICE] TR요청 스킵/실패 attempt={attempt} code={self._code_tag(code)}",
                    key=f"PRICE_FAIL_{code}",
                    throttle=0.5,
                )
                self.report.emit("PRICE_FAIL", {"code": code, "attempt": attempt, "via": "TR_NOW"})
            time.sleep(PRICE_RETRY_SLEEP)

        return None

    def _process_price_queue_tick(self):
        if self._risk_running or self._tr_busy:
            return
        if not self.price_req_queue:
            return

        code = self.price_req_queue.popleft()
        self.price_req_set.discard(code)

        cached = self._price_cache.get(code)
        if cached:
            price, ts = cached
            if (time.time() - float(ts)) <= 2.0 and int(price) > 0:
                return

        price = self._request_price_tr_now(code)
        if not price:
            self._enqueue_price_request(code, why="RETRY_AFTER_FAIL")

    # -----------------------
    # TR serial
    # -----------------------
    def _tr_request(self, rq_name: str, tr_code: str, screen: str, prev_next: int = 0, timeout_sec: float = 8.0) -> bool:
        if self._tr_busy:
            end = time.time() + 0.8
            while self._tr_busy and time.time() < end:
                time.sleep(0.02)
            if self._tr_busy:
                self._log(
                    "WARN",
                    f"[TR-SERIAL] busy({self._tr_busy_name}) -> skip rq={rq_name}",
                    key=f"TR_BUSY_{rq_name}",
                    throttle=0.5,
                )
                self.report.emit("TR_SKIP_BUSY", {"rq": rq_name, "busy_rq": self._tr_busy_name})
                return False

        self._tr_busy = True
        self._tr_busy_name = rq_name
        self._tr_deadline = time.time() + float(timeout_sec)

        ret = self.dynamicCall("CommRqData(QString, QString, int, QString)", rq_name, tr_code, int(prev_next), screen)
        if ret != 0:
            self._log("WARN", f"[TR-SERIAL] CommRqData fail ret={ret} rq={rq_name}", key=f"TR_FAIL_{rq_name}", throttle=0.5)
            self.report.emit("TR_FAIL", {"rq": rq_name, "tr": tr_code, "ret": ret})
            self._tr_release()
            return False

        self._tr_timeout_timer.start(50)
        self.tr_loop.exec_()
        return not self._tr_busy

    def _tr_release(self):
        self._tr_busy = False
        self._tr_busy_name = ""
        self._tr_deadline = 0.0
        try:
            self._tr_timeout_timer.stop()
        except Exception:
            pass

    def _tr_timeout_tick(self):
        if not self._tr_busy:
            try:
                self._tr_timeout_timer.stop()
            except Exception:
                pass
            return
        if time.time() >= self._tr_deadline:
            self._log("ERROR", f"[TR-SERIAL] timeout rq={self._tr_busy_name} -> force release", key=f"TR_TO_{self._tr_busy_name}", throttle=0.5)
            self.report.emit("TR_TIMEOUT", {"rq": self._tr_busy_name})
            self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()

    # -----------------------
    # Login
    # -----------------------
    def comm_connect(self):
        self._log("INFO", f"[LOGIN] GetConnectState={self.dynamicCall('GetConnectState()')} (1=연결,0=미연결)")
        self._log("INFO", "[LOGIN] CommConnect() 호출")
        self.dynamicCall("CommConnect()")
        self.login_loop.exec_()

    def _on_event_connect(self, err_code):
        self._log("INFO", f"[EVENT] OnEventConnect err_code={err_code}")
        if err_code == 0:
            self._log("INFO", "[LOGIN] ✅ 로그인 성공")
            acc_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")
            accs = [a for a in str(acc_list).split(";") if a]
            self._log("INFO", f"[LOGIN] ACCNO(list)={accs}")

            self.account = _digits_only(accs[0]) if accs else ""
            self._log("INFO", f"[LOGIN] 계좌번호: {self.account if self.account else '(EMPTY)'}")

            self.server_gubun = str(self.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")).strip()
            self._log("INFO", f"[LOGIN] 서버구분(1=모의, 0=실): {self.server_gubun}")

            self.report.emit("LOGIN_OK", {"account": self.account, "server_gubun": self.server_gubun})
        else:
            self._log("ERROR", f"[LOGIN] ❌ 로그인 실패 err_code={err_code}")
            self.report.emit("LOGIN_FAIL", {"err_code": err_code})

        if self.login_loop.isRunning():
            self.login_loop.exit()

    # -----------------------
    # Conditions
    # -----------------------
    def load_conditions(self):
        ret = self.dynamicCall("GetConditionLoad()")
        if ret == 0:
            self._log("ERROR", "[COND] GetConditionLoad() 실패")
            self.report.emit("COND_LOAD_FAIL", {})
            return
        self.cond_loop.exec_()

    def _on_receive_condition_ver(self, ret, msg):
        self._log("INFO", "[COND] 조건 목록 로드 완료")
        raw = str(self.dynamicCall("GetConditionNameList()"))
        items = [x for x in raw.split(";") if x]
        self.cond_name_to_idx.clear()
        for it in items:
            try:
                idx_str, name = it.split("^", 1)
                self.cond_name_to_idx[name] = int(idx_str)
            except Exception:
                continue

        buy_idx = {self.cond_name_to_idx.get(n) for n in BUY_COND_NAMES if n in self.cond_name_to_idx}
        sell_idx = {self.cond_name_to_idx.get(n) for n in SELL_COND_NAMES if n in self.cond_name_to_idx}
        buy_idx.discard(None)
        sell_idx.discard(None)

        self._log("INFO", f"[COND] BUY_CONDITIONS={BUY_COND_NAMES} -> idx={buy_idx}")
        self._log("INFO", f"[COND] SELL_CONDITIONS={SELL_COND_NAMES} -> idx={sell_idx}")

        self.report.emit("COND_LIST", {
            "conds_total": len(self.cond_name_to_idx),
            "buy": sorted(list(BUY_COND_NAMES)),
            "sell": sorted(list(SELL_COND_NAMES)),
            "buy_idx": sorted([int(x) for x in buy_idx]),
            "sell_idx": sorted([int(x) for x in sell_idx]),
        })

        if self.cond_loop.isRunning():
            self.cond_loop.exit()

    def subscribe_conditions(self):
        all_conds = list(sorted(set(SELL_COND_NAMES) | set(BUY_COND_NAMES)))
        ordered = []
        for c in sorted(SELL_COND_NAMES):
            if c in all_conds:
                ordered.append(c)
        for c in sorted(BUY_COND_NAMES):
            if c in all_conds and c not in ordered:
                ordered.append(c)

        for name in ordered:
            if name not in self.cond_name_to_idx:
                self._log("WARN", f"[COND] 조건명 없음 -> 스킵: {name}")
                continue
            if name in self.subscribed_conds:
                continue
            idx = self.cond_name_to_idx[name]
            scr = str(SCREEN_COND_BASE + idx)
            ok = self._send_condition_retry(name, idx, scr, search=1)
            if ok:
                self.subscribed_conds.add(name)

    def _send_condition_retry(self, name: str, idx: int, scr: str, search: int) -> bool:
        for i in range(1, 6):
            ret = self.dynamicCall("SendCondition(QString, QString, int, int)", scr, name, idx, search)
            self._log("INFO", f"[COND-SUB] 구독 시도: {name} idx={idx} scr={scr} ret={ret} (try {i}/5)", key=f"COND_SUB_{name}", throttle=0.2)
            if ret == 1:
                self._log("INFO", f"[COND-SUB] ✅ 구독 성공: {name}")
                self.report.emit("COND_SUB_OK", {"cond": name, "idx": idx, "screen": scr})
                return True
            time.sleep(0.2)
        self._log("ERROR", f"[COND-SUB] ❌ 구독 실패: {name}")
        self.report.emit("COND_SUB_FAIL", {"cond": name, "idx": idx, "screen": scr})
        return False

    def _on_receive_tr_condition(self, scr_no, code_list, cond_name, cond_index, next_):
        codes = [c for c in str(code_list).split(";") if c]
        cond_name = str(cond_name)

        self._log("INFO", f"[TRCOND] scr={scr_no} cond={cond_name} idx={cond_index} next={next_}")
        self._log("INFO", f"[TRCOND] {self._fmt_codelist(codes)}", key=f"TRCOND_{cond_name}", throttle=0.5)

        self.report.emit("COND_TR_SNAPSHOT", {
            "cond": cond_name, "idx": int(cond_index), "count": len(codes),
            "codes_sample": list(map(_norm_code, codes[:min(50, len(codes))])),
        })

        if cond_name in SELL_COND_NAMES:
            self.sell_cond_members[cond_name] = set(map(_norm_code, codes))
            self._log("INFO", f"[SELL-COND-INIT] cond={cond_name} -> {self._fmt_codelist(self.sell_cond_members[cond_name])}",
                      key=f"SELLINIT_{cond_name}", throttle=0.5)

        if cond_name in BUY_COND_NAMES:
            for c in codes:
                self._enqueue_buy(_norm_code(c), cond_name, ev="INITIAL_TRCOND")

    def _on_receive_real_condition(self, code, event_type, cond_name, cond_index):
        code = _norm_code(code)
        cond_name = str(cond_name)
        event_type = str(event_type)
        if not code:
            return

        if event_type == "I":
            self._log("INFO", f"[COND-REAL] cond={cond_name} 편입(I): {self._code_tag(code)}", key=f"COND_I_{cond_name}_{code}", throttle=0.5)
            self.report.emit("COND_REAL_I", {"cond": cond_name, "code": code, "idx": int(cond_index)})
            if cond_name in SELL_COND_NAMES:
                self.sell_cond_members[cond_name].add(code)
            if cond_name in BUY_COND_NAMES:
                self._enqueue_buy(code, cond_name, ev="REAL_I")

        elif event_type == "D":
            self._log("INFO", f"[COND-REAL] cond={cond_name} 이탈(D): {self._code_tag(code)}", key=f"COND_D_{cond_name}_{code}", throttle=0.5)
            self.report.emit("COND_REAL_D", {"cond": cond_name, "code": code, "idx": int(cond_index)})
            if cond_name in SELL_COND_NAMES:
                self.sell_cond_members[cond_name].discard(code)

                if code in self.bought_today:
                    self._log("INFO", f"[SELL-DELAY] 스킵(당일매수 조건매도 제외): {self._code_tag(code)} cond={cond_name}",
                              key=f"SELL_DELAY_SKIP_BOUGHT_{code}_{cond_name}", throttle=0.0)
                    self.report.emit("SELL_DELAY_SKIP_BOUGHT_TODAY", {"code": code, "cond": cond_name})
                    return

                self._schedule_delayed_sell(code, cond_name, reason="COND_EXIT")

    # -----------------------
    # TR Receive
    # -----------------------
    def _on_receive_tr_data(self, scr_no, rq_name, tr_code, record_name, prev_next, data_len, err_code, msg1, msg2):
        rq_name = str(rq_name)
        prev_next = str(prev_next).strip()

        if rq_name == "opt10001_req":
            code = _norm_code(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "종목코드"))
            if not code:
                code = _norm_code(self._last_price_req_code)
            price_raw = self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "현재가")
            price = abs(_safe_int(price_raw, 0))
            if code:
                self._price_resp[code] = price
            self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()
            return

        # ✅ 잔고 연속조회
        if rq_name == "opw00018_req":
            self._parse_balance_page(tr_code, rq_name, is_first=not self._bal_is_accumulating)

            if prev_next == "2":
                self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
                self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", PASSWD)
                self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", PASSWD_MEDIA)
                self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
                ret = self.dynamicCall("CommRqData(QString, QString, int, QString)", "opw00018_req", "opw00018", 2, SCREEN_TR_BAL)
                if ret != 0:
                    self._log("WARN", f"[TR-SERIAL] CommRqData fail(ret={ret}) during BALANCE next-page", key="BAL_NEXT_FAIL", throttle=0.5)
                    self.report.emit("BAL_NEXT_FAIL", {"ret": ret})
                    self._finalize_balance_accum()
                    self._tr_release()
                    if self.tr_loop.isRunning():
                        self.tr_loop.exit()
                else:
                    self._bal_is_accumulating = True
                return
            else:
                self._finalize_balance_accum()
                self._tr_release()
                if self.tr_loop.isRunning():
                    self.tr_loop.exit()
                return

        # ✅ 미체결 연속조회(안전)
        if rq_name == "opt10075_req":
            self._parse_unfilled_page(tr_code, rq_name, is_first=not self._unf_is_accumulating)

            if prev_next == "2":
                self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
                self.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
                self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
                self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
                ret = self.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10075_req", "opt10075", 2, SCREEN_TR_UNFILLED)
                if ret != 0:
                    self._log("WARN", f"[TR-SERIAL] CommRqData fail(ret={ret}) during UNFILLED next-page", key="UNF_NEXT_FAIL", throttle=0.5)
                    self.report.emit("UNF_NEXT_FAIL", {"ret": ret})
                    self._finalize_unfilled_accum()
                    self._tr_release()
                    if self.tr_loop.isRunning():
                        self.tr_loop.exit()
                else:
                    self._unf_is_accumulating = True
                return
            else:
                self._finalize_unfilled_accum()
                self._tr_release()
                if self.tr_loop.isRunning():
                    self.tr_loop.exit()
                return

        if self._tr_busy:
            self._tr_release()
        if self.tr_loop.isRunning():
            self.tr_loop.exit()

    def _on_receive_msg(self, scr_no, rq_name, tr_code, msg):
        self._log("DEBUG", f"[MSG] scr={scr_no} rq={rq_name} tr={tr_code} msg={msg}", key="MSG", throttle=1.0)

    # -----------------------
    # TR Request
    # -----------------------
    def request_price(self, code: str) -> Optional[int]:
        code = _norm_code(code)
        if not code:
            return None

        cached = self._price_cache.get(code)
        if cached:
            price, ts = cached
            if (time.time() - float(ts)) <= 2.0 and int(price) > 0:
                return int(price)

        if self._tr_busy:
            self._enqueue_price_request(code, why="TR_BUSY")
            self.report.emit("PRICE_QUEUED", {"code": code, "why": "TR_BUSY"})
            return None

        return self._request_price_tr_now(code)

    def request_balance(self):
        now = time.time()
        if now - self._last_balance_req_ts < BALANCE_COOLDOWN_SEC:
            return
        self._last_balance_req_ts = now

        self._bal_tmp_qty = defaultdict(int)
        self._bal_tmp_avg = defaultdict(int)
        self._bal_tmp_name = defaultdict(str)
        self._bal_is_accumulating = False

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", PASSWD)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", PASSWD_MEDIA)
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.report.emit("BAL_REQ", {})
        self._tr_request("opw00018_req", "opw00018", SCREEN_TR_BAL, prev_next=0, timeout_sec=12.0)

    def request_unfilled(self):
        self._unf_tmp = {}
        self._unf_is_accumulating = False

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
        self.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.report.emit("UNF_REQ", {})
        self._tr_request("opt10075_req", "opt10075", SCREEN_TR_UNFILLED, prev_next=0, timeout_sec=12.0)

    # -----------------------
    # Balance / Unfilled parse (paged)
    # -----------------------
    def _get_comm_data_multi(self, tr_code, rq_name, i: int, fields):
        for f in fields:
            try:
                v = self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, f)
                if str(v).strip():
                    return v, f
            except Exception:
                pass
        return "", ""

    def _parse_balance_page(self, tr_code, rq_name, is_first: bool):
        cnt = _safe_int(self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name), 0)
        for i in range(cnt):
            code = _norm_code(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목번호"))
            name = str(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목명")).strip()
            qty = _safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "보유수량"), 0)

            avg_raw, avg_field = self._get_comm_data_multi(tr_code, rq_name, i, ["평균단가", "매입가", "매입단가", "평균매입단가"])
            avg = abs(_safe_int(avg_raw, 0))

            if code and qty > 0:
                self._bal_tmp_qty[code] = qty
                self._bal_tmp_avg[code] = avg
                self._bal_tmp_name[code] = name
                if name:
                    self.code_name_cache[code] = name

            if qty > 0 and avg <= 0:
                self._log(
                    "WARN",
                    f"[BALANCE] avg=0 감지: {code}({name}) field={avg_field or 'N/A'} raw='{_strip_html(avg_raw)}'",
                    key=f"AVG0_{code}",
                    throttle=30.0,
                )
                self.report.emit("BAL_AVG0_WARN", {"code": code, "name": name, "field": avg_field, "raw": _strip_html(avg_raw)})

    def _finalize_balance_accum(self):
        new_qty = self._bal_tmp_qty
        new_avg = self._bal_tmp_avg
        new_name = self._bal_tmp_name

        new_sig = tuple((c, int(new_qty.get(c, 0)), int(new_avg.get(c, 0))) for c in sorted(new_qty.keys()))
        changed = (new_sig != self._last_balance_sig)
        self._last_balance_sig = new_sig

        self.holdings_qty = new_qty
        self.holdings_avg = new_avg
        self.holdings_name = new_name

        for code in list(self.peak_price.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.peak_price.pop(code, None)
        for code in list(self.stoploss_stage.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.stoploss_stage.pop(code, None)
        for code in list(self.stoploss_base_qty.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.stoploss_base_qty.pop(code, None)
        for code in list(self.rebuy_block_pending):
            if self.holdings_qty.get(code, 0) <= 0:
                self.rebuy_block_pending.discard(code)

        if not changed:
            return

        self._log(
            "INFO",
            f"[BALANCE] 변경감지: 보유종목({len(self.holdings_qty)}): "
            + "{%s}" % ", ".join([f"{self._code_tag(c)}={q}" for c, q in list(self.holdings_qty.items())[:12]])
            + (" ..." if len(self.holdings_qty) > 12 else ""),
            key="BAL_CHANGED",
            throttle=0.0,
        )

        sample = []
        for c, q in list(self.holdings_qty.items())[:50]:
            sample.append({"code": c, "qty": int(q), "avg": int(self.holdings_avg.get(c, 0) or 0)})
        self.report.emit("BAL_SNAPSHOT", {"count": len(self.holdings_qty), "sample": sample})

    def _parse_unfilled_page(self, tr_code, rq_name, is_first: bool):
        cnt = _safe_int(self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name), 0)
        for i in range(cnt):
            order_no = str(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "주문번호")).strip()
            code = _norm_code(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목코드"))
            bs = str(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "매매구분")).strip()
            qty = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "주문수량"), 0))
            unfilled = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "미체결수량"), 0))
            price = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "주문가격"), 0))
            if order_no and code and unfilled > 0:
                self._unf_tmp[order_no] = {"code": code, "bs": bs, "qty": qty, "unfilled": unfilled, "price": price}

    def _finalize_unfilled_accum(self):
        new_unfilled = self._unf_tmp
        new_sig = tuple(
            (str(o), _norm_code(v.get("code", "")), str(v.get("bs", "")),
             int(v.get("qty", 0) or 0), int(v.get("unfilled", 0) or 0), int(v.get("price", 0) or 0))
            for o, v in sorted(new_unfilled.items(), key=lambda x: x[0])
        )
        changed = (new_sig != self._last_unfilled_sig)
        self._last_unfilled_sig = new_sig

        self.unfilled_orders = new_unfilled
        live_pending_codes = set()
        for od in self.unfilled_orders.values():
            live_pending_codes.add(od["code"])
        self.pending_codes = live_pending_codes

        if not changed:
            return

        sample = []
        for order_no in list(sorted(self.unfilled_orders.keys()))[:50]:
            od = self.unfilled_orders[order_no]
            sample.append({
                "order_no": order_no,
                "code": od.get("code"),
                "bs": od.get("bs"),
                "unfilled": int(od.get("unfilled", 0) or 0),
                "price": int(od.get("price", 0) or 0),
            })

        self._log("INFO", f"[UNFILLED] 변경감지: 미체결 {len(self.unfilled_orders)}건", key="UNF_CHANGED", throttle=0.0)
        self.report.emit("UNF_SNAPSHOT", {"count": len(self.unfilled_orders), "sample": sample})

    # -----------------------
    # Chejan
    # -----------------------
    def _on_receive_chejan_data(self, gubun, item_cnt, fid_list):
        gubun = str(gubun)

        if gubun == "0":
            order_no = str(self.dynamicCall("GetChejanData(int)", 9203)).strip()
            code = _norm_code(self.dynamicCall("GetChejanData(int)", 9001))
            bs = str(self.dynamicCall("GetChejanData(int)", 907)).strip()
            unfilled = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 902), 0))
            qty = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 900), 0))

            if order_no and code:
                if unfilled > 0:
                    self.unfilled_orders[order_no] = {"code": code, "bs": bs, "qty": qty, "unfilled": unfilled, "price": 0}
                    self.pending_codes.add(code)
                else:
                    self.unfilled_orders.pop(order_no, None)
                    still = any(v["code"] == code and v.get("unfilled", 0) > 0 for v in self.unfilled_orders.values())
                    if not still:
                        self.pending_codes.discard(code)

                try:
                    self._last_unfilled_sig = tuple(
                        (str(o), _norm_code(v.get("code", "")), str(v.get("bs", "")),
                         int(v.get("qty", 0) or 0), int(v.get("unfilled", 0) or 0), int(v.get("price", 0) or 0))
                        for o, v in sorted(self.unfilled_orders.items(), key=lambda x: x[0])
                    )
                except Exception:
                    pass

                self.report.emit("CHEJAN_ORDER", {"order_no": order_no, "code": code, "bs": bs, "qty": qty, "unfilled": unfilled})

        elif gubun == "1":
            code = _norm_code(self.dynamicCall("GetChejanData(int)", 9001))
            name = str(self.dynamicCall("GetChejanData(int)", 302)).strip()
            qty = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 930), 0))
            avg = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 931), 0))

            if code:
                before_qty = int(self.holdings_qty.get(code, 0) or 0)

                if qty > 0:
                    self.holdings_qty[code] = qty
                    if avg > 0:
                        self.holdings_avg[code] = avg
                    self.holdings_name[code] = name
                    if name:
                        self.code_name_cache[code] = name

                    if code not in self.peak_price:
                        base = self.holdings_avg.get(code, 0) or 0
                        if base > 0:
                            self.peak_price[code] = base

                    if before_qty <= 0 and code not in self.bought_today:
                        self.bought_today.add(code)
                        self._log("INFO", f"[BOUGHT-TODAY] ✅ 신규매수 감지(체잔) -> 당일매수 등록: {self._code_tag(code)}",
                                  key=f"BOUGHT_ADD_{code}", throttle=0.0)
                        if BOUGHT_TODAY_PERSIST:
                            self._save_bought_today()

                    self.report.emit("CHEJAN_HOLDING", {"code": code, "name": name, "qty": qty, "avg": avg})

                else:
                    self.holdings_qty.pop(code, None)
                    self.holdings_avg.pop(code, None)
                    self.holdings_name.pop(code, None)
                    self.rebuy_block_pending.discard(code)
                    self.peak_price.pop(code, None)
                    self.stoploss_stage.pop(code, None)
                    self.stoploss_base_qty.pop(code, None)

                    if code and (code not in self.sold_today):
                        self.sold_today.add(code)
                        self._log("INFO", f"[SOLD-TODAY] ✅ 체잔 qty=0 확정 -> 당일 재매수 금지 등록: {self._code_tag(code)}",
                                  key=f"SOLD_TODAY_ADD_{code}", throttle=0.0)
                        if SOLD_TODAY_PERSIST:
                            self._save_sold_today()

                    self.report.emit("CHEJAN_HOLDING_ZERO", {"code": code, "name": name})

    # -----------------------
    # Buy / Sell core
    # -----------------------
    def _enqueue_buy(self, code: str, cond_name: str, ev: str = ""):
        code = _norm_code(code)
        if not code:
            return
        if code in self.buy_queue_set:
            return
        if code in self.sold_today:
            self._log("INFO", f"[BUY-QUEUE] 스킵: 재매수 금지 -> {self._code_tag(code)}", key=f"BUYQ_SOLD_{code}", throttle=1.0)
            self.report.emit("BUYQ_SKIP_SOLD_TODAY", {"code": code, "cond": cond_name, "ev": ev})
            return
        self.buy_queue.append((code, cond_name, time.time()))
        self.buy_queue_set.add(code)
        self._log("INFO", f"[BUY-QUEUE] 추가: {self._code_tag(code)} cond={cond_name} queue_len={len(self.buy_queue)}",
                  key="BUYQ_ADD", throttle=0.2)
        self.report.emit("BUYQ_ADD", {"code": code, "cond": cond_name, "ev": ev, "queue_len": len(self.buy_queue)})

    def _send_order(self, name: str, screen: str, acc: str, order_type: int,
                    code: str, qty: int, price: int, hoga: str, org_order_no: str) -> int:
        sig = "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)"
        args = [name, screen, acc, int(order_type), code, int(qty), int(price), hoga, org_order_no]
        return self.dynamicCall(sig, args)

    def _estimate_exposure_code(self, code: str, price_hint: int = 0) -> int:
        code = _norm_code(code)
        hold_qty = int(self.holdings_qty.get(code, 0) or 0)
        hold_avg = int(self.holdings_avg.get(code, 0) or 0)
        hold_value = hold_qty * (hold_avg if hold_avg > 0 else price_hint)

        pending_buy_value = 0
        for od in self.unfilled_orders.values():
            if od.get("code") != code:
                continue
            bs = str(od.get("bs", ""))
            if "매수" in bs or bs.startswith("+"):
                p = int(od.get("price", 0) or price_hint)
                pending_buy_value += int(od.get("unfilled", 0)) * p

        queued_budget = 0
        for (c, _cond, _ts) in self.buy_queue:
            if c == code:
                queued_budget += TARGET_BUY_AMOUNT

        reserved = int(self.reserved_exposure.get(code, 0) or 0)
        return int(hold_value + pending_buy_value + queued_budget + reserved)

    def _can_buy_code(self, code: str, price: int) -> Tuple[bool, str]:
        code = _norm_code(code)
        if not code:
            return False, "INVALID_CODE"

        if code in self.sold_today:
            return False, "SOLD_TODAY_BLOCK"

        until = self.rebuy_block_until.get(code, 0)
        if time.time() < until:
            return False, "REBUY_COOLDOWN"

        if code in self.rebuy_block_pending:
            return False, "REBUY_PENDING_QTY0"

        if (not ALLOW_ADD_BUY) and (code in self.bought_today):
            return False, "BOUGHT_TODAY_NO_ADD"

        if self.holdings_qty.get(code, 0) > 0 and not ALLOW_ADD_BUY:
            return False, "ALREADY_HOLDING"

        if code in self.pending_codes:
            return False, "PENDING_ORDER"

        exposure = self._estimate_exposure_code(code, price_hint=price)
        if exposure + TARGET_BUY_AMOUNT > MAX_POSITION_PER_CODE:
            return False, f"MAX_POSITION_PER_CODE(exposure={exposure})"

        if price <= 0 or TARGET_BUY_AMOUNT < price:
            return False, f"INSUFFICIENT_FUNDS_1SHARE(price={price})"

        return True, "OK"

    def _process_buy_queue_tick(self):
        if not self.buy_queue:
            return
        if self._risk_running:
            return

        code, cond, ts = self.buy_queue.popleft()
        self.buy_queue_set.discard(code)

        self._log("INFO", f"[BUY] 진입: target={self._code_tag(code)} budget={TARGET_BUY_AMOUNT} cond={cond}",
                  key=f"BUY_ENTER_{code}", throttle=0.2)

        price = self.request_price(code)
        if not price:
            self._log("WARN", f"[BUY] 현재가 실패 -> 스킵: {self._code_tag(code)}", key=f"BUY_NOPRICE_{code}", throttle=1.0)
            self.report.emit("BUY_SKIP_NO_PRICE", {"code": code, "cond": cond})
            return

        ok, reason = self._can_buy_code(code, price)
        if not ok:
            self._log("INFO", f"[BUY] 스킵: {self._code_tag(code)} reason={reason}", key=f"BUY_BLOCK_{code}", throttle=0.5)
            self.report.emit("BUY_BLOCK", {"code": code, "cond": cond, "reason": reason, "price": price})
            return

        qty = TARGET_BUY_AMOUNT // price
        if qty <= 0:
            self.report.emit("BUY_SKIP_QTY0", {"code": code, "cond": cond, "price": price})
            return

        order_amount = qty * price
        self._log("INFO", f"[BUY] 계산: {self._code_tag(code)} price={price} qty={qty} order_amount={order_amount}",
                  key=f"BUY_CALC_{code}", throttle=0.2)

        self.report.emit("BUY_ATTEMPT", {"code": code, "cond": cond, "price": price, "qty": qty, "order_amount": order_amount})

        ret = self._send_order("BUY", "0101", self.account, 1, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[BUY] ✅ 주문성공: {self._code_tag(code)} qty={qty}", key=f"BUY_OK_{code}", throttle=0.2)
            self.pending_codes.add(code)

            self.reserved_exposure[code] += int(order_amount)
            self._log("INFO", f"[BUY-LIMIT] reserved_exposure += {order_amount} -> now={self.reserved_exposure[code]} for {self._code_tag(code)}",
                      key=f"RESERVE_ADD_{code}", throttle=0.0)

            self.report.emit("BUY_OK", {"code": code, "cond": cond, "qty": qty, "price": price, "order_amount": order_amount,
                                        "reserved_exposure": int(self.reserved_exposure.get(code, 0) or 0)})

            if code not in self.bought_today:
                self.bought_today.add(code)
                self._log("INFO", f"[BOUGHT-TODAY] ✅ 주문성공 선등록(체잔 전): {self._code_tag(code)}",
                          key=f"BOUGHT_PRE_{code}", throttle=0.0)
                if BOUGHT_TODAY_PERSIST:
                    self._save_bought_today()
                self.report.emit("BOUGHT_TODAY_ADD", {"code": code, "via": "BUY_OK_PRE_CHEJAN"})
        else:
            self._log("WARN", f"[BUY] ❌ 주문실패: {self._code_tag(code)} ret={ret}", key=f"BUY_FAIL_{code}", throttle=1.0)
            self.report.emit("BUY_FAIL", {"code": code, "cond": cond, "ret": ret, "price": price, "qty": qty})

    # -----------------------
    # Delayed sell / risk / orphan
    # -----------------------
    def _schedule_delayed_sell(self, code: str, cond: str, reason: str):
        code = _norm_code(code)
        if not code:
            return
        due = time.time() + float(SELL_DELAY_SEC)
        self.delayed_sells[(code, cond)] = due
        self._log("INFO", f"[SELL-DELAY] 예약: {self._code_tag(code)} cond={cond} after={SELL_DELAY_SEC}s (reason={reason})",
                  key=f"SELL_SCHED_{code}_{cond}", throttle=0.2)
        self.report.emit("SELL_DELAY_SCHEDULE", {"code": code, "cond": cond, "due_in_sec": float(SELL_DELAY_SEC), "reason": reason})

    def _has_pending_sell(self, code: str) -> bool:
        code = _norm_code(code)
        if not code:
            return False
        for od in (self.unfilled_orders or {}).values():
            if _norm_code(od.get("code", "")) != code:
                continue
            bs = str(od.get("bs", "") or "")
            if ("매도" in bs) or bs.startswith("-"):
                if int(od.get("unfilled", 0) or 0) > 0:
                    return True
        return False

    def _request_sell_qty(self, code: str, qty: int, reason: str):
        code = _norm_code(code)
        hold_qty = int(self.holdings_qty.get(code, 0) or 0)
        qty = int(qty or 0)
        if hold_qty <= 0 or qty <= 0:
            return
        if qty > hold_qty:
            qty = hold_qty

        self._log("INFO", f"[SELL] 주문전송: {self._code_tag(code)} qty={qty} 시장가 reason={reason}",
                  key=f"SELL_SEND_{code}_{reason}", throttle=0.0)

        self.report.emit("SELL_ATTEMPT", {"code": code, "qty": qty, "reason": reason, "mode": "PARTIAL"})

        ret = self._send_order("SELL", "0102", self.account, 2, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[SELL] ✅ 주문성공: {self._code_tag(code)} qty={qty} reason={reason}",
                      key=f"SELL_OK_{code}_{reason}", throttle=0.0)
            self.pending_codes.add(code)
            self._set_rebuy_block(code, cooldown_sec=REBUY_COOLDOWN_SEC)
            self.report.emit("SELL_OK", {"code": code, "qty": qty, "reason": reason, "mode": "PARTIAL"})
        else:
            self._log("WARN", f"[SELL] ❌ 주문실패: {self._code_tag(code)} ret={ret} reason={reason}",
                      key=f"SELL_FAIL_{code}_{reason}", throttle=0.5)
            self.report.emit("SELL_FAIL", {"code": code, "qty": qty, "reason": reason, "ret": ret, "mode": "PARTIAL"})

    def _request_sell_all(self, code: str, reason: str):
        code = _norm_code(code)
        qty = int(self.holdings_qty.get(code, 0) or 0)
        if qty <= 0:
            return

        self._log("INFO", f"[SELL] 주문전송(전량): {self._code_tag(code)} qty={qty} 시장가 reason={reason}",
                  key=f"SELL_SEND_{code}", throttle=0.2)

        self.report.emit("SELL_ATTEMPT", {"code": code, "qty": qty, "reason": reason, "mode": "ALL"})

        ret = self._send_order("SELL", "0102", self.account, 2, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[SELL] ✅ 주문성공: {self._code_tag(code)} qty={qty} reason={reason}",
                      key=f"SELL_OK_{code}", throttle=0.2)
            self.pending_codes.add(code)
            self._set_rebuy_block(code, cooldown_sec=REBUY_COOLDOWN_SEC)
            self.report.emit("SELL_OK", {"code": code, "qty": qty, "reason": reason, "mode": "ALL"})
        else:
            self._log("WARN", f"[SELL] ❌ 주문실패: {self._code_tag(code)} ret={ret}",
                      key=f"SELL_FAIL_{code}", throttle=1.0)
            self.report.emit("SELL_FAIL", {"code": code, "qty": qty, "reason": reason, "ret": ret, "mode": "ALL"})

    def _set_rebuy_block(self, code: str, cooldown_sec: float):
        code = _norm_code(code)
        if not code:
            return
        self.rebuy_block_until[code] = time.time() + float(cooldown_sec)
        self.rebuy_block_pending.add(code)
        self.report.emit("REBUY_BLOCK_SET", {"code": code, "cooldown_sec": float(cooldown_sec)})

    def _is_in_any_sell_condition(self, code: str) -> bool:
        code = _norm_code(code)
        for s in self.sell_cond_members.values():
            if code in s:
                return True
        return False

    def _has_any_delayed_sell(self, code: str) -> bool:
        code = _norm_code(code)
        if not code:
            return False
        for (c, _cond) in self.delayed_sells.keys():
            if c == code:
                return True
        return False

    def _clear_delayed_sells_code(self, code: str):
        code = _norm_code(code)
        if not code:
            return
        removed = 0
        for k in list(self.delayed_sells.keys()):
            if k[0] == code:
                self.delayed_sells.pop(k, None)
                removed += 1
        if removed > 0:
            self._log("INFO", f"[SELL-DELAY] 정리: {self._code_tag(code)} 관련 예약 {removed}건 삭제(즉시매도 우선)",
                      key=f"SELL_DELAY_CLR_{code}", throttle=0.0)
            self.report.emit("SELL_DELAY_CLEAR_CODE", {"code": code, "removed": removed})

    def _process_delayed_sells_tick(self):
        if self._risk_running:
            return

        now = time.time()
        keys = list(self.delayed_sells.keys())
        for (code, cond) in keys:
            due = self.delayed_sells.get((code, cond), 0)
            if now < due:
                continue

            if cond in SELL_COND_NAMES and code in self.sell_cond_members.get(cond, set()):
                self.delayed_sells.pop((code, cond), None)
                self._log("INFO", f"[SELL-DELAY] 만기취소: {self._code_tag(code)} cond={cond} (re-entered)",
                          key=f"SELL_DELAY_CANCEL_{code}_{cond}", throttle=0.2)
                self.report.emit("SELL_DELAY_CANCEL_REENTER", {"code": code, "cond": cond})
                continue

            if code in self.bought_today:
                self.delayed_sells.pop((code, cond), None)
                self._log("INFO", f"[SELL-DELAY] 만기스킵(당일매수 조건매도 제외): {self._code_tag(code)} cond={cond}",
                          key=f"SELL_DELAY_DUE_SKIP_BOUGHT_{code}_{cond}", throttle=0.0)
                self.report.emit("SELL_DELAY_DUE_SKIP_BOUGHT_TODAY", {"code": code, "cond": cond})
                continue

            self.delayed_sells.pop((code, cond), None)

            hold_qty = int(self.holdings_qty.get(code, 0) or 0)
            if hold_qty <= 0:
                continue
            if code in self.pending_codes:
                continue

            self._request_sell_all(code, reason=f"EXIT_{cond}_DELAY{SELL_DELAY_SEC}s")

    def _orphan_sweeper_tick(self):
        if self._risk_running:
            return
        try:
            now = time.time()
            if not self.holdings_qty:
                return

            for code, qty in list(self.holdings_qty.items()):
                if qty <= 0:
                    self.orphan_first_seen.pop(code, None)
                    continue
                            # ✅ 추가: 당일 매수 종목은 ORPHAN 매도 대상에서 제외
                if code in self.bought_today:
                   self.orphan_first_seen.pop(code, None)
                   continue


                if code in self.pending_codes:
                    continue
                if self._has_any_delayed_sell(code):
                    self.orphan_first_seen.pop(code, None)
                    continue
                if self._is_in_any_sell_condition(code):
                    self.orphan_first_seen.pop(code, None)
                    continue

                first = self.orphan_first_seen.get(code)
                if first is None:
                    self.orphan_first_seen[code] = now
                    continue
                if (now - first) >= float(ORPHAN_GRACE_SEC):
                    self.report.emit("ORPHAN_FIRE", {"code": code, "qty": int(qty), "grace_sec": float(ORPHAN_GRACE_SEC)})
                    self._request_sell_all(code, reason="ORPHAN_SWEEPER")
                    self.orphan_first_seen[code] = now + 999999

        except Exception as e:
            self._log("ERROR", f"[ORPHAN] sweeper error: {e}", key="ORPHAN_ERR", throttle=5.0)
            self.report.emit("ORPHAN_ERR", {"err": repr(e)})

    def _risk_monitor_tick(self):
        if self._risk_running:
            return

        was_buy = self.buy_timer.isActive()
        was_delay = self.delay_sell_timer.isActive()
        was_orphan = self.orphan_timer.isActive()
        was_sync = self.sync_timer.isActive()

        try:
            self._risk_running = True

            if was_buy:
                self.buy_timer.stop()
            if was_delay:
                self.delay_sell_timer.stop()
            if was_orphan:
                self.orphan_timer.stop()
            if was_sync:
                self.sync_timer.stop()

            if not self.holdings_qty:
                return

            try:
                tiers = list(STOPLOSS_TIERS) if isinstance(STOPLOSS_TIERS, (list, tuple)) else []
                tiers = sorted(tiers, key=lambda x: float(x[0]), reverse=True)  # -3 > -5 > -7
            except Exception:
                tiers = [(-3.0, 0.30), (-5.0, 0.50), (-7.0, 1.00)]

            for code, qty in list(self.holdings_qty.items()):
                qty = int(qty or 0)
                if qty <= 0:
                    continue

                avg = int(self.holdings_avg.get(code, 0) or 0)
                if avg <= 0:
                    self._log("WARN", f"[RISK] 평단(avg) 없음/0 -> 스킵: {self._code_tag(code)} avg={avg}",
                              key=f"RISK_NOAVG_{code}", throttle=2.0)
                    self.report.emit("RISK_SKIP_NOAVG", {"code": code, "avg": avg})
                    continue

                if self._has_pending_sell(code):
                    continue

                cur = self.request_price(code)
                if not cur or cur <= 0:
                    self._log("DEBUG", f"[RISK] 현재가 미확보(큐잉/실패) -> 스킵: {self._code_tag(code)}",
                              key=f"RISK_NOPRICE_{code}", throttle=1.0)
                    self.report.emit("RISK_SKIP_NOPRICE", {"code": code})
                    continue

                peak = int(self.peak_price.get(code, 0) or 0)
                if TRAILING_ENABLED:
                    if peak <= 0 or cur > peak:
                        self.peak_price[code] = cur
                        peak = cur

                pnl_pct = (float(cur - avg) / float(avg)) * 100.0
                stage = int(self.stoploss_stage.get(code, 0) or 0)

                # 1) STOPLOSS tiers (한 tick에 한 단계만)
                if tiers and stage < len(tiers):
                    thr_pct, ratio = tiers[stage]
                    if pnl_pct <= float(thr_pct):
                        if code not in self.stoploss_base_qty:
                            self.stoploss_base_qty[code] = qty

                        cur_qty = qty
                        if stage == 0:
                            base_qty = int(self.stoploss_base_qty.get(code, cur_qty))
                            sell_qty = int(base_qty * float(ratio))
                        else:
                            sell_qty = int(cur_qty * float(ratio))

                        if float(ratio) >= 1.0:
                            sell_qty = cur_qty

                        sell_qty = max(1, min(cur_qty, sell_qty))

                        self._log(
                            "INFO",
                            f"[RISK] STOPLOSS 발동: {self._code_tag(code)} pnl={pnl_pct:.2f}% "
                            f"stage={stage+1}/{len(tiers)} thr={thr_pct}% ratio={ratio} sell_qty={sell_qty}",
                            key=f"STOPLOSS_FIRE_{code}_{stage}",
                            throttle=0.0,
                        )
                        self._clear_delayed_sells_code(code)
                        self.report.emit("STOPLOSS_FIRE", {
                            "code": code, "cur": cur, "avg": avg, "pnl_pct": pnl_pct,
                            "stage": stage, "thr_pct": float(thr_pct), "ratio": float(ratio),
                            "sell_qty": int(sell_qty), "cur_qty": int(cur_qty),
                        })

                        if sell_qty >= cur_qty:
                            self._request_sell_all(code, reason=f"STOPLOSS{stage+1}@{thr_pct}%")
                        else:
                            self._request_sell_qty(code, sell_qty, reason=f"STOPLOSS{stage+1}@{thr_pct}%")

                        self.stoploss_stage[code] = stage + 1
                        continue

                # 2) TRAILING (즉시 전량)
                if TRAILING_ENABLED and peak > 0:
                    drawdown_pct = (float(cur - peak) / float(peak)) * 100.0
                    if drawdown_pct <= -abs(float(TRAILING_STOP_PCT)):
                        self._log(
                            "INFO",
                            f"[RISK] TRAILING 발동: {self._code_tag(code)} cur={cur} peak={peak} dd={drawdown_pct:.2f}%",
                            key=f"TRAIL_FIRE_{code}",
                            throttle=0.0,
                        )
                        self._clear_delayed_sells_code(code)
                        self.report.emit("TRAILING_FIRE", {
                            "code": code, "cur": cur, "peak": peak,
                            "drawdown_pct": drawdown_pct, "trail_pct": float(TRAILING_STOP_PCT)
                        })
                        self._request_sell_all(code, reason=f"TRAILING_{TRAILING_STOP_PCT:.0f}")
                        continue

        except Exception as e:
            self._log("ERROR", f"[RISK] monitor error: {e}", key="RISK_ERR", throttle=5.0)
            self.report.emit("RISK_ERR", {"err": repr(e)})

        finally:
            self._risk_running = False
            if was_buy:
                self.buy_timer.start(250)
            if was_delay:
                self.delay_sell_timer.start(200)
            if was_orphan:
                self.orphan_timer.start(int(ORPHAN_CHECK_INTERVAL_SEC * 1000))
            if was_sync:
                self.sync_timer.start(int(max(5, float(SYNC_INTERVAL_SEC)) * 1000))

    def _periodic_sync_tick(self):
        if self._risk_running:
            return
        if self._tr_busy:
            return
        self.request_balance()
        self.request_unfilled()

    # -----------------------
    # Persistence
    # -----------------------
    def _load_sold_today(self):
        try:
            if not os.path.exists(SOLD_TODAY_FILE):
                return
            with open(SOLD_TODAY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            day = data.get("day", "")
            codes = set(map(_norm_code, data.get("codes", [])))
            if day == _today_yyyymmdd():
                self.sold_today = set(c for c in codes if c)
                self._log("INFO", f"[SOLD-TODAY] 로드: {len(self.sold_today)}개", key="SOLD_LOAD", throttle=0.2)
            else:
                self.sold_today = set()
        except Exception as e:
            self._log("WARN", f"[SOLD-TODAY] 로드 실패: {e}", key="SOLD_LOAD_FAIL", throttle=1.0)

    def _save_sold_today(self):
        try:
            payload = {"day": _today_yyyymmdd(), "codes": sorted(list(self.sold_today))}
            with open(SOLD_TODAY_FILE, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._log("WARN", f"[SOLD-TODAY] 저장 실패: {e}", key="SOLD_SAVE_FAIL", throttle=1.0)

    def _load_bought_today(self):
        try:
            if not os.path.exists(BOUGHT_TODAY_FILE):
                return
            with open(BOUGHT_TODAY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            day = data.get("day", "")
            codes = set(map(_norm_code, data.get("codes", [])))
            if day == _today_yyyymmdd():
                self.bought_today = set(c for c in codes if c)
                self._log("INFO", f"[BOUGHT-TODAY] 로드: {len(self.bought_today)}개", key="BOUGHT_LOAD", throttle=0.2)
            else:
                self.bought_today = set()
        except Exception as e:
            self._log("WARN", f"[BOUGHT-TODAY] 로드 실패: {e}", key="BOUGHT_LOAD_FAIL", throttle=1.0)

    def _save_bought_today(self):
        try:
            payload = {"day": _today_yyyymmdd(), "codes": sorted(list(self.bought_today))}
            with open(BOUGHT_TODAY_FILE, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._log("WARN", f"[BOUGHT-TODAY] 저장 실패: {e}", key="BOUGHT_SAVE_FAIL", throttle=1.0)

    # -----------------------
    # Start / UI
    # -----------------------
    def show_account_window(self):
        self._log("INFO", "[UI] KOA_Functions('ShowAccountWindow') 호출(계좌/비번 입력창 유도)", key="SHOW_ACC_WIN", throttle=1.0)
        try:
            ret = self.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
            self._log("INFO", f"[UI] ShowAccountWindow ret={ret}", key="SHOW_ACC_WIN_RET", throttle=1.0)
        except Exception as e:
            self._log("WARN", f"[UI] ShowAccountWindow 호출 실패: {repr(e)}", key="SHOW_ACC_WIN_FAIL", throttle=2.0)

    def start(self):
        self._log("INFO", "[INIT] 프로그램 초기화 시작")
        self._log("INFO", "[INIT] 프로그램 초기화 완료")

        self.comm_connect()
        if not self.account:
            self._log("ERROR", "[FATAL] 계좌번호 없음. 종료.")
            self.report.emit("FATAL_NO_ACCOUNT", {})
            return

        self.show_account_window()
        time.sleep(1.5)

        self.load_conditions()
        self.subscribe_conditions()

        self.request_balance()
        self.request_unfilled()

        self.buy_timer.start(250)
        self.delay_sell_timer.start(200)
        self.orphan_timer.start(int(ORPHAN_CHECK_INTERVAL_SEC * 1000))
        self.sync_timer.start(int(max(5, float(SYNC_INTERVAL_SEC)) * 1000))
        self.risk_timer.start(int(max(5, int(AUTO_SELL_INTERVAL_SEC)) * 1000))
        self.price_queue_timer.start(150)

        self._log("INFO", "[RUN] 타이머 시작 완료")
        self.report.emit("RUN_START", {})


def main():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    try:
        kiwoom.start()
        sys.exit(app.exec_())
    finally:
        try:
            kiwoom.report.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
