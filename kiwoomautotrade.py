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
- (A) 주문 거부/사유 MSG를 INFO로도 출력
- (B) 체잔(gubun=0)에서 주문상태(913)/거부사유(919) 로그 + 리포트
- (C) SendOrder ret=0은 "주문요청 성공"으로만 표기(체결 성공 착시 제거)
20260323 손익로그 추가
20260330 조건별 수익/리스크 분리
20260331 스탑로스 중복 제거

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
BUY_COND_NAMES = {"A", "x2", "w3"}
SELL_COND_NAMES = {"A","A_", "x2", "w3", "w"}

TARGET_BUY_AMOUNT = 50000
MAX_POSITION_PER_CODE = 50000
ALLOW_ADD_BUY = False

SELL_DELAY_SEC = 5.0
REBUY_COOLDOWN_SEC = 600.0

AUTO_SELL_INTERVAL_SEC = 60

TRAILING_STOP_PCT = 8.0
BALANCE_COOLDOWN_SEC = 3.0
SYNC_INTERVAL_SEC = 15.0
PRICE_REQ_INTERVAL = 0.25
PRICE_RETRY_MAX = 3
PRICE_RETRY_SLEEP = 0.8
PRICE_CACHE_HARD_TTL_SEC = 15.0
ORDER_INTENT_TIMEOUT_SEC = 12.0
CONDITION_CHATTER_WINDOW_SEC = 15.0
CONDITION_CHATTER_COUNT = 4
CONDITION_CHATTER_EXTRA_DELAY_SEC = 2.0
METRICS_EMIT_INTERVAL_SEC = 60.0

ORPHAN_CHECK_INTERVAL_SEC = 10
ORPHAN_GRACE_SEC = 60
ORPHAN_ENABLE_AFTER_HHMM = "0905"  # 장초반 오퍼런 스위퍼 지연(예: 09:05 이후부터만 검사 시작)

TODAY = datetime.datetime.now().strftime("%y%m%d")   # 예: 260312
MONTH = datetime.datetime.now().strftime("%y%m")     # 예: 2603

LOG_LEVEL = "INFO"  # "DEBUG"/"INFO"/"WARN"/"ERROR"
LOG_TO_FILE = True
LOG_DIR = "logs"
LOG_FILE_PATH = f"{LOG_DIR}/{MONTH}/kiwoomautotrade_{TODAY}.log"
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
# === Holdings snapshot (계좌 보유종목 JSON 저장) ===
HOLDINGS_SNAPSHOT_ENABLED = True
HOLDINGS_SNAPSHOT_FILE = "holdings_snapshot.json"

# === Take-profit (부분익절) ===
# +10% 도달 시 보유수량의 30% (기준수량=시작/신규편입 시점 수량) 익절
# +15% 도달 시 보유수량의 30% (기준수량의 30%) 추가 익절
TAKEPROFIT_LEVELS = [(8.0, 0.50), (13.0, 0.30)]  # (pnl_pct_threshold, sell_ratio_of_base_qty)
TAKEPROFIT_MIN_QTY = 1

# === Risk feature master switches ===
STOPLOSS_ENABLED = True
TAKEPROFIT_ENABLED = True
TRAILING_ENABLED = True

# === Condition-specific risk config ===
# - 상단 설정만 바꿔서 조건별 손절/익절/트레일링 on/off 및 수치 조절
# - 조건별 설정이 없으면 DEFAULT_RISK_CONFIG 사용
DEFAULT_RISK_CONFIG = {
    "stoploss_enabled": True,
    "takeprofit_enabled": True,
    "trailing_enabled": True,
    "stoploss_tiers": [(-2.0, 0.50), (-2.7, 0.50), (-3.2, 1.00)],
    "takeprofit_levels": [(8.0, 0.30), (13.0, 0.30)],
    "takeprofit_min_qty": 1,
    "trailing_stop_pct": 8.0,
}

COND_RISK_CONFIG = {
    "A": {
        "stoploss_enabled": True,
        "takeprofit_enabled": True,
        "trailing_enabled": True,
        "stoploss_tiers": [(-2.0, 0.50), (-2.7, 0.50), (-3.2, 1.00)],
        "takeprofit_levels": [(8.0, 0.30), (13.0, 0.30)],
        "takeprofit_min_qty": 1,
        "trailing_stop_pct": 8.0,
    },
    "x2": {
        "stoploss_enabled": True,
        "takeprofit_enabled": False,
        "trailing_enabled": True,
        "stoploss_tiers": [(-2.0, 0.50), (-2.7, 0.50), (-3.2, 1.00)],
        "takeprofit_levels": [(8.0, 0.30), (13.0, 0.30)],
        "takeprofit_min_qty": 1,
        "trailing_stop_pct": 8.0,
    },
    "w3": {
        "stoploss_enabled": True,
        "takeprofit_enabled": True,
        "trailing_enabled": True,
        "stoploss_tiers": [(-2.0, 0.50), (-2.7, 0.50), (-3.2, 1.00)],
        "takeprofit_levels": [(8.0, 0.60), (15.0, 0.20)],
        "takeprofit_min_qty": 1,
        "trailing_stop_pct": 8.0,
    },
}

# === Condition-specific scheduled force-sell config ===
# - 상단 설정만 바꿔서 조건별 시간 강제청산 가능
# - action: SELL_ALL | SELL_PARTIAL
# - weekdays: ["MON", "TUE", "WED", "THU", "FRI"]
# - time은 hour/minute 또는 hhmm("1400")로 지정 가능
# - once_per_day=True 이면 같은 규칙이 하루에 한 번만 발동
CONDITION_FORCE_SELL_ENABLED = True
COND_FORCE_SELL_CONFIG = {
    "A": [
        # {"weekdays": ["FRI"], "hour": 14, "minute": 0, "action": "SELL_ALL", "reason": "A_FRI_1400_FORCE_EXIT", "once_per_day": True},
    ],
    "x2": [{"weekdays": ["FRI"], "hour": 14, "minute": 0, "action": "SELL_ALL", "reason": "A_FRI_1400_FORCE_EXIT", "once_per_day": True}
    ],
    "w3": [
    ],
}


# === 거래 시간 게이트 (KST) ===
# 프로그램을 언제 실행하든 상관없이 아래 시간대에만 자동매매(조건/주문/리스크/오펀)가 동작합니다.
TRADE_WINDOW_ENABLED = True
TRADE_START_HHMM = "0850"
TRADE_END_HHMM   = "2000"  # end is exclusive

# 창 밖에서는 완전 종료할지(앱 종료), 아니면 대기만 할지
EXIT_AFTER_WINDOW = False



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

        # 최소침습 상태머신/롤백 보강
        self.order_state: Dict[str, str] = defaultdict(str)   # NONE/BUY_SENT/BUY_WORKING/BUY_FILLED/SELL_SENT/SELL_WORKING/SELL_FILLED/REJECTED/CANCELED
        self.order_state_ts: Dict[str, float] = {}
        self.order_meta: Dict[str, Dict[str, Any]] = {}       # code -> last sent order meta
        self.order_no_to_code: Dict[str, str] = {}

        # 운영/분석용 메트릭
        self.metrics: Dict[str, int] = defaultdict(int)
        self.metric_reasons: Dict[str, int] = defaultdict(int)
        self._last_metrics_emit_ts = 0.0

        # 조건 진동 완화용
        self.cond_toggle_hist: Dict[Tuple[str, str], deque] = defaultdict(deque)

        # ✅ 로컬 방어막: 잔고 누락 시에도 "추정 노출"로 한도 강제
        self.reserved_exposure: Dict[str, int] = defaultdict(int)

        # buy queue
        self.buy_queue = deque()
        self.buy_queue_set: Set[str] = set()
        # 마지막 미매수(=매수 시도 실패/차단) 사유 기록
        # - 조건 편입(I) 로그와 함께 '왜 미매수였는지' 추적하기 위함
        self.last_buy_block_reason: Dict[str, str] = {}  # code -> reason
        self.last_buy_block_ts: Dict[str, float] = {}    # code -> time.time()

        # delayed sells
        self.delayed_sells: Dict[Tuple[str, str], float] = {}

        # sold/bought today
        self.sold_today: Set[str] = set()
        self.bought_today: Set[str] = set()

        self.rebuy_block_until: Dict[str, float] = {}
        self.rebuy_block_pending: Set[str] = set()

        # 매도 컨텍스트(ORPHAN/조건이탈 포함) 추적용
        self.pending_sell_meta: Dict[str, Dict[str, Any]] = {}

        # 종목별 최초/주요 진입 조건 저장 (조건별 리스크 설정 참조용)
        self.buy_condition_by_code: Dict[str, str] = {}

        # Report enhanced: 조건 편입시각 / trade_id 추적
        self.condition_enter_ts: Dict[Tuple[str, str], float] = {}
        self.active_trade_id_by_code: Dict[str, str] = {}

        # 조건별 시간 강제청산 중복 방지
        self.force_sell_fired_today: Set[str] = set()

        # orphan
        self.orphan_first_seen: Dict[str, float] = {}

        # risk
        self.peak_price: Dict[str, int] = {}
        self.stoploss_stage: Dict[str, int] = defaultdict(int)
        self.stoploss_base_qty: Dict[str, int] = defaultdict(int)
        # take-profit
        self.tp_stage: Dict[str, int] = defaultdict(int)      # code -> stage index (0=none)
        self.tp_base_qty: Dict[str, int] = defaultdict(int)   # code -> base qty for partial take-profit

        # holdings snapshot
        self._holdings_snapshot_sig = None
        self._last_holdings_snapshot_ts = 0.0


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
        self._log("INFO", f"[BOOT-CONFIG] AUTO_SELL_INTERVAL_SEC={AUTO_SELL_INTERVAL_SEC}")
        self._log("INFO", f"[BOOT-CONFIG] STOPLOSS_ENABLED={STOPLOSS_ENABLED} default_tiers={DEFAULT_RISK_CONFIG.get('stoploss_tiers', [])}")
        self._log("INFO", f"[BOOT-CONFIG] TAKEPROFIT_ENABLED={TAKEPROFIT_ENABLED} levels={TAKEPROFIT_LEVELS} min_qty={TAKEPROFIT_MIN_QTY}")
        self._log("INFO", f"[BOOT-CONFIG] TRAILING_ENABLED={TRAILING_ENABLED} / TRAILING_STOP_PCT={TRAILING_STOP_PCT}")
        self._log("INFO", f"[BOOT-CONFIG] DEFAULT_RISK_CONFIG={DEFAULT_RISK_CONFIG}")
        self._log("INFO", f"[BOOT-CONFIG] COND_RISK_CONFIG={COND_RISK_CONFIG}")
        self._log("INFO", f"[BOOT-CONFIG] CONDITION_FORCE_SELL_ENABLED={CONDITION_FORCE_SELL_ENABLED} / COND_FORCE_SELL_CONFIG={COND_FORCE_SELL_CONFIG}")
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
            "STOPLOSS_ENABLED": STOPLOSS_ENABLED,
            "TAKEPROFIT_ENABLED": TAKEPROFIT_ENABLED,
            "TAKEPROFIT_LEVELS": TAKEPROFIT_LEVELS,
            "TAKEPROFIT_MIN_QTY": TAKEPROFIT_MIN_QTY,
            "AUTO_SELL_INTERVAL_SEC": AUTO_SELL_INTERVAL_SEC,
            "TRAILING_ENABLED": TRAILING_ENABLED,
            "TRAILING_STOP_PCT": TRAILING_STOP_PCT,
            "DEFAULT_RISK_CONFIG": DEFAULT_RISK_CONFIG,
            "COND_RISK_CONFIG": COND_RISK_CONFIG,
            "CONDITION_FORCE_SELL_ENABLED": CONDITION_FORCE_SELL_ENABLED,
            "COND_FORCE_SELL_CONFIG": COND_FORCE_SELL_CONFIG,
        })

        if SOLD_TODAY_PERSIST:
            self._load_sold_today()
        if BOUGHT_TODAY_PERSIST:
            self._load_bought_today()
        if HOLDINGS_SNAPSHOT_ENABLED:
            self._load_holdings_snapshot()

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

        self.force_sell_timer = QTimer()
        self.force_sell_timer.timeout.connect(self._force_sell_schedule_tick)

        # trade window guard
        self._trade_window_last: Optional[bool] = None
        self.trade_window_timer = QTimer()
        self.trade_window_timer.timeout.connect(self._trade_window_guard_tick)
        self.trade_window_timer.start(1000)  # 1초마다 거래창 상태 감시


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

    def _metric_inc(self, key: str, amount: int = 1, reason: str = ""):
        try:
            key = str(key or "").strip()
            if not key:
                return
            self.metrics[key] += int(amount)
            if reason:
                self.metric_reasons[f"{key}:{reason}"] += int(amount)
        except Exception:
            pass

    def _emit_metrics_snapshot(self, reason: str = "", force: bool = False):
        now = time.time()
        if (not force) and (now - float(self._last_metrics_emit_ts or 0.0) < float(METRICS_EMIT_INTERVAL_SEC)):
            return
        self._last_metrics_emit_ts = now
        try:
            counters = {k: int(v) for k, v in sorted(self.metrics.items()) if int(v) > 0}
            reasons = {k: int(v) for k, v in sorted(self.metric_reasons.items()) if int(v) > 0}
            state_counts = defaultdict(int)
            for _code, st in list(self.order_state.items()):
                st = str(st or "")
                if st:
                    state_counts[st] += 1
            payload = {
                "reason": reason,
                "counters": counters,
                "reasons": reasons,
                "order_states": {k: int(v) for k, v in sorted(state_counts.items()) if int(v) > 0},
                "pending_codes": sorted(list(self.pending_codes)),
            }
            self.report.emit("METRICS_SNAPSHOT", payload)
            self._log("INFO", f"[METRICS] reason={reason} counters={counters} order_states={payload['order_states']}", key="METRICS_SNAPSHOT", throttle=5.0)
        except Exception as e:
            self._log("WARN", f"[METRICS] snapshot emit failed: {e}", key="METRICS_SNAPSHOT_FAIL", throttle=5.0)

    def _set_order_state(self, code: str, state: str, order_no: str = "", **extra):
        code = _norm_code(code)
        if not code:
            return
        state = str(state or "")
        self.order_state[code] = state
        self.order_state_ts[code] = time.time()
        meta = dict(self.order_meta.get(code, {}) or {})
        if state:
            meta["state"] = state
        if order_no:
            meta["order_no"] = str(order_no)
            self.order_no_to_code[str(order_no)] = code
        for k, v in extra.items():
            meta[k] = v
        if meta:
            self.order_meta[code] = meta
        self.report.emit("ORDER_STATE", {"code": code, "state": state, "order_no": str(order_no or meta.get('order_no', '')), **extra})

    def _clear_order_state(self, code: str, keep_meta: bool = False):
        code = _norm_code(code)
        if not code:
            return
        self.order_state.pop(code, None)
        self.order_state_ts.pop(code, None)
        if not keep_meta:
            meta = self.order_meta.pop(code, None)
            try:
                order_no = str((meta or {}).get("order_no", "")).strip()
                if order_no:
                    self.order_no_to_code.pop(order_no, None)
            except Exception:
                pass

    def _live_unfilled_exists(self, code: str) -> bool:
        code = _norm_code(code)
        if not code:
            return False
        for od in (self.unfilled_orders or {}).values():
            if _norm_code(od.get("code", "")) == code and int(od.get("unfilled", 0) or 0) > 0:
                return True
        return False

    def _register_sent_order(self, code: str, side: str, qty: int, price: int, reason: str = "", cond: str = "", reserved_amount: int = 0, pre_bought_added: bool = False):
        code = _norm_code(code)
        if not code:
            return
        side = str(side or "").upper()
        self._set_order_state(
            code,
            f"{side}_SENT",
            side=side,
            qty=int(qty or 0),
            price=int(price or 0),
            reason=str(reason or ""),
            cond=str(cond or ""),
            sent_ts=time.time(),
            reserved_amount=int(reserved_amount or 0),
            pre_bought_added=bool(pre_bought_added),
            rollback_done=False,
        )

    def _rollback_sent_order(self, code: str, source: str, detail: str = ""):
        code = _norm_code(code)
        if not code:
            return False
        meta = dict(self.order_meta.get(code, {}) or {})
        if not meta:
            return False
        if meta.get("rollback_done"):
            return False
        if self._live_unfilled_exists(code):
            return False

        side = str(meta.get("side", "")).upper()
        order_no = str(meta.get("order_no", "")).strip()
        reserved_amount = int(meta.get("reserved_amount", 0) or 0)
        pre_bought_added = bool(meta.get("pre_bought_added", False))
        hold_qty = int(self.holdings_qty.get(code, 0) or 0)

        if side == "BUY":
            if hold_qty > 0:
                return False
            if reserved_amount > 0:
                self.reserved_exposure[code] = max(0, int(self.reserved_exposure.get(code, 0) or 0) - reserved_amount)
            if pre_bought_added and code in self.bought_today:
                self.bought_today.discard(code)
                if BOUGHT_TODAY_PERSIST:
                    self._save_bought_today()
            self.pending_codes.discard(code)
        elif side == "SELL":
            if hold_qty <= 0:
                return False
            self.pending_codes.discard(code)
            self.rebuy_block_pending.discard(code)
            self.rebuy_block_until.pop(code, None)
        else:
            self.pending_codes.discard(code)

        meta["rollback_done"] = True
        meta["rollback_source"] = str(source or "")
        meta["rollback_detail"] = str(detail or "")
        self.order_meta[code] = meta
        self._set_order_state(code, "CANCELED", order_no=order_no, side=side, source=str(source or ""), detail=str(detail or ""))
        self._metric_inc("order_rollback", reason=f"{side}:{source}")
        self._log("WARN", f"[ORDER-ROLLBACK] code={self._code_tag(code)} side={side} source={source} detail={detail}", key=f"ORDER_ROLLBACK_{code}_{side}_{source}", throttle=0.0)
        self.report.emit("ORDER_ROLLBACK", {"code": code, "side": side, "order_no": order_no, "source": str(source or ""), "detail": str(detail or "")})
        return True

    def _cleanup_stale_order_intents(self):
        now = time.time()
        for code, meta in list(self.order_meta.items()):
            try:
                state = str(meta.get("state", "") or "")
                if state not in {"BUY_SENT", "SELL_SENT", "BUY_WORKING", "SELL_WORKING"}:
                    continue
                sent_ts = float(meta.get("sent_ts", self.order_state_ts.get(code, 0.0)) or 0.0)
                if sent_ts <= 0:
                    continue
                if (now - sent_ts) < float(ORDER_INTENT_TIMEOUT_SEC):
                    continue
                if self._live_unfilled_exists(code):
                    continue
                side = str(meta.get("side", "") or "").upper()
                hold_qty = int(self.holdings_qty.get(code, 0) or 0)
                if side == "BUY" and hold_qty > 0:
                    continue
                if side == "SELL" and hold_qty <= 0:
                    continue
                self._rollback_sent_order(code, source="TIMEOUT", detail=f"state={state}")
            except Exception as e:
                self._log("WARN", f"[ORDER-TIMEOUT] cleanup failed code={code}: {e}", key=f"ORDER_TIMEOUT_ERR_{code}", throttle=5.0)

    def _recent_sent_candidates(self, side: str, within_sec: float = 3.0):
        side = str(side or "").upper()
        now = time.time()
        out = []
        for code, meta in list(self.order_meta.items()):
            try:
                if str(meta.get("side", "")).upper() != side:
                    continue
                st = str(meta.get("state", ""))
                if st not in {f"{side}_SENT", f"{side}_WORKING"}:
                    continue
                sent_ts = float(meta.get("sent_ts", self.order_state_ts.get(code, 0.0)) or 0.0)
                if sent_ts > 0 and (now - sent_ts) <= float(within_sec):
                    out.append(code)
            except Exception:
                continue
        return out

    def _get_cached_price(self, code: str, max_age_sec: float = 2.0) -> Optional[int]:
        code = _norm_code(code)
        cached = self._price_cache.get(code)
        if not cached:
            return None
        price, ts = cached
        if int(price or 0) <= 0:
            return None
        if (time.time() - float(ts)) <= float(max_age_sec):
            return int(price)
        return None

    def _is_zero_like_text(self, s: str) -> bool:
        s = _strip_html(s).strip()
        if not s:
            return True
        compact = re.sub(r"\s+", "", s)
        return bool(re.fullmatch(r"0+", compact))

    def _is_chejan_reject(self, status: str, reject: str) -> bool:
        status = _strip_html(status).strip()
        reject = _strip_html(reject).strip()
        reject_keywords = ["거부", "실패", "취소거부", "정정거부"]

        if reject and (not self._is_zero_like_text(reject)):
            return True
        if any(k in status for k in reject_keywords):
            return True
        return False

    def _track_condition_toggle(self, code: str, cond_name: str, event_type: str) -> float:
        code = _norm_code(code)
        cond_name = str(cond_name or "")
        event_type = str(event_type or "")
        key = (code, cond_name)
        dq = self.cond_toggle_hist[key]
        now = time.time()
        dq.append((now, event_type))
        while dq and (now - float(dq[0][0])) > float(CONDITION_CHATTER_WINDOW_SEC):
            dq.popleft()
        if len(dq) >= int(CONDITION_CHATTER_COUNT):
            self._metric_inc("cond_chatter_detected", reason=cond_name)
            self._log("WARN", f"[COND-CHATTER] 감지: {self._code_tag(code)} cond={cond_name} toggles={len(dq)} window={CONDITION_CHATTER_WINDOW_SEC}s", key=f"COND_CHATTER_{code}_{cond_name}", throttle=2.0)
            self.report.emit("COND_CHATTER", {"code": code, "cond": cond_name, "toggle_count": len(dq), "window_sec": float(CONDITION_CHATTER_WINDOW_SEC)})
            return float(CONDITION_CHATTER_EXTRA_DELAY_SEC)
        return 0.0

    def _get_primary_condition_for_code(self, code: str) -> str:
        code = _norm_code(code)
        if not code:
            return ""
        cond = str(self.buy_condition_by_code.get(code, "") or "").strip()
        if cond:
            return cond
        meta = dict(self.order_meta.get(code, {}) or {})
        cond = str(meta.get("cond", "") or "").strip()
        return cond

    def _get_risk_config_for_code(self, code: str) -> Dict[str, Any]:
        code = _norm_code(code)
        cond = self._get_primary_condition_for_code(code)
        cfg = dict(DEFAULT_RISK_CONFIG)
        cond_cfg = dict(COND_RISK_CONFIG.get(cond, {}) or {})
        cfg.update(cond_cfg)

        cfg["cond"] = cond
        cfg["stoploss_enabled"] = bool(STOPLOSS_ENABLED and bool(cfg.get("stoploss_enabled", True)))
        cfg["takeprofit_enabled"] = bool(TAKEPROFIT_ENABLED and bool(cfg.get("takeprofit_enabled", True)))
        cfg["trailing_enabled"] = bool(TRAILING_ENABLED and bool(cfg.get("trailing_enabled", True)))
        cfg["stoploss_tiers"] = list(cfg.get("stoploss_tiers", DEFAULT_RISK_CONFIG.get("stoploss_tiers", [])) or [])
        cfg["takeprofit_levels"] = list(cfg.get("takeprofit_levels", TAKEPROFIT_LEVELS) or [])
        cfg["takeprofit_min_qty"] = int(cfg.get("takeprofit_min_qty", TAKEPROFIT_MIN_QTY) or 1)
        cfg["trailing_stop_pct"] = float(cfg.get("trailing_stop_pct", TRAILING_STOP_PCT) or 0.0)
        return cfg

    def _weekday_key(self, now: datetime.datetime = None) -> str:
        if now is None:
            now = datetime.datetime.now()
        return ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"][now.weekday()]

    def _normalize_force_sell_rule(self, cond: str, rule: Dict[str, Any]) -> Dict[str, Any]:
        rule = dict(rule or {})
        weekdays = [str(x).upper() for x in list(rule.get("weekdays", []) or [])]
        hhmm = str(rule.get("hhmm", "") or "").strip()
        hour = rule.get("hour", None)
        minute = rule.get("minute", None)
        if hhmm and len(hhmm) == 4 and hhmm.isdigit():
            hour = int(hhmm[:2])
            minute = int(hhmm[2:])
        hour = int(hour if hour is not None else 0)
        minute = int(minute if minute is not None else 0)
        action = str(rule.get("action", "SELL_ALL") or "SELL_ALL").upper()
        ratio = float(rule.get("ratio", 1.0) or 1.0)
        once_per_day = bool(rule.get("once_per_day", True))
        only_if_profit = bool(rule.get("only_if_profit", False))
        only_if_loss = bool(rule.get("only_if_loss", False))
        reason = str(rule.get("reason", "") or "").strip() or f"{cond}_{'_'.join(weekdays) if weekdays else 'ANY'}_{hour:02d}{minute:02d}_{action}"
        rule_id = str(rule.get("rule_id", "") or "").strip() or reason
        return {
            "cond": str(cond or ""),
            "weekdays": weekdays,
            "hour": max(0, min(23, hour)),
            "minute": max(0, min(59, minute)),
            "action": action,
            "ratio": ratio,
            "once_per_day": once_per_day,
            "only_if_profit": only_if_profit,
            "only_if_loss": only_if_loss,
            "reason": reason,
            "rule_id": rule_id,
        }

    def _get_force_sell_rules_for_code(self, code: str):
        code = _norm_code(code)
        cond = self._get_primary_condition_for_code(code)
        rules = []
        for raw in list(COND_FORCE_SELL_CONFIG.get(cond, []) or []):
            try:
                rules.append(self._normalize_force_sell_rule(cond, raw))
            except Exception as e:
                self._log("WARN", f"[FORCE-SELL] invalid rule cond={cond} raw={raw} err={e}", key=f"FORCE_RULE_BAD_{cond}", throttle=5.0)
        return cond, rules

    def _make_force_sell_fire_key(self, code: str, rule: Dict[str, Any], now: datetime.datetime = None) -> str:
        if now is None:
            now = datetime.datetime.now()
        return f"{now.strftime('%Y%m%d')}|{_norm_code(code)}|{str(rule.get('rule_id', ''))}"

    def _force_sell_rule_is_due(self, rule: Dict[str, Any], now: datetime.datetime = None) -> bool:
        if now is None:
            now = datetime.datetime.now()
        weekdays = list(rule.get("weekdays", []) or [])
        if weekdays and self._weekday_key(now) not in weekdays:
            return False
        rule_hhmm = int(rule.get("hour", 0)) * 100 + int(rule.get("minute", 0))
        now_hhmm = now.hour * 100 + now.minute
        return now_hhmm >= rule_hhmm

    def _force_sell_schedule_tick(self):
        if not self._in_trade_window():
            return
        if not CONDITION_FORCE_SELL_ENABLED:
            return
        if self._risk_running:
            return
        if not self.holdings_qty:
            return

        now = datetime.datetime.now()
        today_prefix = now.strftime('%Y%m%d') + "|"
        stale = [x for x in list(self.force_sell_fired_today) if not str(x).startswith(today_prefix)]
        for k in stale:
            self.force_sell_fired_today.discard(k)

        for code, qty in list(self.holdings_qty.items()):
            code = _norm_code(code)
            qty = int(qty or 0)
            if not code or qty <= 0:
                continue
            if code in self.pending_codes or self._has_pending_sell(code):
                continue

            cond, rules = self._get_force_sell_rules_for_code(code)
            if not cond or not rules:
                continue

            snap = None
            for rule in rules:
                if not self._force_sell_rule_is_due(rule, now=now):
                    continue

                fire_key = self._make_force_sell_fire_key(code, rule, now=now)
                if bool(rule.get("once_per_day", True)) and fire_key in self.force_sell_fired_today:
                    continue

                if snap is None:
                    snap = self._build_sell_context_snapshot(code)
                est_pnl_amt = int(snap.get("est_pnl_amt", 0) or 0)
                if bool(rule.get("only_if_profit", False)) and est_pnl_amt <= 0:
                    continue
                if bool(rule.get("only_if_loss", False)) and est_pnl_amt >= 0:
                    continue

                reason = str(rule.get("reason", "FORCE_SELL") or "FORCE_SELL")
                action = str(rule.get("action", "SELL_ALL") or "SELL_ALL").upper()
                extra = {
                    "source": "CONDITION_FORCE_SELL",
                    "cond": cond,
                    "rule_id": str(rule.get("rule_id", "") or ""),
                    "weekday": self._weekday_key(now),
                    "scheduled_hhmm": f"{int(rule.get('hour', 0)):02d}{int(rule.get('minute', 0)):02d}",
                    "only_if_profit": bool(rule.get("only_if_profit", False)),
                    "only_if_loss": bool(rule.get("only_if_loss", False)),
                }

                fired = False
                if action == "SELL_PARTIAL":
                    ratio = float(rule.get("ratio", 1.0) or 1.0)
                    sell_qty = max(1, min(qty, int(qty * ratio)))
                    self._request_sell_qty(code, sell_qty, reason=reason, price_hint=int(snap.get("cur", 0) or 0), extra=extra)
                    fired = bool(code in self.pending_codes or self._has_pending_sell(code))
                else:
                    self._request_sell_all(code, reason=reason, price_hint=int(snap.get("cur", 0) or 0), extra=extra)
                    fired = bool(code in self.pending_codes or self._has_pending_sell(code))

                if fired:
                    self.force_sell_fired_today.add(fire_key)
                    self._log("INFO", f"[FORCE-SELL] fired: {self._code_tag(code)} cond={cond} action={action} reason={reason}", key=f"FORCE_SELL_FIRE_{code}_{reason}", throttle=0.0)
                    self.report.emit("FORCE_SELL_FIRED", {"code": code, "cond": cond, "action": action, "reason": reason, "rule": rule, "snapshot": snap})
                else:
                    self._log("WARN", f"[FORCE-SELL] trigger but order not sent: {self._code_tag(code)} cond={cond} action={action} reason={reason}", key=f"FORCE_SELL_NOFIRE_{code}_{reason}", throttle=2.0)
                    self.report.emit("FORCE_SELL_NOT_FIRED", {"code": code, "cond": cond, "action": action, "reason": reason, "rule": rule, "snapshot": snap})
                break

    # -----------------------
    # Trade window gate (KST)
    # -----------------------
    def _in_trade_window(self, now: datetime.datetime = None) -> bool:
        if not TRADE_WINDOW_ENABLED:
            return True
        if now is None:
            now = datetime.datetime.now()
        hhmm = now.strftime('%H%M')
        return (TRADE_START_HHMM <= hhmm < TRADE_END_HHMM)

    def _start_trade_session(self):
        """거래창 진입 시: 조건구독/동기화/타이머 시작"""
        try:
            self.subscribe_conditions()
        except Exception as e:
            self._log('WARN', f'[TRADE-WINDOW] subscribe_conditions error: {e}', key='TW_SUB_ERR', throttle=2.0)

        try:
            self.request_balance()
            self.request_unfilled()
        except Exception as e:
            self._log('WARN', f'[TRADE-WINDOW] initial sync error: {e}', key='TW_SYNC_ERR', throttle=2.0)

        self.buy_timer.start(250)
        self.delay_sell_timer.start(200)
        self.orphan_timer.start(int(ORPHAN_CHECK_INTERVAL_SEC * 1000))
        self.sync_timer.start(int(max(5, float(SYNC_INTERVAL_SEC)) * 1000))
        self.risk_timer.start(int(max(5, int(AUTO_SELL_INTERVAL_SEC)) * 1000))
        self.force_sell_timer.start(5000)
        self.price_queue_timer.start(150)

    def _stop_all_conditions(self):
        """조건식 실시간 구독 중지"""
        for name in list(self.subscribed_conds):
            try:
                idx = int(self.cond_name_to_idx.get(name, -1))
                if idx < 0:
                    continue
                scr = str(SCREEN_COND_BASE + idx)
                self.dynamicCall('SendConditionStop(QString, QString, int)', scr, name, idx)
            except Exception:
                pass
        self.subscribed_conds.clear()

    def _stop_trade_session(self):
        """거래창 이탈 시: 주문/리스크 관련 타이머 정지 + 큐/예약 정리 + 조건중지"""
        try:
            self.buy_timer.stop()
            self.delay_sell_timer.stop()
            self.orphan_timer.stop()
            self.sync_timer.stop()
            self.risk_timer.stop()
            self.force_sell_timer.stop()
            self.price_queue_timer.stop()
        except Exception:
            pass

        try:
            self.buy_queue.clear()
            self.buy_queue_set.clear()
            self.delayed_sells.clear()
        except Exception:
            pass

        self._cleanup_stale_order_intents()
        self._stop_all_conditions()

    def _trade_window_guard_tick(self):
        inwin = self._in_trade_window()
        if self._trade_window_last is None:
            self._trade_window_last = inwin
            return

        if inwin == self._trade_window_last:
            return

        self._trade_window_last = inwin
        if inwin:
            self._log('INFO', f'[TRADE-WINDOW] ✅ 거래창 진입: {TRADE_START_HHMM}-{TRADE_END_HHMM} -> 자동매매 ON', key='TW_ON', throttle=0.0)
            self.report.emit('TRADE_WINDOW_ON', {'start': TRADE_START_HHMM, 'end': TRADE_END_HHMM})
            self._start_trade_session()
        else:
            self._log('INFO', f'[TRADE-WINDOW] ⛔ 거래창 종료: {TRADE_START_HHMM}-{TRADE_END_HHMM} -> 자동매매 OFF', key='TW_OFF', throttle=0.0)
            self.report.emit('TRADE_WINDOW_OFF', {'start': TRADE_START_HHMM, 'end': TRADE_END_HHMM})
            self._stop_trade_session()
            if EXIT_AFTER_WINDOW:
                self._log('INFO', '[TRADE-WINDOW] EXIT_AFTER_WINDOW=True 이므로 프로그램 종료', key='TW_EXIT', throttle=0.0)
                QApplication.instance().quit()

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
                self._metric_inc("price_tr_fail", reason=code)
                self.report.emit("PRICE_FAIL", {"code": code, "attempt": attempt, "via": "TR_NOW"})
            time.sleep(PRICE_RETRY_SLEEP)

        return None

    def _process_price_queue_tick(self):
        if not self._in_trade_window():
            return
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
                self._metric_inc("tr_skip_busy", reason=rq_name)
                self.report.emit("TR_SKIP_BUSY", {"rq": rq_name, "busy_rq": self._tr_busy_name})
                return False

        self._tr_busy = True
        self._tr_busy_name = rq_name
        self._tr_deadline = time.time() + float(timeout_sec)

        ret = self.dynamicCall("CommRqData(QString, QString, int, QString)", rq_name, tr_code, int(prev_next), screen)
        if ret != 0:
            self._log("WARN", f"[TR-SERIAL] CommRqData fail ret={ret} rq={rq_name}", key=f"TR_FAIL_{rq_name}", throttle=0.5)
            self._metric_inc("tr_fail", reason=rq_name)
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
            self._metric_inc("tr_timeout", reason=self._tr_busy_name)
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
                cc = _norm_code(c)
                self.condition_enter_ts.setdefault((cc, cond_name), time.time())
                queued, reason = self._enqueue_buy(cc, cond_name, ev="INITIAL_TRCOND")
                if not queued:
                    cc = _norm_code(c)
                    self._log("INFO", f"[COND-INIT] 미매수(큐 등록 실패): {self._code_tag(cc)} cond={cond_name} reason={reason}",
                              key=f"COND_INIT_NOBUY_{cond_name}_{cc}", throttle=0.5)
                    self.report.emit("COND_INIT_NOBUY", {"cond": cond_name, "code": cc, "reason": reason})
    def _on_receive_real_condition(self, code, event_type, cond_name, cond_index):
        code = _norm_code(code)
        cond_name = str(cond_name)
        event_type = str(event_type)
        self._log("INFO", f"[COND-REAL-ENTRY] code={code} type={event_type} cond={cond_name} idx={cond_index}", key=f"COND_ENTRY_{cond_name}_{code}_{event_type}", throttle=0.2)
        if not code:
            return

        extra_delay = self._track_condition_toggle(code, cond_name, event_type)

        if event_type == "I":
            self.condition_enter_ts[(code, cond_name)] = time.time()
            self._log("INFO", f"[COND-REAL] cond={cond_name} 편입(I): {self._code_tag(code)}", key=f"COND_I_{cond_name}_{code}", throttle=0.5)
            self.report.emit("COND_REAL_I", {"cond": cond_name, "code": code, "idx": int(cond_index), "condition_enter_ts": _now_iso()})
            if cond_name in SELL_COND_NAMES:
                self.sell_cond_members[cond_name].add(code)
            if cond_name in BUY_COND_NAMES:
                queued, reason = self._enqueue_buy(code, cond_name, ev="REAL_I")
                if not queued:
                    self._log("INFO", f"[COND-REAL] 미매수(큐 등록 실패): {self._code_tag(_norm_code(code))} reason={reason}", key=f"COND_I_NOBUY_{cond_name}_{code}", throttle=0.5)
                    self.report.emit("COND_REAL_I_NOBUY", {"cond": cond_name, "code": _norm_code(code), "reason": reason})

        elif event_type == "D":
            duration_sec = self._get_condition_duration_sec(code, cond_name)
            self._log("INFO", f"[COND-REAL] cond={cond_name} 이탈(D): {self._code_tag(code)}", key=f"COND_D_{cond_name}_{code}", throttle=0.5)
            self.report.emit("COND_REAL_D", {"cond": cond_name, "code": code, "idx": int(cond_index), "condition_duration_sec": float(duration_sec)})
            self.condition_enter_ts.pop((code, cond_name), None)
            if cond_name in SELL_COND_NAMES:
                self.sell_cond_members[cond_name].discard(code)

                if code in self.bought_today:
                    self._log("INFO", f"[SELL-DELAY] 스킵(당일매수 조건매도 제외): {self._code_tag(code)} cond={cond_name}",
                              key=f"SELL_DELAY_SKIP_BOUGHT_{code}_{cond_name}", throttle=0.0)
                    self.report.emit("SELL_DELAY_SKIP_BOUGHT_TODAY", {"code": code, "cond": cond_name})
                    return

                self._schedule_delayed_sell(code, cond_name, reason="COND_EXIT", extra_delay_sec=extra_delay)

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

            self._log("INFO", f"[HOLDINGS-UPDATE] opw00018 page_parsed prev_next={prev_next} tmp_count={len(self._bal_tmp_qty)}", key="HOLD_UPD", throttle=0.2)

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

    # ✅ (A) 주문 거부/사유 MSG를 INFO로도 출력
    def _on_receive_msg(self, scr_no, rq_name, tr_code, msg):
        msg = str(msg)
        self._log("INFO", f"[MSG] scr={scr_no} rq={rq_name} tr={tr_code} msg={msg}",
                  key="MSG_INFO", throttle=0.0)
        self.report.emit("MSG", {"scr": str(scr_no), "rq": str(rq_name), "tr": str(tr_code), "msg": msg})

        sell_fail_keywords = ["매도가능수량", "주문가능수량", "잔고가 부족", "수량이 부족"]
        buy_fail_keywords = ["증거금", "예수금", "주문가능금액", "주문금액"]

        if any(k in msg for k in sell_fail_keywords):
            self._metric_inc("msg_sell_reject")
            cands = self._recent_sent_candidates("SELL", within_sec=3.0)
            if len(cands) == 1:
                self._metric_inc("msg_sell_reject_heuristic")
                self._rollback_sent_order(cands[0], source="MSG", detail=msg)
        elif any(k in msg for k in buy_fail_keywords):
            self._metric_inc("msg_buy_reject")
            cands = self._recent_sent_candidates("BUY", within_sec=3.0)
            if len(cands) == 1:
                self._metric_inc("msg_buy_reject_heuristic")
                self._rollback_sent_order(cands[0], source="MSG", detail=msg)

    # -----------------------
    # TR Request
    # -----------------------
    def request_price(self, code: str) -> Optional[int]:
        code = _norm_code(code)
        if not code:
            return None

        cached = self._get_cached_price(code, max_age_sec=2.0)
        if cached:
            return int(cached)

        if self._tr_busy:
            self._enqueue_price_request(code, why="TR_BUSY")
            self.report.emit("PRICE_QUEUED", {"code": code, "why": "TR_BUSY"})
            stale = self._get_cached_price(code, max_age_sec=PRICE_CACHE_HARD_TTL_SEC)
            if stale:
                self._metric_inc("price_use_stale_cache")
                self.report.emit("PRICE_STALE_CACHE_USED", {"code": code, "price": int(stale), "why": "TR_BUSY", "max_age_sec": float(PRICE_CACHE_HARD_TTL_SEC)})
                return int(stale)
            return None

        price = self._request_price_tr_now(code)
        if price:
            return int(price)

        stale = self._get_cached_price(code, max_age_sec=PRICE_CACHE_HARD_TTL_SEC)
        if stale:
            self._metric_inc("price_use_stale_cache")
            self.report.emit("PRICE_STALE_CACHE_USED", {"code": code, "price": int(stale), "why": "TR_FAIL", "max_age_sec": float(PRICE_CACHE_HARD_TTL_SEC)})
            return int(stale)
        return None

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

                cur_raw, _cur_field = self._get_comm_data_multi(tr_code, rq_name, i, ["현재가", "현재가(평가)", "평가금액"])
                cur = abs(_safe_int(cur_raw, 0))
                if cur > 0:
                    self._price_cache[code] = (cur, time.time())

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

        # take-profit base qty init (기준수량): 봇 시작/잔고 동기화 시 최초 보유수량을 고정
        for c, q in list(self.holdings_qty.items()):
            q = int(q or 0)
            if q > 0 and int(self.tp_base_qty.get(c, 0) or 0) <= 0:
                self.tp_base_qty[c] = q

        # take-profit state cleanup (보유수량 0이면 정리)
        for c in list(self.tp_stage.keys()):
            if self.holdings_qty.get(c, 0) <= 0:
                self.tp_stage.pop(c, None)
        for c in list(self.tp_base_qty.keys()):
            if self.holdings_qty.get(c, 0) <= 0:
                self.tp_base_qty.pop(c, None)

        for code in list(self.order_meta.keys()):
            if self.holdings_qty.get(code, 0) <= 0 and (not self._live_unfilled_exists(code)):
                if self.order_state.get(code) not in {"SELL_FILLED", "CANCELED", "REJECTED"}:
                    self._clear_order_state(code)
                self.reserved_exposure.pop(code, None)

        for code in list(self.peak_price.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.peak_price.pop(code, None)
        for code in list(self.stoploss_stage.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.stoploss_stage.pop(code, None)
        for code in list(self.stoploss_base_qty.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.stoploss_base_qty.pop(code, None)
        for code in list(self.buy_condition_by_code.keys()):
            if self.holdings_qty.get(code, 0) <= 0:
                self.buy_condition_by_code.pop(code, None)
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
        self._save_holdings_snapshot(reason="BAL_CHANGED")


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
        side_map = {}
        for od in self.unfilled_orders.values():
            c = _norm_code(od["code"])
            live_pending_codes.add(c)
            bs = str(od.get("bs", "") or "")
            side_map[c] = "BUY" if (("매수" in bs) or bs.startswith("+")) else ("SELL" if (("매도" in bs) or bs.startswith("-")) else "")
        self.pending_codes = live_pending_codes
        for c in list(live_pending_codes):
            side = side_map.get(c, str((self.order_meta.get(c, {}) or {}).get("side", "")))
            if side:
                self._set_order_state(c, f"{side}_WORKING", side=side)

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
            fill_price = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 910), 0))
            fill_qty = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 911), 0))

            # ✅ (B) 주문상태/거부사유를 반드시 로깅
            status = str(self.dynamicCall("GetChejanData(int)", 913)).strip()  # 주문상태
            reject = str(self.dynamicCall("GetChejanData(int)", 919)).strip()  # 거부사유(있을 때만)

            if order_no and code:
                self._log(
                    "INFO",
                    f"[CHEJAN-ORDER] no={order_no} {self._code_tag(code)} bs={bs} status={status} "
                    f"qty={qty} fill_qty={fill_qty} fill_price={fill_price} unfilled={unfilled} reject='{reject}'",
                    key=f"CH0_{order_no}",
                    throttle=0.0
                )
                self.report.emit("CHEJAN_ORDER_DETAIL", {
                    "order_no": order_no,
                    "code": code,
                    "bs": bs,
                    "status": status,
                    "reject": reject,
                    "qty": qty,
                    "fill_qty": fill_qty,
                    "fill_price": fill_price,
                    "unfilled": unfilled,
                })

            if order_no and code:
                side = "BUY" if (("매수" in bs) or str(bs).startswith("+")) else ("SELL" if (("매도" in bs) or str(bs).startswith("-")) else "")
                if side == "SELL" and code in self.pending_sell_meta:
                    try:
                        self.pending_sell_meta[code]["order_no"] = order_no
                        self.pending_sell_meta[code]["last_status"] = status
                    except Exception:
                        pass

                self._set_order_state(code, f"{side}_WORKING" if side and unfilled > 0 else (f"{side}_FILLED" if side and unfilled <= 0 else str(self.order_state.get(code, ""))), order_no=order_no, side=side, status=status, reject=reject, unfilled=int(unfilled), qty=int(qty))

                reject_hit = self._is_chejan_reject(status, reject)

                if unfilled > 0:
                    self.unfilled_orders[order_no] = {"code": code, "bs": bs, "qty": qty, "unfilled": unfilled, "price": 0}
                    self.pending_codes.add(code)
                else:
                    self.unfilled_orders.pop(order_no, None)
                    still = any(v["code"] == code and v.get("unfilled", 0) > 0 for v in self.unfilled_orders.values())
                    if not still:
                        self.pending_codes.discard(code)

                if side == "SELL" and fill_qty > 0:
                    self._log_sell_fill_result(code, fill_price, fill_qty, order_no=order_no, bs=bs, status=status, unfilled=unfilled)

                if reject_hit:
                    self._metric_inc("chejan_reject", reason=side or "UNKNOWN")
                    self._set_order_state(code, "REJECTED", order_no=order_no, side=side, status=status, reject=reject)
                    self._rollback_sent_order(code, source="CHEJAN_REJECT", detail=f"status={status} reject={reject}")
                    if side == "SELL":
                        self._clear_pending_sell_context(code, why="CHEJAN_REJECT")

                if side == "SELL" and unfilled <= 0 and (not reject_hit):
                    self._clear_pending_sell_context(code, why="ORDER_FILLED")

                try:
                    self._last_unfilled_sig = tuple(
                        (str(o), _norm_code(v.get("code", "")), str(v.get("bs", "")),
                         int(v.get("qty", 0) or 0), int(v.get("unfilled", 0) or 0), int(v.get("price", 0) or 0))
                        for o, v in sorted(self.unfilled_orders.items(), key=lambda x: x[0])
                    )
                except Exception:
                    pass

                self.report.emit("CHEJAN_ORDER", {"order_no": order_no, "code": code, "bs": bs, "qty": qty, "fill_qty": fill_qty, "fill_price": fill_price, "unfilled": unfilled})

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

                    if not self.buy_condition_by_code.get(code):
                        last_cond = str((self.order_meta.get(code, {}) or {}).get("cond", "") or "").strip()
                        if last_cond:
                            self.buy_condition_by_code[code] = last_cond

                    # take-profit base qty init (체잔 신규편입 시점 수량을 기준으로 고정)
                    if int(self.tp_base_qty.get(code, 0) or 0) <= 0:
                        self.tp_base_qty[code] = int(qty)
                        # tp_stage는 기본 0 유지

                    if before_qty <= 0:
                        self._set_order_state(code, "BUY_FILLED", side="BUY", holding_qty=int(qty), avg=int(avg or 0))
                        trade_id = str(self.active_trade_id_by_code.get(code, "") or self._make_trade_id(code))
                        self.active_trade_id_by_code[code] = trade_id
                        cond_name = str(self.buy_condition_by_code.get(code, "") or str((self.order_meta.get(code, {}) or {}).get("cond", "")) or "")
                        duration_sec = self._get_condition_duration_sec(code, cond_name)
                        self.report.emit("BUY_FILLED_EVENT", {
                            "code": code,
                            "trade_id": trade_id,
                            "condition": cond_name,
                            "condition_duration_sec": float(duration_sec),
                            "qty": int(qty),
                            "avg": int(avg or 0),
                            "name": name,
                        })
                    else:
                        cur_state = str(self.order_state.get(code, ""))
                        if cur_state.startswith("SELL_"):
                            self._set_order_state(code, cur_state, side="SELL", holding_qty=int(qty), avg=int(avg or 0))

                    self._save_holdings_snapshot(reason="CHEJAN_HOLDING")
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

                    self.tp_stage.pop(code, None)
                    self.tp_base_qty.pop(code, None)
                    self.buy_condition_by_code.pop(code, None)
                    self.active_trade_id_by_code.pop(code, None)
                    self._set_order_state(code, "SELL_FILLED", side="SELL", holding_qty=0)
                    self._clear_pending_sell_context(code, why="HOLDING_QTY_ZERO")
                    self._save_holdings_snapshot(reason="CHEJAN_HOLDING_ZERO")
                    self.report.emit("CHEJAN_HOLDING_ZERO", {"code": code, "name": name})

    # -----------------------
    # Buy / Sell core
    # -----------------------
    def _set_buy_block_reason(self, code: str, reason: str):
        """미매수(매수 차단/실패) 사유를 최근값으로 저장."""
        try:
            if not code:
                return
            self.last_buy_block_reason[code] = str(reason) if reason else ""
            self.last_buy_block_ts[code] = time.time()
        except Exception:
            pass

    def _get_buy_block_reason(self, code: str) -> str:
        try:
            return self.last_buy_block_reason.get(code, "")
        except Exception:
            return ""

    def _make_trade_id(self, code: str) -> str:
        code = _norm_code(code)
        now = datetime.datetime.now()
        return f"{code}_{now.strftime('%Y%m%d_%H%M%S_%f')[:-3]}"

    def _get_condition_duration_sec(self, code: str, cond: str) -> float:
        code = _norm_code(code)
        cond = str(cond or "").strip()
        if not code or not cond:
            return 0.0
        ts = float(self.condition_enter_ts.get((code, cond), 0.0) or 0.0)
        if ts <= 0:
            return 0.0
        return max(0.0, time.time() - ts)

    def _normalize_exit_reason(self, reason: str) -> str:
        s = str(reason or "").upper()
        if "STOPLOSS" in s:
            return "STOPLOSS"
        if s.startswith("TP") or "TAKEPROFIT" in s:
            return "TAKEPROFIT"
        if "TRAILING" in s:
            return "TRAILING_STOP"
        if "ORPHAN" in s:
            return "ORPHAN"
        if "FORCE" in s:
            return "FORCE_SELL"
        if "EXIT_" in s or "COND_EXIT" in s:
            return "CONDITION_EXIT"
        return "OTHER"

    def _build_sell_context_snapshot(self, code: str, price_hint: int = 0) -> Dict[str, Any]:
        code = _norm_code(code)
        hold_qty = int(self.holdings_qty.get(code, 0) or 0)
        avg = int(self.holdings_avg.get(code, 0) or 0)
        cur = int(price_hint or 0)

        if cur <= 0:
            cur = int(self._get_cached_price(code, max_age_sec=PRICE_CACHE_HARD_TTL_SEC) or 0)
        if cur <= 0:
            try:
                cur = int(self.request_price(code) or 0)
            except Exception:
                cur = 0

        pnl_amt = 0
        pnl_pct = 0.0
        if hold_qty > 0 and avg > 0 and cur > 0:
            pnl_amt = int((cur - avg) * hold_qty)
            pnl_pct = ((float(cur) - float(avg)) / float(avg)) * 100.0

        return {
            "code": code,
            "name": self._get_code_name(code),
            "hold_qty": int(hold_qty),
            "avg": int(avg),
            "cur": int(cur),
            "est_pnl_amt": int(pnl_amt),
            "est_pnl_pct": float(pnl_pct),
        }

    def _remember_sell_context(self, code: str, qty: int, reason: str, mode: str = "ALL", price_hint: int = 0, extra: Optional[Dict[str, Any]] = None):
        code = _norm_code(code)
        if not code:
            return
        extra = dict(extra or {})
        snap = self._build_sell_context_snapshot(code, price_hint=price_hint)
        meta = {
            "reason": str(reason or ""),
            "exit_reason": self._normalize_exit_reason(reason),
            "trade_id": str(self.active_trade_id_by_code.get(code, "") or ""),
            "mode": str(mode or ""),
            "req_qty": int(qty or 0),
            "snapshot_ts": _now_iso(),
            **snap,
            "extra": extra,
        }
        self.pending_sell_meta[code] = meta
        self._log(
            "INFO",
            f"[SELL-CONTEXT] {self._code_tag(code)} reason={reason} mode={mode} req_qty={int(qty or 0)} "
            f"hold_qty={snap['hold_qty']} avg={snap['avg']} now={snap['cur']} "
            f"est_pnl_amt={snap['est_pnl_amt']} est_pnl_pct={snap['est_pnl_pct']:.2f} extra={extra}",
            key=f"SELL_CONTEXT_{code}_{reason}",
            throttle=0.0,
        )
        self.report.emit("SELL_CONTEXT", meta)

    def _clear_pending_sell_context(self, code: str, why: str = ""):
        code = _norm_code(code)
        if not code:
            return
        meta = self.pending_sell_meta.pop(code, None)
        if meta:
            self.report.emit("SELL_CONTEXT_CLEAR", {"code": code, "why": str(why or ""), "reason": str(meta.get("reason", ""))})

    def _log_sell_fill_result(self, code: str, fill_price: int, fill_qty: int, order_no: str = "", bs: str = "", status: str = "", unfilled: int = 0):
        code = _norm_code(code)
        fill_price = int(fill_price or 0)
        fill_qty = int(fill_qty or 0)
        if not code or fill_price <= 0 or fill_qty <= 0:
            return

        meta = dict(self.pending_sell_meta.get(code, {}) or {})
        avg = int(meta.get("avg", 0) or self.holdings_avg.get(code, 0) or 0)
        realized_amt = 0
        realized_pct = 0.0
        if avg > 0:
            realized_amt = int((fill_price - avg) * fill_qty)
            realized_pct = ((float(fill_price) - float(avg)) / float(avg)) * 100.0

        reason = str(meta.get("reason", "UNKNOWN") or "UNKNOWN")
        self._log(
            "INFO",
            f"[SELL-FILL] no={order_no} {self._code_tag(code)} bs={bs} reason={reason} fill_qty={fill_qty} "
            f"fill_price={fill_price} avg={avg} realized_pnl_amt={realized_amt} realized_pnl_pct={realized_pct:.2f} "
            f"status={status} unfilled={int(unfilled or 0)}",
            key=f"SELL_FILL_{order_no or code}_{fill_price}_{fill_qty}",
            throttle=0.0,
        )
        self.report.emit("SELL_FILL_RESULT", {
            "code": code,
            "trade_id": str(meta.get("trade_id", self.active_trade_id_by_code.get(code, "")) or ""),
            "order_no": str(order_no or ""),
            "bs": str(bs or ""),
            "reason": reason,
            "exit_reason": str(meta.get("exit_reason", self._normalize_exit_reason(reason)) or "OTHER"),
            "fill_qty": int(fill_qty),
            "fill_price": int(fill_price),
            "avg": int(avg),
            "realized_pnl_amt": int(realized_amt),
            "realized_pnl_pct": float(realized_pct),
            "status": str(status or ""),
            "unfilled": int(unfilled or 0),
            "context": meta,
        })

    def _enqueue_buy(self, code: str, cond_name: str, ev: str = "") -> Tuple[bool, str]:
        """매수 큐에 등록. (queued, reason) 반환"""
        if not self._in_trade_window():
            reason = "OUT_OF_TRADE_WINDOW"
            code_n = _norm_code(code)
            if code_n:
                self._set_buy_block_reason(code_n, reason)
            self._log("INFO", f"[BUY-QUEUE] 스킵: 거래창 밖 -> {self._code_tag(code_n)} cond={cond_name} ev={ev} reason={reason}", key=f"BUYQ_BLOCK_WIN_{code_n}", throttle=0.5)
            self.report.emit("BUYQ_BLOCK_OUT_OF_WINDOW", {"code": code_n, "cond": cond_name, "ev": ev, "reason": reason})
            return False, reason
        code = _norm_code(code)
        if not code:
            self._log("INFO", f"[BUY-QUEUE] 스킵: 빈 종목코드 cond={cond_name} ev={ev}", key="BUYQ_EMPTY", throttle=1.0)
            self.report.emit("BUYQ_BLOCK_EMPTY_CODE", {"cond": cond_name, "ev": ev})
            return False, "EMPTY_CODE"
        if code in self.buy_queue_set:
            reason = "ALREADY_QUEUED"
            self._set_buy_block_reason(code, reason)
            self._log("INFO", f"[BUY-QUEUE] 스킵: 이미 큐에 있음 -> {self._code_tag(code)} cond={cond_name} ev={ev} reason={reason}", key=f"BUYQ_ALREADY_{code}", throttle=0.5)
            self.report.emit("BUYQ_BLOCK_ALREADY_QUEUED", {"code": code, "cond": cond_name, "ev": ev, "reason": reason})
            return False, reason
        if code in self.sold_today:
            reason = "SOLD_TODAY_REBUY_BLOCK"
            self._set_buy_block_reason(code, reason)
            self._log("INFO", f"[BUY-QUEUE] 스킵: 재매수 금지 -> {self._code_tag(code)} cond={cond_name} ev={ev} reason={reason}", key=f"BUYQ_SOLD_{code}", throttle=1.0)
            self.report.emit("BUYQ_SKIP_SOLD_TODAY", {"code": code, "cond": cond_name, "ev": ev})
            return False, reason
        self.buy_queue.append((code, cond_name, time.time()))
        self.buy_queue_set.add(code)
        self._log("INFO", f"[BUY-QUEUE] 추가: {self._code_tag(code)} cond={cond_name} queue_len={len(self.buy_queue)}",
                  key="BUYQ_ADD", throttle=0.2)
        self.report.emit("BUYQ_ADD", {"code": code, "cond": cond_name, "ev": ev, "queue_len": len(self.buy_queue)})
        return True, "QUEUED"
    def _send_order(self, name: str, screen: str, acc: str, order_type: int,
                    code: str, qty: int, price: int, hoga: str, org_order_no: str) -> int:
        
        # 장중(거래 시간) 여부 확인
        if not self._in_trade_window():
            self._log('INFO', f"[TRADE-WINDOW] 주문차단(창 밖): name={name} type={order_type} code={self._code_tag(code)} qty={qty}", key=f"TW_BLOCK_{code}", throttle=0.5)
            return -999
            
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
        if not self._in_trade_window():
            return
        if not self.buy_queue:
            return
        if self._risk_running:
            return

        code, cond, ts = self.buy_queue.popleft()
        self._log("INFO", f"[BUY-ENTRY] code={code} cond={cond} queue_ts={ts:.0f}", key=f"BUY_ENTRY_{code}", throttle=0.2)
        prev_reason = self._get_buy_block_reason(code)
        if prev_reason:
            self._log("INFO", f"[BUY-ENTRY] 최근 미매수 사유: {self._code_tag(code)} reason={prev_reason}", key=f"BUY_ENTRY_REASON_{code}", throttle=0.5)
        self.buy_queue_set.discard(code)

        self._log("INFO", f"[BUY] 진입: target={self._code_tag(code)} budget={TARGET_BUY_AMOUNT} cond={cond}",
                  key=f"BUY_ENTER_{code}", throttle=0.2)

        price = self.request_price(code)
        if not price:
            self._set_buy_block_reason(code, "NO_PRICE")
            self._log("WARN", f"[BUY] 현재가 실패 -> 스킵: {self._code_tag(code)} cond={cond}", key=f"BUY_NOPRICE_{code}", throttle=1.0)
            self._metric_inc("buy_skip_no_price", reason=cond)
            self.report.emit("BUY_SKIP_NO_PRICE", {"code": code, "cond": cond})
            return

        ok, reason = self._can_buy_code(code, price)
        if not ok:
            self._set_buy_block_reason(code, reason)
            self._log("INFO", f"[BUY] 스킵: {self._code_tag(code)} cond={cond} reason={reason}", key=f"BUY_BLOCK_{code}", throttle=0.5)
            self._metric_inc("buy_block", reason=reason)
            self.report.emit("BUY_BLOCK", {"code": code, "cond": cond, "reason": reason, "price": price})
            return

        qty = TARGET_BUY_AMOUNT // price
        if qty <= 0:
            self._set_buy_block_reason(code, "QTY_ZERO")
            self._log("INFO", f"[BUY] 스킵: 수량=0 -> {self._code_tag(code)} cond={cond} price={price} budget={TARGET_BUY_AMOUNT}", key=f"BUY_QTY0_{code}", throttle=1.0)
            self.report.emit("BUY_SKIP_QTY0", {"code": code, "cond": cond, "price": price})
            return

        order_amount = qty * price
        self._log("INFO", f"[BUY] 계산: {self._code_tag(code)} price={price} qty={qty} order_amount={order_amount}",
                  key=f"BUY_CALC_{code}", throttle=0.2)

        self.report.emit("BUY_ATTEMPT", {"code": code, "cond": cond, "price": price, "qty": qty, "order_amount": order_amount})

        ret = self._send_order("BUY", "0101", self.account, 1, code, qty, 0, "03", "")
        if ret == 0:
            # ✅ (C) ret=0은 "주문요청 전송 성공"일 뿐. 체결/접수는 체잔으로 확인.
            self._log(
                "INFO",
                f"[BUY] ✅ 주문요청 성공(SendOrder ret=0): {self._code_tag(code)} qty={qty} cond={cond}",
                key=f"BUY_SENT_{code}",
                throttle=0.0,
            )
            self._set_buy_block_reason(code, "")  # 성공 시 최근 미매수 사유 초기화
            self.pending_codes.add(code)

            self.reserved_exposure[code] += int(order_amount)
            self._log("INFO", f"[BUY-LIMIT] reserved_exposure += {order_amount} -> now={self.reserved_exposure[code]} for {self._code_tag(code)}",
                      key=f"RESERVE_ADD_{code}", throttle=0.0)

            pre_bought_added = False
            if code not in self.bought_today:
                self.bought_today.add(code)
                pre_bought_added = True
                self._log("INFO", f"[BOUGHT-TODAY] ✅ 주문요청 선등록(체잔 전): {self._code_tag(code)}",
                          key=f"BOUGHT_PRE_{code}", throttle=0.0)
                if BOUGHT_TODAY_PERSIST:
                    self._save_bought_today()
                self.report.emit("BOUGHT_TODAY_ADD", {"code": code, "via": "BUY_SENT_PRE_CHEJAN"})

            self._register_sent_order(code, "BUY", qty, price, reason="BUY", cond=cond, reserved_amount=order_amount, pre_bought_added=pre_bought_added)
            self.buy_condition_by_code[code] = str(cond or "")
            self._save_holdings_snapshot(reason="BUY_SENT_COND_TRACK")
            self.report.emit("BUY_SENT", {"code": code, "cond": cond, "qty": qty, "price": price, "order_amount": order_amount,
                                          "reserved_exposure": int(self.reserved_exposure.get(code, 0) or 0), "pre_bought_added": bool(pre_bought_added)})
        else:
            self._log("WARN", f"[BUY] ❌ 주문실패: {self._code_tag(code)} cond={cond} ret={ret}", key=f"BUY_FAIL_{code}", throttle=1.0)
            self._set_buy_block_reason(code, f"ORDER_FAIL_{ret}")
            self.report.emit("BUY_FAIL", {"code": code, "cond": cond, "ret": ret, "price": price, "qty": qty})

    # -----------------------
    # Delayed sell / risk / orphan
    # -----------------------
    def _schedule_delayed_sell(self, code: str, cond: str, reason: str, extra_delay_sec: float = 0.0):
        if not self._in_trade_window():
            return
        code = _norm_code(code)
        if not code:
            return
        extra_delay_sec = max(0.0, float(extra_delay_sec or 0.0))
        due_in_sec = float(SELL_DELAY_SEC) + extra_delay_sec
        due = time.time() + due_in_sec
        self.delayed_sells[(code, cond)] = due
        if extra_delay_sec > 0:
            self._metric_inc("sell_delay_extra_due_chatter", reason=cond)
        self._log("INFO", f"[SELL-DELAY] 예약: {self._code_tag(code)} cond={cond} after={due_in_sec}s (base={SELL_DELAY_SEC}s extra={extra_delay_sec}s reason={reason})",
                  key=f"SELL_SCHED_{code}_{cond}", throttle=0.2)
        self.report.emit("SELL_DELAY_SCHEDULE", {"code": code, "cond": cond, "due_in_sec": due_in_sec, "base_delay_sec": float(SELL_DELAY_SEC), "extra_delay_sec": extra_delay_sec, "reason": reason})

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

    def _request_sell_qty(self, code: str, qty: int, reason: str, price_hint: int = 0, extra: Optional[Dict[str, Any]] = None):
        code = _norm_code(code)
        hold_qty = int(self.holdings_qty.get(code, 0) or 0)
        qty = int(qty or 0)
        if hold_qty <= 0 or qty <= 0:
            return
        if qty > hold_qty:
            qty = hold_qty
        if code in self.pending_codes or self._has_pending_sell(code):
            self._metric_inc("sell_skip_pending", reason=reason)
            self._log("INFO", f"[SELL] 스킵: 이미 주문/미체결 존재 -> {self._code_tag(code)} qty={qty} reason={reason}", key=f"SELL_SKIP_PENDING_{code}_{reason}", throttle=0.2)
            self.report.emit("SELL_SKIP_PENDING", {"code": code, "qty": qty, "reason": reason, "mode": "PARTIAL"})
            return

        self._remember_sell_context(code, qty=qty, reason=reason, mode="PARTIAL", price_hint=price_hint, extra=extra)
        self._log("INFO", f"[SELL] 주문전송: {self._code_tag(code)} qty={qty} 시장가 reason={reason}",
                  key=f"SELL_SEND_{code}_{reason}", throttle=0.0)

        self.report.emit("SELL_ATTEMPT", {"code": code, "qty": qty, "reason": reason, "mode": "PARTIAL"})

        ret = self._send_order("SELL", "0102", self.account, 2, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[SELL] ✅ 주문성공: {self._code_tag(code)} qty={qty} reason={reason}",
                      key=f"SELL_OK_{code}_{reason}", throttle=0.0)
            self.pending_codes.add(code)
            self._set_rebuy_block(code, cooldown_sec=REBUY_COOLDOWN_SEC)
            self._register_sent_order(code, "SELL", qty, 0, reason=reason)
            self.report.emit("SELL_OK", {"code": code, "qty": qty, "reason": reason, "mode": "PARTIAL"})
        else:
            self._log("WARN", f"[SELL] ❌ 주문실패: {self._code_tag(code)} ret={ret} reason={reason}",
                      key=f"SELL_FAIL_{code}_{reason}", throttle=0.5)
            self._clear_pending_sell_context(code, why=f"SEND_FAIL_{ret}")
            self.report.emit("SELL_FAIL", {"code": code, "qty": qty, "reason": reason, "ret": ret, "mode": "PARTIAL"})

    def _request_sell_all(self, code: str, reason: str, price_hint: int = 0, extra: Optional[Dict[str, Any]] = None):
        code = _norm_code(code)
        qty = int(self.holdings_qty.get(code, 0) or 0)
        if qty <= 0:
            return
        if code in self.pending_codes or self._has_pending_sell(code):
            self._metric_inc("sell_skip_pending", reason=reason)
            self._log("INFO", f"[SELL] 스킵(전량): 이미 주문/미체결 존재 -> {self._code_tag(code)} reason={reason}", key=f"SELL_SKIP_PENDING_ALL_{code}_{reason}", throttle=0.2)
            self.report.emit("SELL_SKIP_PENDING", {"code": code, "qty": qty, "reason": reason, "mode": "ALL"})
            return

        self._remember_sell_context(code, qty=qty, reason=reason, mode="ALL", price_hint=price_hint, extra=extra)
        self._log("INFO", f"[SELL] 주문전송(전량): {self._code_tag(code)} qty={qty} 시장가 reason={reason}",
                  key=f"SELL_SEND_{code}", throttle=0.2)

        self.report.emit("SELL_ATTEMPT", {"code": code, "qty": qty, "reason": reason, "mode": "ALL"})

        ret = self._send_order("SELL", "0102", self.account, 2, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[SELL] ✅ 주문성공: {self._code_tag(code)} qty={qty} reason={reason}",
                      key=f"SELL_OK_{code}", throttle=0.2)
            self.pending_codes.add(code)
            self._set_rebuy_block(code, cooldown_sec=REBUY_COOLDOWN_SEC)
            self._register_sent_order(code, "SELL", qty, 0, reason=reason)
            self.report.emit("SELL_OK", {"code": code, "qty": qty, "reason": reason, "mode": "ALL"})
        else:
            self._log("WARN", f"[SELL] ❌ 주문실패: {self._code_tag(code)} ret={ret}",
                      key=f"SELL_FAIL_{code}", throttle=1.0)
            self._clear_pending_sell_context(code, why=f"SEND_FAIL_{ret}")
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
        if not self._in_trade_window():
            return
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
                self._metric_inc("sell_delay_skip_pending", reason=cond)
                continue

            self._request_sell_all(code, reason=f"EXIT_{cond}_DELAY{SELL_DELAY_SEC}s", extra={"cond": cond, "source": "DELAYED_COND_EXIT"})

    def _orphan_sweeper_tick(self):
        if not self._in_trade_window():
            return
        if self._risk_running:
            return
        try:
            now_dt = datetime.datetime.now()
            if str(now_dt.strftime("%H%M")) < str(ORPHAN_ENABLE_AFTER_HHMM):
                if self.orphan_first_seen:
                    self.orphan_first_seen.clear()
                self._log("INFO", f"[ORPHAN] 장초반 유예중: {ORPHAN_ENABLE_AFTER_HHMM} 이전에는 스위퍼 비활성화", key="ORPHAN_STARTUP_GRACE", throttle=30.0)
                self.report.emit("ORPHAN_STARTUP_GRACE", {"enable_after_hhmm": str(ORPHAN_ENABLE_AFTER_HHMM)})
                return

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
                    self._request_sell_all(code, reason="ORPHAN_SWEEPER", extra={"source": "ORPHAN_SWEEPER", "grace_sec": float(ORPHAN_GRACE_SEC)})
                    self.orphan_first_seen[code] = now + 999999

        except Exception as e:
            self._log("ERROR", f"[ORPHAN] sweeper error: {e}", key="ORPHAN_ERR", throttle=5.0)
            self.report.emit("ORPHAN_ERR", {"err": repr(e)})

    def _risk_monitor_tick(self):
        if not self._in_trade_window():
            return
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

                if code in self.pending_codes:
                    self._metric_inc("risk_skip_pending", reason=code)
                    self.report.emit("RISK_SKIP_PENDING", {"code": code, "order_state": str(self.order_state.get(code, ""))})
                    continue

                if self._has_pending_sell(code):
                    self._metric_inc("risk_skip_pending_sell", reason=code)
                    continue

                cfg = self._get_risk_config_for_code(code)
                cond_name = str(cfg.get("cond", "") or "")
                try:
                    tiers = list(cfg.get("stoploss_tiers", []) or [])
                    tiers = sorted(tiers, key=lambda x: float(x[0]), reverse=True)
                except Exception:
                    tiers = []
                try:
                    tp_levels = list(cfg.get("takeprofit_levels", []) or [])
                    tp_levels = sorted(tp_levels, key=lambda x: float(x[0]))
                except Exception:
                    tp_levels = []
                tp_min_qty = max(1, int(cfg.get("takeprofit_min_qty", TAKEPROFIT_MIN_QTY) or 1))
                trailing_stop_pct = abs(float(cfg.get("trailing_stop_pct", TRAILING_STOP_PCT) or 0.0))

                cur = self.request_price(code)
                if (not cur or cur <= 0):
                    cur = self._get_cached_price(code, max_age_sec=PRICE_CACHE_HARD_TTL_SEC)
                    if cur and cur > 0:
                        self._metric_inc("risk_use_stale_price", reason=code)
                        self.report.emit("RISK_USE_STALE_PRICE", {"code": code, "price": int(cur), "max_age_sec": float(PRICE_CACHE_HARD_TTL_SEC)})
                if not cur or cur <= 0:
                    self._log("DEBUG", f"[RISK] 현재가 미확보(큐잉/실패) -> 스킵: {self._code_tag(code)}",
                              key=f"RISK_NOPRICE_{code}", throttle=1.0)
                    self._metric_inc("risk_skip_noprice", reason=code)
                    self.report.emit("RISK_SKIP_NOPRICE", {"code": code})
                    continue

                peak = int(self.peak_price.get(code, 0) or 0)
                if bool(cfg.get("trailing_enabled", False)):
                    if peak <= 0 or cur > peak:
                        self.peak_price[code] = cur
                        peak = cur

                pnl_pct = (float(cur - avg) / float(avg)) * 100.0
                stage = int(self.stoploss_stage.get(code, 0) or 0)

                # 1) STOPLOSS tiers (한 tick에 한 단계만)
                if bool(cfg.get("stoploss_enabled", False)) and tiers and stage < len(tiers):
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
                            "sell_qty": int(sell_qty), "cur_qty": int(cur_qty), "cond": cond_name,
                        })

                        if sell_qty >= cur_qty:
                            self._request_sell_all(code, reason=f"STOPLOSS{stage+1}@{thr_pct}%", price_hint=cur, extra={"source": "STOPLOSS", "stage": int(stage + 1), "threshold_pct": float(thr_pct)})
                        else:
                            self._request_sell_qty(code, sell_qty, reason=f"STOPLOSS{stage+1}@{thr_pct}%", price_hint=cur, extra={"source": "STOPLOSS", "stage": int(stage + 1), "threshold_pct": float(thr_pct)})

                        self.stoploss_stage[code] = stage + 1
                        continue

                # 2) TRAILING (즉시 전량)
                if bool(cfg.get("trailing_enabled", False)) and peak > 0 and trailing_stop_pct > 0:
                    drawdown_pct = (float(cur - peak) / float(peak)) * 100.0
                    if drawdown_pct <= -abs(float(trailing_stop_pct)):
                        self._log(
                            "INFO",
                            f"[RISK] TRAILING 발동: {self._code_tag(code)} cur={cur} peak={peak} dd={drawdown_pct:.2f}%",
                            key=f"TRAIL_FIRE_{code}",
                            throttle=0.0,
                        )
                        self._clear_delayed_sells_code(code)
                        self.report.emit("TRAILING_FIRE", {
                            "code": code, "cur": cur, "peak": peak,
                            "drawdown_pct": drawdown_pct, "trail_pct": float(trailing_stop_pct), "cond": cond_name
                        })
                        self._request_sell_all(code, reason=f"TRAILING_{trailing_stop_pct:.0f}", price_hint=cur, extra={"source": "TRAILING", "trail_pct": float(trailing_stop_pct), "peak": int(peak), "cond": cond_name})
                        continue
                # 3) TAKEPROFIT (부분익절): +10% 30%, +15% 30% (기준수량의 30%)
                if bool(cfg.get("takeprofit_enabled", False)):
                    levels = tp_levels

                    tp_stage = int(self.tp_stage.get(code, 0) or 0)
                    if tp_stage < len(levels):
                        thr_pct, ratio = levels[tp_stage]

                        if pnl_pct >= float(thr_pct):
                            # pending_codes에 이미 있으면(주문/미체결) stage 업데이트 판별이 애매하므로 스킵
                            if code in self.pending_codes:
                                self.report.emit("TAKEPROFIT_SKIP_PENDING", {"code": code, "stage": tp_stage, "thr_pct": float(thr_pct)})
                            else:
                                base_qty = int(self.tp_base_qty.get(code, 0) or 0)
                                if base_qty <= 0:
                                    base_qty = int(qty)

                                sell_qty = int(base_qty * float(ratio))
                                sell_qty = max(int(tp_min_qty), sell_qty)
                                sell_qty = max(1, min(int(qty), int(sell_qty)))

                                self._log(
                                    "INFO",
                                    f"[TAKEPROFIT] 발동: {self._code_tag(code)} pnl={pnl_pct:.2f}% "
                                    f"stage={tp_stage+1}/{len(levels)} thr=+{thr_pct}% ratio={ratio} base_qty={base_qty} sell_qty={sell_qty}",
                                    key=f"TP_FIRE_{code}_{tp_stage}",
                                    throttle=0.0,
                                )
                                self.report.emit("TAKEPROFIT_FIRE", {
                                    "code": code, "cur": cur, "avg": avg, "pnl_pct": pnl_pct,
                                    "stage": tp_stage, "thr_pct": float(thr_pct), "ratio": float(ratio),
                                    "base_qty": int(base_qty), "sell_qty": int(sell_qty), "cur_qty": int(qty),
                                })

                                self._save_holdings_snapshot(reason="BEFORE_TAKEPROFIT")

                                before_pending = (code in self.pending_codes)
                                self._request_sell_qty(code, sell_qty, reason=f"TP{tp_stage+1}@+{thr_pct}%", price_hint=cur, extra={"source": "TAKEPROFIT", "stage": int(tp_stage + 1), "threshold_pct": float(thr_pct)})
                                after_pending = (code in self.pending_codes)

                                # 주문요청(ret=0) 성공이면 pending_codes에 들어가므로 그때만 stage 업데이트
                                if (not before_pending) and after_pending:
                                    self.tp_stage[code] = tp_stage + 1
                                    self.report.emit("TAKEPROFIT_STAGE_SET", {"code": code, "stage": int(self.tp_stage.get(code, 0) or 0)})
                                else:
                                    self._log(
                                        "WARN",
                                        f"[TAKEPROFIT] 주문요청 실패/미확인 -> stage 유지: {self._code_tag(code)} stage={tp_stage} thr=+{thr_pct}%",
                                        key=f"TP_NOADV_{code}_{tp_stage}",
                                        throttle=2.0,
                                    )
                                    self.report.emit("TAKEPROFIT_NO_STAGE_ADVANCE", {"code": code, "stage": tp_stage, "thr_pct": float(thr_pct)})

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
        if not self._in_trade_window():
            return
        if self._risk_running:
            return
        self._cleanup_stale_order_intents()
        self._emit_metrics_snapshot(reason="PERIODIC_SYNC")
        if self._tr_busy:
            return
        self.request_balance()
        self.request_unfilled()

    # -----------------------
    # Persistence
    # -----------------------
    def _make_holdings_snapshot(self) -> Dict[str, Any]:
        # 현재 보유종목(계좌) 스냅샷을 JSON으로 저장하기 위한 dict 생성
        items = []
        for code in sorted(self.holdings_qty.keys()):
            qty = int(self.holdings_qty.get(code, 0) or 0)
            if qty <= 0:
                continue
            items.append({
                "code": code,
                "name": self._get_code_name(code),
                "qty": qty,
                "avg": int(self.holdings_avg.get(code, 0) or 0),
                "buy_cond": str(self.buy_condition_by_code.get(code, "") or ""),
                "trade_id": str(self.active_trade_id_by_code.get(code, "") or ""),
                "tp_stage": int(self.tp_stage.get(code, 0) or 0),
                "tp_base_qty": int(self.tp_base_qty.get(code, 0) or 0),
            })
        return {
            "day": _today_yyyymmdd(),
            "ts": _now_iso(),
            "account": self.account,
            "holdings": items,
        }

    def _save_holdings_snapshot(self, reason: str = "", force: bool = False):
        if not HOLDINGS_SNAPSHOT_ENABLED:
            return
        try:
            snap = self._make_holdings_snapshot()
            sig = tuple((x["code"], int(x["qty"]), int(x["avg"]), str(x.get("buy_cond", "")), str(x.get("trade_id", "")), int(x.get("tp_stage", 0)), int(x.get("tp_base_qty", 0)))
                        for x in snap.get("holdings", []))
            if (not force) and (sig == self._holdings_snapshot_sig):
                return
            self._holdings_snapshot_sig = sig

            with open(HOLDINGS_SNAPSHOT_FILE, "w", encoding="utf-8") as f:
                json.dump(snap, f, ensure_ascii=False, indent=2)

            self.report.emit("HOLDINGS_SNAPSHOT_SAVED", {
                "file": HOLDINGS_SNAPSHOT_FILE,
                "count": len(snap.get("holdings", [])),
                "reason": reason,
            })
        except Exception as e:
            self._log("WARN", f"[HOLDINGS] snapshot save failed: {e}", key="HOLD_SNAP_FAIL", throttle=2.0)
            try:
                self.report.emit("HOLDINGS_SNAPSHOT_SAVE_FAIL", {"err": repr(e), "reason": reason})
            except Exception:
                pass

    def _load_holdings_snapshot(self):
        try:
            if not os.path.exists(HOLDINGS_SNAPSHOT_FILE):
                return
            with open(HOLDINGS_SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            day = str(data.get("day", "") or "")
            if day != _today_yyyymmdd():
                return
            loaded = 0
            for item in list(data.get("holdings", []) or []):
                code = _norm_code(item.get("code", ""))
                if not code:
                    continue
                cond = str(item.get("buy_cond", "") or "").strip()
                if cond:
                    self.buy_condition_by_code[code] = cond
                trade_id = str(item.get("trade_id", "") or "").strip()
                if trade_id:
                    self.active_trade_id_by_code[code] = trade_id
                tp_stage = int(item.get("tp_stage", 0) or 0)
                tp_base_qty = int(item.get("tp_base_qty", 0) or 0)
                if tp_stage > 0:
                    self.tp_stage[code] = tp_stage
                if tp_base_qty > 0:
                    self.tp_base_qty[code] = tp_base_qty
                loaded += 1
            if loaded > 0:
                self._log("INFO", f"[HOLDINGS] snapshot load: {loaded}개 buy_condition/tp 복원", key="HOLD_SNAP_LOAD", throttle=0.0)
                self.report.emit("HOLDINGS_SNAPSHOT_LOADED", {"count": int(loaded), "file": HOLDINGS_SNAPSHOT_FILE})
        except Exception as e:
            self._log("WARN", f"[HOLDINGS] snapshot load failed: {e}", key="HOLD_SNAP_LOAD_FAIL", throttle=2.0)
            try:
                self.report.emit("HOLDINGS_SNAPSHOT_LOAD_FAIL", {"err": repr(e)})
            except Exception:
                pass

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

        # 거래 시간 게이트: 창 안이면 즉시 구독/타이머 시작, 창 밖이면 대기(guard가 진입 시 시작)
        if self._in_trade_window():
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
            self._emit_metrics_snapshot(reason="RUN_START", force=True)
        else:
            self._log("INFO", f"[TRADE-WINDOW] 현재 창 밖({TRADE_START_HHMM}-{TRADE_END_HHMM}) -> 자동매매 대기(ON 되면 시작)")
            self.report.emit("RUN_WAIT_WINDOW", {"start": TRADE_START_HHMM, "end": TRADE_END_HHMM})
            self._emit_metrics_snapshot(reason="RUN_WAIT_WINDOW", force=True)
            # 혹시 모를 잔여 구독/타이머 정리
            self._stop_trade_session()



def main():
    month_dir = os.path.join(LOG_DIR, MONTH)
    os.makedirs(month_dir, exist_ok=True)

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