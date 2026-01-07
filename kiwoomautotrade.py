# -*- coding: utf-8 -*-
"""
Kiwoom OpenAPI+ 자동매매 통합본 (단일 파일)


✅ 반영/수정 사항
1) SendOrder dynamicCall TypeError 해결: 인자를 리스트로 넘기도록 고정
2) BUY/SELL 조건 겹칠 때 중복 구독 방지 (w3, x2 같은 케이스)
3) 매입 한도(MAX_POSITION_PER_CODE) = 보유+미체결+대기예산 포함 강제
4) 재시작해도 당일 재매수 금지 유지(sold_today.json)
5) 매도 직후 재매수 방지: 쿨다운 + qty=0 확정 전까지 차단
6) 조건 이탈 매도는 지연매도(SELL_DELAY_SEC)만 스케줄 (즉시 매도 금지)
7) 고아종목 제거(SELL 조건 어디에도 미편입이면 ORPHAN_GRACE 후 전량 매도)
8) ✅ TR 직렬화(중복 QEventLoop exec 방지): request_price/balance/unfilled 모두 단일 TR 게이트로 처리
9) ✅ [INFO] 로그에 종목코드 옆에 종목명 표시(간섭 최소)
    - TR 요청 추가 없이 GetMasterCodeName(QString)로 조회
    - 캐시 적용(동일 종목 반복 호출 최소화)
10) ✅ 잔고/미체결 로그를 '변화가 있을 때만' 출력
11) ✅ (요청사항) 비밀번호/ShowAccountWindow 흐름을 "아래 코드 방식"으로 교체
    - PASSWD="" 가능
    - start()에서 show_account_window() 호출
    - opw00018 잔고조회 SetInputValue("비밀번호", PASSWD) 사용


환경: Windows / Python 32-bit / PyQt5 / Kiwoom OpenAPI+
중요: QAxWidget import = PyQt5.QAxContainer
"""


import sys
import os
import time
import json
import datetime
from collections import deque, defaultdict
from typing import Dict, Any, Optional, Set, Tuple


from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer


# =========================
# 사용자 설정 (여기만 건드리면 됨)
# =========================
BUY_COND_NAMES = {"w3", "x2"} # ✅ 매수 조건
SELL_COND_NAMES = {"w", "w3", "x2"} # ✅ 매도 조건(이탈 트리거 → 지연매도)


TARGET_BUY_AMOUNT = 100000 # 1회 매수 예산(분할매수 단위)
MAX_POSITION_PER_CODE = 100000 # ✅ 종목당 누적 노출(보유+미체결+대기예산) 최대 한도
ALLOW_ADD_BUY = False  # ✅ 동일 종목 추가매수 허용 여부


SELL_DELAY_SEC = 3.0    # ✅ 조건 이탈 시 지연매도 (우선)
REBUY_COOLDOWN_SEC = 60.0  # ✅ 매도 직후 임시 재매수 금지(초)
SOLD_TODAY_PERSIST = True  # ✅ 당일 재매수 금지 저장/로드


STOPLOSS_PCT = -10.0    # (옵션)
AUTO_SELL_INTERVAL_SEC = 60 # (옵션)


BALANCE_COOLDOWN_SEC = 3.0 # 잔고 TR 과호출 방지
PRICE_REQ_INTERVAL = 0.25  # 현재가 TR 과호출 방지
PRICE_RETRY_MAX = 3
PRICE_RETRY_SLEEP = 0.8


# 고아종목 제거
ORPHAN_CHECK_INTERVAL_SEC = 10
ORPHAN_GRACE_SEC = 60


# 로그 설정
LOG_LEVEL = "INFO"  # "DEBUG" / "INFO" / "WARN" / "ERROR"
LOG_TO_FILE = True
LOG_FILE_PATH = "kiwoomautotrade.log"
LOG_THROTTLE_SEC = 2.0
LOG_PRINT_CODELIST_MAX = 8


# ✅ (교체) 비밀번호 입력 방식(아래 코드 방식)
# - 보안상 비추천: 하드코딩하지 말고 PASSWD=""로 두고 ShowAccountWindow에서 입력해도 됨
PASSWD = ""    # "" 가능
PASSWD_MEDIA = "00" # 00: 공통


# 화면번호(중복 사용 금지)
SCREEN_LOGIN = "0000"
SCREEN_TR_PRICE = "2000"
SCREEN_TR_BAL = "2100"
SCREEN_TR_UNFILLED = "2200"
SCREEN_COND_BASE = 5000


# 파일(당일 재매수 금지 저장)
SOLD_TODAY_FILE = "sold_today.json"




def _today_yyyymmdd() -> str:
    return datetime.datetime.now().strftime("%Y%m%d")




def _now_hms() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")




def _safe_int(x, default=0) -> int:
    try:
        return int(str(x).strip())
    except:
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




class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()


        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")


        # 이벤트 연결
        self.OnEventConnect.connect(self._on_event_connect)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)
        self.OnReceiveMsg.connect(self._on_receive_msg)


        self.OnReceiveConditionVer.connect(self._on_receive_condition_ver)
        self.OnReceiveTrCondition.connect(self._on_receive_tr_condition)
        self.OnReceiveRealCondition.connect(self._on_receive_real_condition)


        self.OnReceiveChejanData.connect(self._on_receive_chejan_data)


        # 루프
        self.login_loop = QEventLoop()
        self.tr_loop = QEventLoop()
        self.cond_loop = QEventLoop()


        # ✅ TR 직렬화(중복 exec 방지)
        self._tr_busy = False
        self._tr_busy_name = ""
        self._tr_deadline = 0.0
        self._tr_timeout_timer = QTimer()
        self._tr_timeout_timer.timeout.connect(self._tr_timeout_tick)


        # 상태
        self.account = ""
        self.server_gubun = ""


        # 조건
        self.cond_name_to_idx: Dict[str, int] = {}
        self.subscribed_conds: Set[str] = set() # ✅ 중복구독 방지용


        # sell 조건 편입 집합(고아 판정용)
        self.sell_cond_members: Dict[str, Set[str]] = defaultdict(set)


        # 보유/미체결
        self.holdings_qty: Dict[str, int] = defaultdict(int)
        self.holdings_avg: Dict[str, int] = defaultdict(int)
        self.holdings_name: Dict[str, str] = defaultdict(str)


        self.unfilled_orders: Dict[str, Dict[str, Any]] = {}
        self.pending_codes: Set[str] = set()


        # 매수 큐
        self.buy_queue = deque() # (code, cond_name, ts)
        self.buy_queue_set: Set[str] = set()


        # 지연매도 예약
        self.delayed_sells: Dict[Tuple[str, str], float] = {}


        # 재매수 금지
        self.sold_today: Set[str] = set()
        self.rebuy_block_until: Dict[str, float] = {}
        self.rebuy_block_pending: Set[str] = set()


        # 고아 추적
        self.orphan_first_seen: Dict[str, float] = {}


        # TR 결과 임시
        self._price_resp: Dict[str, int] = {}
        self._last_price_req_ts = 0.0
        self._last_balance_req_ts = 0.0


        # ✅ (추가) 종목명 캐시 (TR 없이 조회)
        self.code_name_cache: Dict[str, str] = {}


        # ✅ (추가) 잔고/미체결 "변화 있을 때만" 로그용 signature
        self._last_balance_sig = None
        self._last_unfilled_sig = None


        # 로거
        self._init_logger()


        self._log("INFO", "======================================================================")
        self._log("INFO", "[BOOT-CONFIG] 실제 실행중 설정값 확인 (이게 로그랑 다르면 '다른 파일 실행중'임)")
        self._log("INFO", f"[BOOT-CONFIG] __file__={__file__}")
        self._log("INFO", f"[BOOT-CONFIG] BUY_COND_NAMES={BUY_COND_NAMES}")
        self._log("INFO", f"[BOOT-CONFIG] SELL_COND_NAMES={SELL_COND_NAMES}")
        self._log("INFO", f"[BOOT-CONFIG] TARGET_BUY_AMOUNT={TARGET_BUY_AMOUNT}")
        self._log("INFO", f"[BOOT-CONFIG] MAX_POSITION_PER_CODE={MAX_POSITION_PER_CODE}")
        self._log("INFO", f"[BOOT-CONFIG] ALLOW_ADD_BUY={ALLOW_ADD_BUY}")
        self._log("INFO", f"[BOOT-CONFIG] PASSWD={'(EMPTY)' if PASSWD=='' else '(SET)'} / PASSWD_MEDIA={PASSWD_MEDIA}")
        self._log("INFO", "======================================================================")


        if SOLD_TODAY_PERSIST:
            self._load_sold_today()


        # 타이머
        self.buy_timer = QTimer()
        self.buy_timer.timeout.connect(self._process_buy_queue_tick)


        self.delay_sell_timer = QTimer()
        self.delay_sell_timer.timeout.connect(self._process_delayed_sells_tick)


        self.orphan_timer = QTimer()
        self.orphan_timer.timeout.connect(self._orphan_sweeper_tick)


        self.sync_timer = QTimer()
        self.sync_timer.timeout.connect(self._periodic_sync_tick)


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
            except:
                pass


    # -----------------------
    # ✅ (추가) ShowAccountWindow (아래 코드 방식)
    # -----------------------
    def show_account_window(self):
        """
        ✅ KOA_Functions('ShowAccountWindow') 호출
        - PASSWD를 비워둔 경우에도 사용자 입력창 흐름을 통과하기 위한 용도
        """
        self._log("INFO", "[UI] KOA_Functions('ShowAccountWindow') 호출(계좌/비번 입력창 유도)", key="SHOW_ACC_WIN", throttle=1.0)
        try:
            ret = self.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
            self._log("INFO", f"[UI] ShowAccountWindow ret={ret}", key="SHOW_ACC_WIN_RET", throttle=1.0)
        except Exception as e:
            self._log("WARN", f"[UI] ShowAccountWindow 호출 실패: {repr(e)}", key="SHOW_ACC_WIN_FAIL", throttle=2.0)


    # -----------------------
    # ✅ 종목명 표시 (TR 간섭 없이)
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
        except:
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
    # ✅ TR 직렬화 게이트
    # -----------------------
    def _tr_request(self, rq_name: str, tr_code: str, screen: str, timeout_sec: float = 8.0) -> bool:
        if self._tr_busy:
            self._log(
                "WARN",
                f"[TR-SERIAL] busy({self._tr_busy_name}) -> skip rq={rq_name}",
                key=f"TR_BUSY_{rq_name}",
                throttle=0.5,
            )
            return False


        self._tr_busy = True
        self._tr_busy_name = rq_name
        self._tr_deadline = time.time() + float(timeout_sec)


        ret = self.dynamicCall("CommRqData(QString, QString, int, QString)", rq_name, tr_code, 0, screen)
        if ret != 0:
            self._log(
                "WARN",
                f"[TR-SERIAL] CommRqData fail ret={ret} rq={rq_name}",
                key=f"TR_FAIL_{rq_name}",
                throttle=0.5,
            )
            self._tr_busy = False
            self._tr_busy_name = ""
            self._tr_deadline = 0.0
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
        except:
            pass


    def _tr_timeout_tick(self):
        if not self._tr_busy:
            try:
                self._tr_timeout_timer.stop()
            except:
                pass
            return


        if time.time() >= self._tr_deadline:
            self._log(
                "ERROR",
                f"[TR-SERIAL] timeout rq={self._tr_busy_name} -> force release",
                key=f"TR_TO_{self._tr_busy_name}",
                throttle=0.5,
            )
            self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()


    # -----------------------
    # Login
    # -----------------------
    def comm_connect(self):
        self._log("INFO", "[LOGIN] GetConnectState=%s (1=연결,0=미연결)" % self.dynamicCall("GetConnectState()"))
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
            self.account = accs[0] if accs else ""
            self._log("INFO", f"[LOGIN] 계좌번호: {self.account}")
            self.server_gubun = str(self.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")).strip()
            self._log("INFO", f"[LOGIN] 서버구분(1=모의, 0=실): {self.server_gubun}")
        else:
            self._log("ERROR", f"[LOGIN] ❌ 로그인 실패 err_code={err_code}")


        if self.login_loop.isRunning():
            self.login_loop.exit()


    # -----------------------
    # Conditions
    # -----------------------
    def load_conditions(self):
        ret = self.dynamicCall("GetConditionLoad()")
        if ret == 0:
            self._log("ERROR", "[COND] GetConditionLoad() 실패")
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
            except:
                continue


        buy_idx = {self.cond_name_to_idx.get(n) for n in BUY_COND_NAMES if n in self.cond_name_to_idx}
        sell_idx = {self.cond_name_to_idx.get(n) for n in SELL_COND_NAMES if n in self.cond_name_to_idx}
        buy_idx.discard(None)
        sell_idx.discard(None)


        self._log("INFO", f"[COND] BUY_CONDITIONS={BUY_COND_NAMES} -> idx={buy_idx}")
        self._log("INFO", f"[COND] SELL_CONDITIONS={SELL_COND_NAMES} -> idx={sell_idx}")


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
            self._log(
                "INFO",
                f"[COND-SUB] 구독 시도: {name} idx={idx} scr={scr} ret={ret} (try {i}/5)",
                key=f"COND_SUB_{name}",
                throttle=0.2,
            )
            if ret == 1:
                self._log("INFO", f"[COND-SUB] ✅ 구독 성공: {name}")
                return True
            time.sleep(0.2)
        self._log("ERROR", f"[COND-SUB] ❌ 구독 실패: {name}")
        return False


    def _on_receive_tr_condition(self, scr_no, code_list, cond_name, cond_index, next_):
        codes = [c for c in str(code_list).split(";") if c]
        cond_name = str(cond_name)


        self._log("INFO", f"[TRCOND] scr={scr_no} cond={cond_name} idx={cond_index} next={next_}")
        self._log("INFO", f"[TRCOND] {self._fmt_codelist(codes)}", key=f"TRCOND_{cond_name}", throttle=0.5)


        if cond_name in SELL_COND_NAMES:
            self.sell_cond_members[cond_name] = set(map(_norm_code, codes))
            self._log(
                "INFO",
                f"[SELL-COND-INIT] cond={cond_name} -> {self._fmt_codelist(self.sell_cond_members[cond_name])}",
                key=f"SELLINIT_{cond_name}",
                throttle=0.5,
            )


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
            self._log(
                "INFO",
                f"[COND-REAL] cond={cond_name} 편입(I): {self._code_tag(code)}",
                key=f"COND_I_{cond_name}_{code}",
                throttle=0.5,
            )
            if cond_name in SELL_COND_NAMES:
                self.sell_cond_members[cond_name].add(code)
            if cond_name in BUY_COND_NAMES:
                self._enqueue_buy(code, cond_name, ev="REAL_I")


        elif event_type == "D":
            self._log(
                "INFO",
                f"[COND-REAL] cond={cond_name} 이탈(D): {self._code_tag(code)}",
                key=f"COND_D_{cond_name}_{code}",
                throttle=0.5,
            )
            if cond_name in SELL_COND_NAMES:
                self.sell_cond_members[cond_name].discard(code)
                self._schedule_delayed_sell(code, cond_name, reason="COND_EXIT")


    # -----------------------
    # TR
    # -----------------------
    def _on_receive_tr_data(self, scr_no, rq_name, tr_code, record_name, prev_next, data_len, err_code, msg1, msg2):
        rq_name = str(rq_name)


        if rq_name == "opt10001_req":
            code = _norm_code(
                self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "종목코드")
            )
            price_raw = self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "현재가")
            price = abs(_safe_int(price_raw, 0))
            self._price_resp[code] = price


            self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()


        elif rq_name == "opw00018_req":
            self._parse_balance(tr_code, rq_name)


            self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()


        elif rq_name == "opt10075_req":
            self._parse_unfilled(tr_code, rq_name)


            self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()
        else:
            if self._tr_busy:
                self._tr_release()
            if self.tr_loop.isRunning():
                self.tr_loop.exit()


    def _on_receive_msg(self, scr_no, rq_name, tr_code, msg):
        self._log("DEBUG", f"[MSG] scr={scr_no} rq={rq_name} tr={tr_code} msg={msg}", key="MSG", throttle=1.0)


    def request_price(self, code: str) -> Optional[int]:
        code = _norm_code(code)
        if not code:
            return None


        now = time.time()
        gap = now - self._last_price_req_ts
        if gap < PRICE_REQ_INTERVAL:
            time.sleep(PRICE_REQ_INTERVAL - gap)
        self._last_price_req_ts = time.time()


        self._price_resp.pop(code, None)
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)


        for attempt in range(1, PRICE_RETRY_MAX + 1):
            ok = self._tr_request("opt10001_req", "opt10001", SCREEN_TR_PRICE, timeout_sec=6.0)
            if ok:
                price = self._price_resp.get(code, 0)
                if price > 0:
                    return price
            else:
                self._log(
                    "WARN",
                    f"[PRICE] TR요청 스킵/실패 attempt={attempt} code={self._code_tag(code)}",
                    key=f"PRICE_FAIL_{code}",
                    throttle=0.5,
                )
            time.sleep(PRICE_RETRY_SLEEP)


        return None


    def request_balance(self):
        now = time.time()
        if now - self._last_balance_req_ts < BALANCE_COOLDOWN_SEC:
            return
        self._last_balance_req_ts = now


        self._log("DEBUG", "[BALANCE] 요청: opw00018 잔고조회", key="BAL_REQ", throttle=0.5)


        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)


        # ✅ (교체) 아래 코드 방식: PASSWD 사용(빈값 가능)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", PASSWD)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", PASSWD_MEDIA)


        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self._tr_request("opw00018_req", "opw00018", SCREEN_TR_BAL, timeout_sec=8.0)


    def request_unfilled(self):
        self._log("DEBUG", "[UNFILLED] 요청: opt10075 미체결조회", key="UNF_REQ", throttle=0.5)


        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
        self.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self._tr_request("opt10075_req", "opt10075", SCREEN_TR_UNFILLED, timeout_sec=8.0)


    # -----------------------
    # ✅ 잔고/미체결 "변화 있을 때만" 로그를 위한 signature
    # -----------------------
    def _make_balance_sig(self, qty_map: Dict[str, int], avg_map: Dict[str, int]) -> Tuple[Tuple[str, int, int], ...]:
        rows = []
        for c in sorted(qty_map.keys()):
            rows.append((c, int(qty_map.get(c, 0)), int(avg_map.get(c, 0))))
        return tuple(rows)


    def _make_unfilled_sig(self, unfilled: Dict[str, Dict[str, Any]]) -> Tuple[Tuple[str, str, str, int, int, int], ...]:
        rows = []
        for order_no in sorted(unfilled.keys()):
            od = unfilled.get(order_no, {}) or {}
            rows.append((
                str(order_no),
                _norm_code(od.get("code", "")),
                str(od.get("bs", "")),
                int(od.get("qty", 0) or 0),
                int(od.get("unfilled", 0) or 0),
                int(od.get("price", 0) or 0),
            ))
        return tuple(rows)


    def _parse_balance(self, tr_code, rq_name):
        cnt = _safe_int(self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name), 0)


        new_qty = defaultdict(int)
        new_avg = defaultdict(int)
        new_name = defaultdict(str)


        for i in range(cnt):
            code = _norm_code(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목번호"))
            name = str(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목명")).strip()
            qty = _safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "보유수량"), 0)
            avg = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "평균단가"), 0))
            if code and qty > 0:
                new_qty[code] = qty
                new_avg[code] = avg
                new_name[code] = name
                if name:
                    self.code_name_cache[code] = name


        new_sig = self._make_balance_sig(new_qty, new_avg)
        changed = (new_sig != self._last_balance_sig)
        self._last_balance_sig = new_sig


        self.holdings_qty = new_qty
        self.holdings_avg = new_avg
        self.holdings_name = new_name


        for code in list(self.rebuy_block_pending):
            if self.holdings_qty.get(code, 0) <= 0:
                self.rebuy_block_pending.discard(code)


        if not changed:
            return


        self._log(
            "INFO",
            f"[BALANCE] 변경감지: 보유종목({len(self.holdings_qty)}): "
            + "{%s}"
            % ", ".join(
                [
                    f"{self._code_tag(c)}={q}"
                    for c, q in list(self.holdings_qty.items())[:12]
                ]
            )
            + (" ..." if len(self.holdings_qty) > 12 else ""),
            key="BAL_CHANGED",
            throttle=0.0,
        )


    def _parse_unfilled(self, tr_code, rq_name):
        cnt = _safe_int(self.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name), 0)
        new_unfilled = {}
        for i in range(cnt):
            order_no = str(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "주문번호")).strip()
            code = _norm_code(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목코드"))
            bs = str(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "매매구분")).strip()
            qty = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "주문수량"), 0))
            unfilled = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "미체결수량"), 0))
            price = abs(_safe_int(self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "주문가격"), 0))
            if order_no and code and unfilled > 0:
                new_unfilled[order_no] = {"code": code, "bs": bs, "qty": qty, "unfilled": unfilled, "price": price}


        new_sig = self._make_unfilled_sig(new_unfilled)
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
        for order_no in list(sorted(self.unfilled_orders.keys()))[:8]:
            od = self.unfilled_orders[order_no]
            sample.append(
                f"{order_no}:{self._code_tag(od.get('code'))}/{od.get('bs')}/미체결{od.get('unfilled')}"
            )
        sample_txt = ", ".join(sample)
        more = "" if len(self.unfilled_orders) <= 8 else f" ... (+{len(self.unfilled_orders) - 8})"
        self._log(
            "INFO",
            f"[UNFILLED] 변경감지: 미체결 {len(self.unfilled_orders)}건 | {sample_txt}{more}",
            key="UNF_CHANGED",
            throttle=0.0,
        )


    # -----------------------
    # Chejan
    # -----------------------
    def _on_receive_chejan_data(self, gubun, item_cnt, fid_list):
        gubun = str(gubun)
        self._log("DEBUG", f"[CHEJAN] gubun={gubun} item_cnt={item_cnt}", key=f"CHEJAN_{gubun}", throttle=1.0)


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
                    self._last_unfilled_sig = self._make_unfilled_sig(self.unfilled_orders)
                except:
                    pass


        elif gubun == "1":
            code = _norm_code(self.dynamicCall("GetChejanData(int)", 9001))
            name = str(self.dynamicCall("GetChejanData(int)", 302)).strip()
            qty = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 930), 0))
            avg = abs(_safe_int(self.dynamicCall("GetChejanData(int)", 931), 0))


            if code:
                if qty > 0:
                    self.holdings_qty[code] = qty
                    self.holdings_avg[code] = avg
                    self.holdings_name[code] = name
                    if name:
                        self.code_name_cache[code] = name
                else:
                    self.holdings_qty.pop(code, None)
                    self.holdings_avg.pop(code, None)
                    self.holdings_name.pop(code, None)
                    self.rebuy_block_pending.discard(code)


                try:
                    self._last_balance_sig = self._make_balance_sig(self.holdings_qty, self.holdings_avg)
                except:
                    pass


    # -----------------------
    # Exposure / Limits
    # -----------------------
    def _estimate_exposure_code(self, code: str, price_hint: int = 0) -> int:
        code = _norm_code(code)
        hold_qty = self.holdings_qty.get(code, 0)
        hold_avg = self.holdings_avg.get(code, 0)
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


        return int(hold_value + pending_buy_value + queued_budget)


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


    # -----------------------
    # Buy Queue
    # -----------------------
    def _enqueue_buy(self, code: str, cond_name: str, ev: str = ""):
        code = _norm_code(code)
        if not code:
            return
        if code in self.buy_queue_set:
            return
        if code in self.sold_today:
            self._log("INFO", f"[BUY-QUEUE] 스킵: 재매수 금지 -> {self._code_tag(code)}", key=f"BUYQ_SOLD_{code}", throttle=1.0)
            return
        self.buy_queue.append((code, cond_name, time.time()))
        self.buy_queue_set.add(code)
        self._log("INFO", f"[BUY-QUEUE] 추가: {self._code_tag(code)} cond={cond_name} queue_len={len(self.buy_queue)}",
                    key="BUYQ_ADD", throttle=0.2)


    def _send_order(self, name: str, screen: str, acc: str, order_type: int,
                    code: str, qty: int, price: int, hoga: str, org_order_no: str) -> int:
        sig = "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)"
        args = [name, screen, acc, int(order_type), code, int(qty), int(price), hoga, org_order_no]
        return self.dynamicCall(sig, args)


    def _process_buy_queue_tick(self):
        if not self.buy_queue:
            return


        code, cond, ts = self.buy_queue.popleft()
        self.buy_queue_set.discard(code)


        self._log("INFO", f"[BUY] 진입: target={self._code_tag(code)} budget={TARGET_BUY_AMOUNT} cond={cond}",
                    key=f"BUY_ENTER_{code}", throttle=0.2)


        price = self.request_price(code)
        if not price:
            self._log("WARN", f"[BUY] 현재가 실패 -> 스킵: {self._code_tag(code)}", key=f"BUY_NOPRICE_{code}", throttle=1.0)
            return


        ok, reason = self._can_buy_code(code, price)
        if not ok:
            self._log("INFO", f"[BUY] 스킵: {self._code_tag(code)} reason={reason}", key=f"BUY_BLOCK_{code}", throttle=0.5)
            return


        qty = TARGET_BUY_AMOUNT // price
        if qty <= 0:
            return


        order_amount = qty * price
        self._log("INFO", f"[BUY] 계산: {self._code_tag(code)} price={price} qty={qty} order_amount={order_amount}",
                    key=f"BUY_CALC_{code}", throttle=0.2)


        ret = self._send_order("BUY", "0101", self.account, 1, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[BUY] ✅ 주문성공: {self._code_tag(code)} qty={qty}", key=f"BUY_OK_{code}", throttle=0.2)
            self.pending_codes.add(code)
        else:
            self._log("WARN", f"[BUY] ❌ 주문실패: {self._code_tag(code)} ret={ret}", key=f"BUY_FAIL_{code}", throttle=1.0)


    # -----------------------
    # Delayed Sell
    # -----------------------
    def _schedule_delayed_sell(self, code: str, cond: str, reason: str):
        code = _norm_code(code)
        if not code:
            return
        due = time.time() + float(SELL_DELAY_SEC)
        self.delayed_sells[(code, cond)] = due
        self._log("INFO", f"[SELL-DELAY] 예약: {self._code_tag(code)} cond={cond} after={SELL_DELAY_SEC}s (reason={reason})",
                    key=f"SELL_SCHED_{code}_{cond}", throttle=0.2)


    def _process_delayed_sells_tick(self):
        now = time.time()
        keys = list(self.delayed_sells.keys())
        for (code, cond) in keys:
            due = self.delayed_sells.get((code, cond), 0)
            if now < due:
                continue


            if cond in SELL_COND_NAMES and code in self.sell_cond_members.get(cond, set()):
                self.delayed_sells.pop((code, cond), None)
                continue


            self.delayed_sells.pop((code, cond), None)


            hold_qty = self.holdings_qty.get(code, 0)
            if hold_qty <= 0:
                continue
            if code in self.pending_codes:
                continue


            self._request_sell_all(code, reason=f"EXIT_{cond}_DELAY{SELL_DELAY_SEC}s")


    def _request_sell_all(self, code: str, reason: str):
        code = _norm_code(code)
        qty = self.holdings_qty.get(code, 0)
        if qty <= 0:
            return


        self._log("INFO", f"[SELL] 주문전송(전량): {self._code_tag(code)} qty={qty} 시장가 reason={reason}",
                    key=f"SELL_SEND_{code}", throttle=0.2)


        ret = self._send_order("SELL", "0102", self.account, 2, code, qty, 0, "03", "")
        if ret == 0:
            self._log("INFO", f"[SELL] ✅ 주문성공: {self._code_tag(code)} qty={qty} reason={reason}",
                        key=f"SELL_OK_{code}", throttle=0.2)
            self.pending_codes.add(code)
            self._set_rebuy_block(code, cooldown_sec=REBUY_COOLDOWN_SEC)
        else:
            self._log("WARN", f"[SELL] ❌ 주문실패: {self._code_tag(code)} ret={ret}", key=f"SELL_FAIL_{code}", throttle=1.0)


    # -----------------------
    # Rebuy block (persist)
    # -----------------------
    def _set_rebuy_block(self, code: str, cooldown_sec: float):
        code = _norm_code(code)
        if not code:
            return
        self.rebuy_block_until[code] = time.time() + float(cooldown_sec)
        self.rebuy_block_pending.add(code)


    def _promote_sold_today_if_confirmed(self):
        for code, until in list(self.rebuy_block_until.items()):
            if self.holdings_qty.get(code, 0) <= 0 and code not in self.rebuy_block_pending:
                if code not in self.sold_today:
                    self.sold_today.add(code)
                    if SOLD_TODAY_PERSIST:
                        self._save_sold_today()


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


    # -----------------------
    # Orphan Sweeper
    # -----------------------
    def _is_in_any_sell_condition(self, code: str) -> bool:
        code = _norm_code(code)
        for s in self.sell_cond_members.values():
            if code in s:
                return True
        return False


    def _orphan_sweeper_tick(self):
        try:
            now = time.time()
            if not self.holdings_qty:
                return


            for code, qty in list(self.holdings_qty.items()):
                if qty <= 0:
                    self.orphan_first_seen.pop(code, None)
                    continue
                if code in self.pending_codes:
                    continue
                if self._is_in_any_sell_condition(code):
                    self.orphan_first_seen.pop(code, None)
                    continue


                first = self.orphan_first_seen.get(code)
                if first is None:
                    self.orphan_first_seen[code] = now
                    continue
                if (now - first) >= float(ORPHAN_GRACE_SEC):
                    self._request_sell_all(code, reason="ORPHAN_SWEEPER")
                    self.orphan_first_seen[code] = now + 999999


        except Exception as e:
            self._log("ERROR", f"[ORPHAN] sweeper error: {e}", key="ORPHAN_ERR", throttle=5.0)


    # -----------------------
    # Periodic Sync
    # -----------------------
    def _periodic_sync_tick(self):
        self.request_balance()
        self.request_unfilled()
        self._promote_sold_today_if_confirmed()


    # -----------------------
    # Run
    # -----------------------
    def start(self):
        self._log("INFO", "[INIT] 프로그램 초기화 시작")
        self._log("INFO", "[INIT] 프로그램 초기화 완료")


        self.comm_connect()
        if not self.account:
            self._log("ERROR", "[FATAL] 계좌번호 없음. 종료.")
            return


        # ✅ (교체) 아래 코드 방식: 로그인 직후 ShowAccountWindow 호출
        # - PASSWD="" 인 경우 특히 유용
        self.show_account_window()
        time.sleep(1.5)


        self.load_conditions()
        self.subscribe_conditions()


        self.request_balance()
        self.request_unfilled()


        self.buy_timer.start(250)
        self.delay_sell_timer.start(200)
        self.orphan_timer.start(int(ORPHAN_CHECK_INTERVAL_SEC * 1000))
        self.sync_timer.start(int(max(5, BALANCE_COOLDOWN_SEC) * 1000))


        self._log("INFO", "[RUN] 타이머 시작 완료")




def main():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.start()
    sys.exit(app.exec_())




if __name__ == "__main__":
    main()