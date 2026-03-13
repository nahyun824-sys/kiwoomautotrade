# -*- coding: utf-8 -*-
\"\"\"Kiwoom OpenAPI client + trading logic (split from single-file script).\"\"\"

import time
import datetime
from collections import deque, defaultdict
from typing import Optional, Dict, Any, Set

from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer

from . import config
from .report import write_daily_trade_report


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        # 이벤트 연결
        self.OnEventConnect.connect(self._on_event_connect)
        self.OnReceiveConditionVer.connect(self._on_receive_condition_ver)
        self.OnReceiveRealCondition.connect(self._on_receive_real_condition)
        self.OnReceiveTrCondition.connect(self._on_receive_tr_condition)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)
        self.OnReceiveChejanData.connect(self._on_receive_chejan_data)

        # 동기 대기를 위한 루프 (TR별로 분리)
        self.login_loop = QEventLoop()
        self.balance_loop = QEventLoop()
        self.price_loop = QEventLoop()

        # 상태
        self.account: Optional[str] = None
        self.conditions: Dict[int, str] = {}          # {idx: name}
        self.holdings: Dict[str, int] = {}            # {code: qty}

        # 기존보유 손절 판단용 손익률 맵
        self.balance_profit_rates: Dict[str, float] = {}  # {code: profit_rate_float}

        # 잔고 로딩 완료 여부 (잔고 반영 전 중복매수 방지)
        self.balance_ready: bool = False

        # pending: 주문 진행 중 코드 잠금
        self.pending_orders: Dict[str, float] = {}        # {code: ts}

        self.last_prices: Dict[str, int] = {}           # {code: price}
        self.server_gubun: Optional[str] = None

        # 조건 인덱스
        self.buy_condition_indices: Set[int] = set()
        self.sell_condition_indices: Set[int] = set()

        # 매수 큐
        self.buy_queue = deque()
        self.is_buying: bool = False

        # 종목별 누적 매수 금액
        self.code_accum_buy_amount = defaultdict(int)
        self.max_position_per_code: int = config.MAX_POSITION_PER_CODE

        # 잔고 조회 쿨타임
        self.last_balance_req_time: float = 0.0

        # 코드명 캐시
        self.code_name_cache: Dict[str, str] = {}

        # 포지션 기록/리포트용
        self.open_positions: Dict[str, Dict[str, Any]] = {}
        self.closed_trades_today = []
        self.last_report_date = None

        # 매도 조건별 현재 편입 종목 집합
        self.sell_condition_codes = defaultdict(set)

        # 최근 매수 시도 쿨다운
        self.last_buy_attempt_ts: Dict[str, float] = {}

        # 손절
        self.STOPLOSS_RATE: float = config.STOPLOSS_RATE

        # opt10001 레이트리밋 (TR 과부하 ret=-200 완화)
        self.last_price_req_ts: float = 0.0

        # 매일 16:00 리포트 체크
        self.report_timer = QTimer()
        self.report_timer.setInterval(30 * 1000)
        self.report_timer.timeout.connect(self._check_and_generate_daily_report)
        self.report_timer.start()

        # 1분마다 손절/고아 점검
        self.force_sell_timer = QTimer()
        self.force_sell_timer.setInterval(60 * 1000)
        self.force_sell_timer.timeout.connect(self._check_force_sell_after_exit)
        self.force_sell_timer.start()

        print("[INIT] 프로그램 초기화 완료")

    # -----------------------------
    # 유틸
    # -----------------------------
    @staticmethod
    def _norm_code(code: str) -> str:
        return (code or "").strip().replace("A", "")

    def get_code_name(self, code: str) -> str:
        code = self._norm_code(code)
        if not code:
            return ""
        if code in self.code_name_cache:
            return self.code_name_cache[code]
        name = self.dynamicCall("GetMasterCodeName(QString)", code)
        name = (name or "").strip()
        self.code_name_cache[code] = name
        return name

    def _expire_pending_orders(self):
        now = time.time()
        expired = [c for c, ts in self.pending_orders.items() if now - ts >= config.PENDING_EXPIRE_SEC]
        for c in expired:
            age = now - self.pending_orders.get(c, now)
            print(f"[PENDING-EXPIRE] {c} pending {age:.1f}s 경과 → 자동 해제")
            self.pending_orders.pop(c, None)

    # -----------------------------
    # 로그인
    # -----------------------------
    def login(self):
        print("[LOGIN] 로그인 요청")
        self.dynamicCall("CommConnect()")
        self.login_loop.exec_()

    def _on_event_connect(self, err_code):
        if err_code == 0:
            print("[LOGIN] 로그인 성공")
            acc_list = self.dynamicCall('GetLoginInfo(QString)', "ACCNO")
            self.account = acc_list.split(';')[0]
            print(f"[LOGIN] 계좌번호: {self.account}")
            self.server_gubun = self.dynamicCall('GetLoginInfo(QString)', "GetServerGubun")
            print(f"[LOGIN] 서버구분(1=모의, 0=실): {self.server_gubun}")

            self.request_balance()
            self.dynamicCall("GetConditionLoad()")
        else:
            print(f"[LOGIN] 로그인 실패 코드: {err_code}")
        self.login_loop.quit()

    # -----------------------------
    # 잔고 조회(opw00018)
    # -----------------------------
    def request_balance(self):
        if self.balance_loop.isRunning():
            print("[BALANCE] balance_loop 동작 중 → 잔고 요청 스킵")
            return

        now = time.time()
        if now - self.last_balance_req_time < 3.0:
            print("[BALANCE] 최근에 조회함 → 잔고 요청 스킵")
            return
        self.last_balance_req_time = now

        print("[BALANCE] 잔고 조회 요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")

        ret = self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "opw00018_req", "opw00018", 0, config.SCREEN_TR_BALANCE
        )
        if ret != 0:
            print(f"[BALANCE] TR 요청 실패 ret={ret}")
            return

        self.balance_loop.exec_()

    def _parse_balance(self, trcode, rqname):
        cnt = int(self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname))
        print(f"[BALANCE-DEBUG] 반복건수 cnt={cnt}")

        holdings = {}
        profit_rate_map = {}

        for i in range(cnt):
            code = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)",
                trcode, rqname, i, "종목번호"
            ).strip()
            code = self._norm_code(code)

            qty_str = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)",
                trcode, rqname, i, "보유수량"
            ).strip()

            rate_str = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)",
                trcode, rqname, i, "손익율"
            ).strip()
            if not rate_str:
                rate_str = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    trcode, rqname, i, "수익률"
                ).strip()

            try:
                qty = int(qty_str)
            except Exception:
                qty = 0

            try:
                rate = float(rate_str) if rate_str else 0.0
            except Exception:
                rate = 0.0

            if code and qty > 0:
                holdings[code] = qty
                profit_rate_map[code] = rate

        self.holdings = holdings
        self.balance_profit_rates = profit_rate_map
        self.balance_ready = True

        print(f"[BALANCE] 보유종목(잔고TR 기준): {self.holdings}")

    # -----------------------------
    # 조건식 로드/구독
    # -----------------------------
    def _on_receive_condition_ver(self, bRet, msg):
        if bRet != 1:
            print(f"[COND] 조건 로드 실패: {msg}")
            return

        raw = self.dynamicCall("GetConditionNameList()")
        print(f"[COND] 조건 목록: {raw}")

        conds = {}
        for item in raw.split(';'):
            if not item:
                continue
            idx, name = item.split('^')
            conds[int(idx)] = name.strip()
        self.conditions = conds

        self.buy_condition_indices = {idx for idx, name in conds.items() if name in config.BUY_CONDITIONS}
        self.sell_condition_indices = {idx for idx, name in conds.items() if name in config.SELL_CONDITIONS}

        print(f"[COND] BUY_CONDITIONS={config.BUY_CONDITIONS}, 인덱스={self.buy_condition_indices}")
        print(f"[COND] SELL_CONDITIONS={config.SELL_CONDITIONS}, 인덱스={self.sell_condition_indices}")

        to_subscribe = [(idx, name) for idx, name in conds.items()
                        if idx in self.buy_condition_indices or idx in self.sell_condition_indices]

        if not to_subscribe:
            print("[COND] 구독할 대상 조건이 없습니다.")
            return

        for idx, name in to_subscribe:
            self.dynamicCall("SendCondition(QString, QString, int, int)", config.SCREEN_CONDITION, name, idx, 1)
            print(f"[COND] 실시간 구독 시작: {name} (idx={idx})")

    def _on_receive_tr_condition(self, scr_no, code_list, cond_name, cond_index, next_):
        cond_name = (cond_name or "").strip()
        try:
            cond_index_int = int(str(cond_index).strip())
        except Exception:
            cond_index_int = None

        print(f"[TRCOND] scr_no={scr_no}, cond_name={cond_name}, cond_index={cond_index_int}, next={next_}")
        print(f"[TRCOND] code_list={code_list}")

        if not code_list or cond_index_int is None:
            return

        if cond_index_int in self.buy_condition_indices:
            for raw in code_list.split(';'):
                code = self._norm_code(raw)
                if not code:
                    continue
                already = self.code_accum_buy_amount.get(code, 0)
                if already >= self.max_position_per_code:
                    print(f"[BUY-TRIGGER] (초기검색) {code} 누적 {already}원 ≥ {self.max_position_per_code}원 → 스킵")
                    continue
                print(f"[BUY-TRIGGER] (초기검색) cond={cond_index_int}({cond_name}), code={code}")
                self.enqueue_buy(code, config.TARGET_BUY_AMOUNT, cond_index_int, cond_name)

        if cond_index_int in self.sell_condition_indices:
            for raw in code_list.split(';'):
                code = self._norm_code(raw)
                if code:
                    self.sell_condition_codes[cond_index_int].add(code)
            print(f"[SELL-COND-INIT] cond_index={cond_index_int}({cond_name}) 초기 편입 종목: {self.sell_condition_codes[cond_index_int]}")

    def _on_receive_real_condition(self, code, type, cond_name, cond_index):
        code = self._norm_code(code)
        type = (type or "").strip()
        cond_name = (cond_name or "").strip()
        try:
            cond_index_int = int(str(cond_index).strip())
        except Exception:
            cond_index_int = None

        event = "편입(I)" if type == "I" else "이탈(D)"
        print(f"[COND] {cond_name} {event}: {code}")
        print(f"[COND-DEBUG] type={repr(type)}, cond_index_int={cond_index_int}")

        self._handle_condition_event(code, type, cond_name, cond_index_int)

    def _handle_condition_event(self, code, type, cond_name, cond_index_int):
        if cond_index_int is None or not code:
            return

        # SELL 조건 관리 + 이탈 즉시 전량매도
        if cond_index_int in self.sell_condition_indices:
            if type == "I":
                self.sell_condition_codes[cond_index_int].add(code)
                print(f"[SELL-COND] 편입(I) {cond_index_int}({cond_name}) +{code}")
            elif type == "D":
                self.sell_condition_codes[cond_index_int].discard(code)
                print(f"[SELL-COND] 이탈(D) {cond_index_int}({cond_name}) -{code}")

                if code in self.pending_orders:
                    print(f"[SELL-EXIT] {code} 진행 중 주문 존재 → 실시간 매도 스킵")
                    return

                hold_qty = self.holdings.get(code, 0)
                has_pos = (code in self.open_positions)
                if hold_qty > 0 or has_pos:
                    name = self.get_code_name(code)
                    print(f"[SELL-EXIT] 이탈(D) 즉시 전량 매도: {code}({name})")
                    self.sell_all_market(code, sell_cond_index=cond_index_int, sell_cond_name=f"EXIT_{cond_name}")

        # BUY 조건 편입 시 매수 트리거
        if type == "I" and cond_index_int in self.buy_condition_indices:
            print(f"[BUY-TRIGGER] (실시간) cond={cond_index_int}({cond_name}), code={code}")

            already = self.code_accum_buy_amount.get(code, 0)
            if already >= self.max_position_per_code:
                print(f"[BUY] {code} 누적 {already}원 ≥ {self.max_position_per_code}원 → 추가 매수 금지")
                return

            if code in self.pending_orders:
                print(f"[BUY] 진행 중 주문 있어 스킵: {code}")
                return

            self.enqueue_buy(code, config.TARGET_BUY_AMOUNT, cond_index_int, cond_name)

    # -----------------------------
    # 매수 큐 관리
    # -----------------------------
    def enqueue_buy(self, code, amount, buy_cond_index=None, buy_cond_name=""):
        now = time.time()

        # 잔고 로딩 전 매수 차단
        if not self.balance_ready:
            print(f"[BUY-QUEUE] 잔고 로딩 전 → 매수 스킵: {code}")
            return

        # 이미 보유면 스킵
        if self.holdings.get(code, 0) > 0:
            print(f"[BUY-QUEUE] 계좌에 이미 보유중 → 추가매수 스킵: {code}")
            return

        # open_positions 있으면 스킵
        if code in self.open_positions:
            print(f"[BUY-QUEUE] open_positions에 존재(보유로 간주) → 스킵: {code}")
            return

        # pending 있으면 스킵
        if code in self.pending_orders:
            print(f"[BUY-QUEUE] pending 주문 진행 중 → 스킵: {code}")
            return

        # 큐 중복 방지
        if any(c == code for (c, _, _, _) in self.buy_queue):
            print(f"[BUY-QUEUE] 이미 대기열에 존재 → 스킵: {code}")
            return

        # 최근 매수 시도 쿨다운
        last_ts = self.last_buy_attempt_ts.get(code, 0)
        if now - last_ts < config.BUY_COOLDOWN_SEC:
            remain = config.BUY_COOLDOWN_SEC - (now - last_ts)
            print(f"[BUY-QUEUE] 최근 매수시도({config.BUY_COOLDOWN_SEC}s) → 스킵: {code}, 남은 {remain:.1f}s")
            return

        # 누적 한도 체크
        already = self.code_accum_buy_amount.get(code, 0)
        if already >= self.max_position_per_code:
            print(f"[BUY-QUEUE] {code} 누적 {already}원 ≥ {self.max_position_per_code}원 → 큐 추가 안 함")
            return

        self.last_buy_attempt_ts[code] = now
        self.buy_queue.append((code, amount, buy_cond_index, buy_cond_name))
        print(f"[BUY-QUEUE] 큐 추가: {code}, 현재대기={len(self.buy_queue)}, cond={buy_cond_index}({buy_cond_name})")

        if not self.is_buying:
            self._process_next_buy()

    def _process_next_buy(self):
        if self.is_buying or not self.buy_queue:
            return
        code, amount, buy_idx, buy_name = self.buy_queue.popleft()
        self.is_buying = True
        QTimer.singleShot(400, lambda: self._buy_market_amount_internal(code, amount, buy_idx, buy_name))

    # -----------------------------
    # 현재가 조회(opt10001)
    # -----------------------------
    def request_price(self, code):
        code = self._norm_code(code)
        if not code:
            return None

        if self.price_loop.isRunning():
            print(f"[PRICE] price_loop 동작 중 → 현재가 요청 스킵: {code}")
            return None

        # 레이트리밋(과부하 방지)
        now = time.time()
        gap = now - self.last_price_req_ts
        if gap < config.PRICE_REQ_INTERVAL:
            time.sleep(config.PRICE_REQ_INTERVAL - gap)
        self.last_price_req_ts = time.time()

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        ret = self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "opt10001_req", "opt10001", 0, config.SCREEN_TR_PRICE
        )
        if ret != 0:
            print(f"[PRICE] TR 요청 실패 ret={ret} code={code}")
            return None

        print(f"[PRICE] TR 요청 성공, 현재가 대기중: {code}")
        self.price_loop.exec_()
        price = self.last_prices.get(code)
        print(f"[PRICE] TR 응답 후 price={price} (code={code})")
        return price

    def _parse_price(self, trcode, rqname):
        if rqname != "opt10001_req":
            return
        code = self.dynamicCall(
            "GetCommData(QString, QString, int, QString)",
            trcode, rqname, 0, "종목코드"
        ).strip()
        code = self._norm_code(code)

        curr_str = self.dynamicCall(
            "GetCommData(QString, QString, int, QString)",
            trcode, rqname, 0, "현재가"
        ).strip()

        try:
            price = abs(int(curr_str))
        except Exception:
            price = None

        if code and price:
            self.last_prices[code] = price
            print(f"[PRICE] {code} 현재가: {price}")

    # -----------------------------
    # 매수
    # -----------------------------
    def _buy_market_amount_internal(self, code, amount, buy_cond_index=None, buy_cond_name=""):
        code = self._norm_code(code)
        print(f"[BUY] 진입: code={code}, amount={amount}, cond={buy_cond_index}({buy_cond_name})")

        # 잔고가 업데이트되기 전/중이면 안전하게 방어
        if self.holdings.get(code, 0) > 0:
            print(f"[BUY] 이미 보유중으로 확인됨 → 매수 중단: {code}")
            self.is_buying = False
            self._process_next_buy()
            return

        already = self.code_accum_buy_amount.get(code, 0)
        remaining_amount = self.max_position_per_code - already
        if remaining_amount <= 0:
            print(f"[BUY] {code} 이미 {already}원 매수 → 한도 초과로 매수 취소")
            self.is_buying = False
            self._process_next_buy()
            return

        price = self.request_price(code)
        if not price or price <= 0:
            print(f"[BUY] 현재가 조회 실패로 매수 불가: code={code}, price={price}")
            self.is_buying = False
            self._process_next_buy()
            return

        qty = remaining_amount // price
        if qty <= 0:
            print(f"[BUY] 남은 한도 {remaining_amount}원으로 수량 0: price={price}")
            self.is_buying = False
            self._process_next_buy()
            return

        order_amount = qty * price
        name = self.get_code_name(code)
        print(f"[BUY] 계산 → {code}({name}) price={price}, qty={qty}, order_amount={order_amount}")

        rqname = "buy_by_condition"
        order_type = 1
        hoga = "03"

        # pending 기록
        self.pending_orders[code] = time.time()

        print(f"[BUY] 주문 전송: {code}({name}) 수량={qty} 시장가")
        ret = self.dynamicCall(
            "SendOrder(QString,QString,QString,int,QString,int,int,QString,QString)",
            [rqname, config.SCREEN_ORDER, self.account, int(order_type), code, int(qty), 0, hoga, ""]
        )

        if ret != 0:
            print(f"[BUY] 주문 실패 ret={ret} code={code}")
            self.pending_orders.pop(code, None)
        else:
            self.code_accum_buy_amount[code] = already + order_amount
            print(f"[BUY] 주문 성공: code={code}, qty={qty}, 누적={self.code_accum_buy_amount[code]}원")
            self._record_open_position(code, qty, price, buy_cond_index, buy_cond_name)

        self.is_buying = False
        self._process_next_buy()

    # -----------------------------
    # 매도
    # -----------------------------
    def sell_all_market(self, code, sell_cond_index=None, sell_cond_name=""):
        code = self._norm_code(code)
        name = self.get_code_name(code)
        print(f"[SELL] 진입: {code}({name}), cond={sell_cond_index}({sell_cond_name})")

        hold_qty = self.holdings.get(code, 0)
        if hold_qty <= 0:
            pos = self.open_positions.get(code)
            if pos:
                try:
                    hold_qty = int(pos.get("qty", 0) or 0)
                except Exception:
                    hold_qty = 0
                print(f"[SELL] holdings 0 → open_positions qty 사용: {code} qty={hold_qty}")

        if hold_qty <= 0:
            print(f"[SELL] 보유수량 0 → 매도 스킵: {code}")
            return

        price = self.request_price(code)
        if price:
            print(f"[SELL] 현재가(참고): {code}({name}) price={price}")

        rqname = "sell_by_condition"
        order_type = 2
        hoga = "03"

        self.pending_orders[code] = time.time()

        print(f"[SELL] 주문 전송(전량): {code}({name}) qty={hold_qty} 시장가")
        ret = self.dynamicCall(
            "SendOrder(QString,QString,QString,int,QString,int,int,QString,QString)",
            [rqname, config.SCREEN_ORDER, self.account, int(order_type), code, int(hold_qty), 0, hoga, ""]
        )

        if ret != 0:
            print(f"[SELL] 주문 실패 ret={ret} code={code}")
            self.pending_orders.pop(code, None)
        else:
            print(f"[SELL] 주문 성공: {code}({name}) qty={hold_qty}")
            self._record_close_position(code, hold_qty, price, sell_cond_index, sell_cond_name)
            self.holdings.pop(code, None)
            self.balance_profit_rates.pop(code, None)
            self.code_accum_buy_amount[code] = 0

    # -----------------------------
    # TR 수신
    # -----------------------------
    def _on_receive_tr_data(self, screenNo, rqname, trcode,
                           recordName, prevNext, dataLen,
                           errCode, msg1, msg2):
        if rqname == "opw00018_req":
            try:
                self._parse_balance(trcode, rqname)
            finally:
                if self.balance_loop.isRunning():
                    self.balance_loop.quit()
        elif rqname == "opt10001_req":
            try:
                self._parse_price(trcode, rqname)
            finally:
                if self.price_loop.isRunning():
                    self.price_loop.quit()
        else:
            print(f"[TR] 수신: {rqname}")

    # -----------------------------
    # 체잔 수신
    # -----------------------------
    def _on_receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(f"[CHEJAN] gubun={gubun} item_cnt={item_cnt}")

        # 주문/체결/잔고 모두에서 코드 받아 pending 해제
        try:
            c = self.dynamicCall("GetChejanData(int)", 9001).strip()
            c = self._norm_code(c)
            if c:
                if c in self.pending_orders:
                    self.pending_orders.pop(c, None)
                    print(f"[CHEJAN-PENDING] pending 해제: {c}")
        except Exception:
            pass

        # 잔고변경(gubun == '1') → holdings 업데이트
        if str(gubun) == '1':
            try:
                code = self.dynamicCall("GetChejanData(int)", 9001).strip()
                code = self._norm_code(code)

                qty_str = self.dynamicCall("GetChejanData(int)", 930).strip()
                try:
                    qty = int(qty_str)
                except Exception:
                    qty = 0

                if code:
                    if qty > 0:
                        self.holdings[code] = qty
                    else:
                        self.holdings.pop(code, None)
                        self.balance_profit_rates.pop(code, None)

                    print(f"[CHEJAN-HOLDINGS] {code} 보유수량={qty} → holdings={self.holdings}")

                    if qty <= 0:
                        self.pending_orders.pop(code, None)
            except Exception as e:
                print(f"[CHEJAN-ERR] holdings 업데이트 실패: {e}")

        # 체잔 후 잔고 TR 재조회
        def do_balance():
            if not self.balance_loop.isRunning():
                self.request_balance()
            else:
                print("[CHEJAN] balance_loop 동작 중 → 잔고 조회 스킵")

        QTimer.singleShot(1500, do_balance)

    # -----------------------------
    # 포지션 기록/리포트
    # -----------------------------
    def _record_open_position(self, code, qty, price, buy_cond_index, buy_cond_name):
        now = datetime.datetime.now()
        name = self.get_code_name(code)
        self.open_positions[code] = {
            "code": code,
            "name": name,
            "buy_price": price,
            "qty": qty,
            "buy_dt": now,
            "buy_cond_index": buy_cond_index,
            "buy_cond_name": buy_cond_name,
        }
        print(f"[POS-OPEN] {code}({name}) 매수 기록: price={price}, qty={qty}, cond={buy_cond_index}({buy_cond_name})")

    def _record_close_position(self, code, qty, price, sell_cond_index, sell_cond_name):
        now = datetime.datetime.now()
        name = self.get_code_name(code)

        pos = self.open_positions.pop(code, None)
        if pos is None:
            print(f"[POS-CLOSE] 기존 기록 없는 포지션 매도 기록: {code}({name})")
            return

        buy_price = pos.get("buy_price", 0)
        buy_dt = pos.get("buy_dt")
        buy_cond_index = pos.get("buy_cond_index")
        buy_cond_name = pos.get("buy_cond_name")

        sell_price = price or buy_price
        sell_qty = qty

        profit = (sell_price - buy_price) * sell_qty
        profit_rate = ((sell_price - buy_price) / buy_price * 100.0) if buy_price > 0 else 0.0

        trade = {
            "code": code,
            "name": name,
            "buy_cond_index": buy_cond_index,
            "buy_cond_name": buy_cond_name,
            "sell_cond_index": sell_cond_index,
            "sell_cond_name": sell_cond_name,
            "buy_price": buy_price,
            "sell_price": sell_price,
            "qty": sell_qty,
            "profit": profit,
            "profit_rate": profit_rate,
            "buy_dt": buy_dt,
            "sell_dt": now,
            "sell_date": now.date(),
        }
        self.closed_trades_today.append(trade)

        print(f"[POS-CLOSE] {code}({name}) profit={profit}, profit%={profit_rate:.2f}, sell_cond={sell_cond_name}")

    # -----------------------------
    # 리포트
    # -----------------------------
    def _check_and_generate_daily_report(self):
        now = datetime.datetime.now()
        today = now.date()

        if self.last_report_date == today:
            return
        if now.hour < 16:
            return

        self._generate_daily_report(today)
        self.last_report_date = today

    def _generate_daily_report(self, target_date):
        trades = [t for t in self.closed_trades_today if t.get("sell_date") == target_date]
        if not trades:
            print(f"[REPORT] {target_date} 기준 매도 내역이 없어 리포트 생성 스킵")
            return

        filename = write_daily_trade_report(config.REPORT_DIR, target_date, trades)
        print(f"[REPORT] 리포트 생성 완료: {filename}")

    # -----------------------------
    # 1분마다: 손절 + 고아 종목
    # -----------------------------
    def _check_force_sell_after_exit(self):
        self._expire_pending_orders()

        holding_codes = set(self.holdings.keys())
        open_codes = set(self.open_positions.keys())
        current_codes = holding_codes | open_codes
        if not current_codes:
            return

        valid_sell_codes = set()
        for codes in self.sell_condition_codes.values():
            valid_sell_codes.update(codes)

        print("[AUTO-SELL] 점검 시작 (1) -10% 손절, (2) 매도조건 1:1 매핑")
        print(f"[AUTO-SELL-DEBUG] current_codes={current_codes}")
        print(f"[AUTO-SELL-DEBUG] valid_sell_codes={valid_sell_codes}")
        print(f"[AUTO-SELL-DEBUG] pending_codes={set(self.pending_orders.keys())}")

        # (A) 손절: open_positions(신규) + balance_profit_rates(기존보유)
        for code in list(current_codes):
            if code in self.pending_orders:
                print(f"[STOPLOSS] {code} 진행 중 주문 존재 → 스킵")
                continue

            # 신규 포지션 손절(매수가 대비)
            pos = self.open_positions.get(code)
            if pos:
                buy_price = float(pos.get("buy_price", 0) or 0)
                if buy_price > 0:
                    cur_price = self.request_price(code)
                    if cur_price:
                        pnl = ((cur_price - buy_price) / buy_price) * 100.0
                        if pnl <= self.STOPLOSS_RATE:
                            name = self.get_code_name(code)
                            print(f"[STOPLOSS] 신규포지션 -10% 이하 → 매도: {code}({name}) pnl%={pnl:.2f}")
                            self.sell_all_market(code, sell_cond_name="STOPLOSS_-10")
                            continue

            # 기존 보유 손절(opw00018 손익율 기반)
            rate = self.balance_profit_rates.get(code, None)
            if rate is not None and rate <= self.STOPLOSS_RATE:
                name = self.get_code_name(code)
                print(f"[STOPLOSS] 기존보유 손익율 -10% 이하 → 매도: {code}({name}) 손익율={rate}")
                self.sell_all_market(code, sell_cond_name="STOPLOSS_BALANCE_RATE")
                continue

        # (B) 고아 종목(매도조건 매핑 실패) 전량매도
        for code in list(current_codes):
            if code in self.pending_orders:
                print(f"[AUTO-SELL] {code} 진행 중 주문 존재 → 스킵")
                continue

            if code not in valid_sell_codes:
                qty = self.holdings.get(code, 0)
                if qty <= 0:
                    pos = self.open_positions.get(code)
                    if pos:
                        try:
                            qty = int(pos.get("qty", 0) or 0)
                        except Exception:
                            qty = 0

                name = self.get_code_name(code)
                print(f"[AUTO-SELL] 매핑 실패 → 전량 매도: {code}({name}) qty={qty}")
                self.sell_all_market(code, sell_cond_name="AUTO_SELL_UNMAPPED")

        print("[AUTO-SELL] 점검 종료")
