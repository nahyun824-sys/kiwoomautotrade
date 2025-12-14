import sys
import time
import os
import datetime
from collections import deque, defaultdict

from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer

# 화면번호(중복 사용 금지)
SCREEN_LOGIN = "0000"
SCREEN_CONDITION = "0001"          # 실시간 조건
SCREEN_TR_PRICE = "1001"           # opt10001
SCREEN_TR_BALANCE = "1002"         # opw00018
SCREEN_ORDER = "2001"

# 매매 한도
TARGET_BUY_AMOUNT = 300000       # ★ 1종목당 목표 매수금액
MAX_POSITION_PER_CODE = 300000   # ★ 1종목당 최대 보유 금액

# ★ 변경 1: 편입/이탈 조건 이름들
# w1, w2, w3, x1 편입(I) 시 매수
BUY_CONDITIONS = {"w1", "w2", "w3", "x1"}

# w, x1 이탈(D) 시 매도 조건 집합
SELL_CONDITIONS = {"w", "x1"}

REPORT_DIR = "reports"  # 리포트 저장 폴더


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
        self.balance_loop = QEventLoop()   # opw00018 용
        self.price_loop = QEventLoop()     # opt10001 용

        # 상태
        self.account = None
        self.conditions = {}             # {idx: name}
        self.holdings = {}               # {code: qty}

        # ✅ pending_orders 패치: set -> dict(timestamp)
        self.pending_orders = {}         # {code: ts}
        self.PENDING_EXPIRE_SEC = 15     # 15초 지나면 pending 자동 해제

        self.last_prices = {}            # {code: price}
        self.server_gubun = None

        # 조건 인덱스
        self.buy_condition_indices = set()
        self.sell_condition_indices = set()

        # 매수 큐 (한 번에 한 종목만 TR 처리)
        self.buy_queue = deque()         # [(code, amount, buy_cond_index, buy_cond_name), ...]
        self.is_buying = False

        # 종목별 누적 매수 금액
        self.code_accum_buy_amount = defaultdict(int)  # {code: 누적 매수 금액}
        self.max_position_per_code = MAX_POSITION_PER_CODE

        # 잔고 조회 쿨타임
        self.last_balance_req_time = 0.0

        # 코드명 캐시
        self.code_name_cache = {}

        # 매매 내역 기록용 (리포트용)
        self.open_positions = {}        # {code: {...}}
        self.closed_trades_today = []   # [{...}, ...]
        self.last_report_date = None

        # ★ 매도 조건별 현재 편입 종목 집합
        self.sell_condition_codes = defaultdict(set)   # {cond_index: {code, ...}}

        # 매일 16:00 리포트 체크용 타이머
        self.report_timer = QTimer()
        self.report_timer.setInterval(30 * 1000)  # 30초마다 체크
        self.report_timer.timeout.connect(self._check_and_generate_daily_report)
        self.report_timer.start()

        # ★ 1분마다 “매도 조건 vs 보유 종목 1:1 매핑” 체크용 타이머
        self.force_sell_timer = QTimer()
        self.force_sell_timer.setInterval(60 * 1000)  # 1분
        self.force_sell_timer.timeout.connect(self._check_force_sell_after_exit)
        self.force_sell_timer.start()

        print("[INIT] 프로그램 초기화 완료")

    # -----------------------------
    # ✅ pending 만료 처리
    # -----------------------------
    def _expire_pending_orders(self):
        now = time.time()
        expired = [c for c, ts in self.pending_orders.items() if now - ts >= self.PENDING_EXPIRE_SEC]
        for c in expired:
            age = now - self.pending_orders.get(c, now)
            print(f"[PENDING-EXPIRE] {c} pending {age:.1f}s 경과 → 자동 해제")
            self.pending_orders.pop(c, None)

    # -----------------------------
    # 유틸: 코드명 조회
    # -----------------------------
    def get_code_name(self, code):
        code = (code or "").strip().replace("A", "")
        if not code:
            return ""
        if code in self.code_name_cache:
            return self.code_name_cache[code]
        name = self.dynamicCall("GetMasterCodeName(QString)", code)
        name = (name or "").strip()
        self.code_name_cache[code] = name
        return name

    # -----------------------------
    # 로그인 및 초기 준비 (✅ 네 원본 그대로 유지)
    # -----------------------------
    def login(self):
        print("[LOGIN] 로그인 요청")
        self.dynamicCall("CommConnect()")
        self.login_loop.exec_()  # OnEventConnect에서 종료

    def _on_event_connect(self, err_code):
        if err_code == 0:
            print("[LOGIN] 로그인 성공")
            acc_list = self.dynamicCall('GetLoginInfo(QString)', "ACCNO")
            self.account = acc_list.split(';')[0]
            print(f"[LOGIN] 계좌번호: {self.account}")
            self.server_gubun = self.dynamicCall('GetLoginInfo(QString)', "GetServerGubun")
            print(f"[LOGIN] 서버구분(1=모의, 0=실): {self.server_gubun}")

            # ✅ 네 원본처럼: 로그인 성공 이벤트 안에서 잔고 조회 + 조건식 로드
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
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")  # 2: 종목별

        ret = self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "opw00018_req", "opw00018", 0, SCREEN_TR_BALANCE
        )
        if ret != 0:
            print(f"[BALANCE] TR 요청 실패 ret={ret}")
            return

        self.balance_loop.exec_()

    def _parse_balance(self, trcode, rqname):
        cnt = int(self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, "계좌평가잔고내역"))
        print(f"[BALANCE-DEBUG] 반복건수 cnt={cnt}")
        holdings = {}
        for i in range(cnt):
            code = self.dynamicCall(
                "CommGetData(QString, QString, QString, int, QString)",
                trcode, "", rqname, i, "종목코드"
            ).strip()
            qty_str = self.dynamicCall(
                "CommGetData(QString, QString, QString, int, QString)",
                trcode, "", rqname, i, "보유수량"
            ).strip()

            try:
                qty = int(qty_str)
            except Exception:
                qty = 0

            code = code.replace("A", "")
            if code and qty > 0:
                holdings[code] = qty

        if holdings:
            self.holdings = holdings
        print(f"[BALANCE] 보유종목(잔고TR 기준): {self.holdings}")

    # -----------------------------
    # 조건식 로드 및 실시간 구독
    # -----------------------------
    def _on_receive_condition_ver(self, bRet, msg):
        if bRet == 1:
            raw = self.dynamicCall("GetConditionNameList()")
            print(f"[COND] 조건 목록: {raw}")
            conds = {}
            for item in raw.split(';'):
                if not item:
                    continue
                idx, name = item.split('^')
                idx = int(idx)
                name = name.strip()
                conds[idx] = name
            self.conditions = conds

            self.buy_condition_indices = {idx for idx, name in conds.items() if name in BUY_CONDITIONS}
            self.sell_condition_indices = {idx for idx, name in conds.items() if name in SELL_CONDITIONS}

            print(f"[COND] BUY_CONDITIONS={BUY_CONDITIONS}, 인덱스={self.buy_condition_indices}")
            print(f"[COND] SELL_CONDITIONS={SELL_CONDITIONS}, 인덱스={self.sell_condition_indices}")

            to_subscribe = []
            for idx, name in conds.items():
                if idx in self.buy_condition_indices or idx in self.sell_condition_indices:
                    to_subscribe.append((idx, name))

            if not to_subscribe:
                print("[COND] 구독할 대상 조건이 없습니다.")
                return

            for idx, name in to_subscribe:
                self.dynamicCall(
                    "SendCondition(QString, QString, int, int)",
                    SCREEN_CONDITION, name, idx, 1
                )
                print(f"[COND] 실시간 구독 시작: {name} (idx={idx})")
        else:
            print(f"[COND] 조건 로드 실패: {msg}")

    # -----------------------------
    # 조건 검색 결과 (초기 리스트)
    # -----------------------------
    def _on_receive_tr_condition(self, scr_no, code_list, cond_name, cond_index, next_):
        cond_name = (cond_name or "").strip()
        cond_index_str = (str(cond_index).strip()) if cond_index is not None else ""
        try:
            cond_index_int = int(cond_index_str) if cond_index_str != "" else None
        except ValueError:
            cond_index_int = None

        print(f"[TRCOND] scr_no={scr_no}, cond_name={cond_name}, cond_index={cond_index_int}, next={next_}")
        print(f"[TRCOND] code_list={code_list}")

        if not code_list:
            return

        if cond_index_int in self.buy_condition_indices:
            for raw in code_list.split(';'):
                code = (raw or "").strip()
                if not code:
                    continue

                already = self.code_accum_buy_amount.get(code, 0)
                if already >= self.max_position_per_code:
                    print(f"[BUY-TRIGGER] (초기검색) {code} 이미 누적 {already}원 ≥ {self.max_position_per_code}원 → 스킵")
                    continue

                print(f"[BUY-TRIGGER] (초기검색) 조건 편입 매수 트리거: cond_index={cond_index_int}({cond_name}), code={code}")
                self.enqueue_buy(code, TARGET_BUY_AMOUNT, cond_index_int, cond_name)

        if cond_index_int in self.sell_condition_indices:
            for raw in code_list.split(';'):
                code = (raw or "").strip()
                if not code:
                    continue
                self.sell_condition_codes[cond_index_int].add(code)
            print(f"[SELL-COND-INIT] cond_index={cond_index_int}({cond_name}) 초기 편입 종목: {self.sell_condition_codes[cond_index_int]}")

    # -----------------------------
    # 실시간 조건 편입/이탈
    # -----------------------------
    def _on_receive_real_condition(self, code, type, cond_name, cond_index):
        code = (code or "").strip()
        type = (type or "").strip()
        cond_name = (cond_name or "").strip()
        cond_index_str = (str(cond_index).strip()) if cond_index is not None else ""
        try:
            cond_index_int = int(cond_index_str) if cond_index_str != "" else None
        except ValueError:
            cond_index_int = None

        event = "편입(I)" if type == 'I' else "이탈(D)"
        print(f"[COND] {cond_name} {event}: {code}")
        print(
            f"[COND-DEBUG] type={repr(type)}, cond_name={repr(cond_name)}, "
            f"cond_index_str={repr(cond_index_str)}, cond_index_int={cond_index_int}"
        )

        self._handle_condition_event(code, type, cond_name, cond_index_int)

    def _handle_condition_event(self, code, type, cond_name, cond_index_int):
        if cond_index_int is None:
            print(f"[COND-HANDLE] cond_index_int is None → 무시 (code={code})")
            return

        if cond_index_int in self.sell_condition_indices:
            if type == 'I':
                self.sell_condition_codes[cond_index_int].add(code)
                print(f"[SELL-COND] 편입(I) cond_index={cond_index_int}({cond_name}) 코드 추가: {code}")
            elif type == 'D':
                if code in self.sell_condition_codes.get(cond_index_int, set()):
                    self.sell_condition_codes[cond_index_int].discard(code)
                    print(f"[SELL-COND] 이탈(D) cond_index={cond_index_int}({cond_name}) 코드 제거: {code}")

        if type == 'I' and cond_index_int in self.buy_condition_indices:
            print(f"[BUY-TRIGGER] (실시간) 조건 편입 매수 트리거: cond_index={cond_index_int}({cond_name}), code={code}")

            already = self.code_accum_buy_amount.get(code, 0)
            if already >= self.max_position_per_code:
                print(f"[BUY] {code} 누적 매수금액 {already}원 ≥ {self.max_position_per_code}원 → 추가 매수 금지")
                return

            if code in self.pending_orders:
                print(f"[BUY] 진행 중 주문 있어 스킵: {code}")
                return

            self.enqueue_buy(code, TARGET_BUY_AMOUNT, cond_index_int, cond_name)
            return

        # 실시간 이탈(D) 즉시 매도 안함 (1분마다 강제매도만)

    # -----------------------------
    # 매수 큐 관리
    # -----------------------------
    def enqueue_buy(self, code, amount, buy_cond_index=None, buy_cond_name=""):
        if code in self.holdings and self.holdings.get(code, 0) > 0:
            print(f"[BUY-QUEUE] 계좌에 이미 보유중 → 추가매수 스킵: {code}")
            return

        if any(c == code for (c, _, _, _) in self.buy_queue) or code in self.pending_orders:
            print(f"[BUY-QUEUE] 이미 대기열 또는 진행중: {code}")
            return

        already = self.code_accum_buy_amount.get(code, 0)
        if already >= self.max_position_per_code:
            print(f"[BUY-QUEUE] {code} 누적 {already}원 ≥ {self.max_position_per_code}원 → 큐 추가 안 함")
            return

        self.buy_queue.append((code, amount, buy_cond_index, buy_cond_name))
        print(f"[BUY-QUEUE] 큐 추가: {code}, 현재대기={len(self.buy_queue)}, buy_cond_index={buy_cond_index}({buy_cond_name})")
        if not self.is_buying:
            self._process_next_buy()

    def _process_next_buy(self):
        if self.is_buying:
            return
        if not self.buy_queue:
            return

        code, amount, buy_idx, buy_name = self.buy_queue.popleft()
        self.is_buying = True
        QTimer.singleShot(400, lambda c=code, a=amount, bi=buy_idx, bn=buy_name: self._buy_market_amount_internal(c, a, bi, bn))

    # -----------------------------
    # 현재가 조회(opt10001)
    # -----------------------------
    def request_price(self, code):
        if self.price_loop.isRunning():
            print(f"[PRICE] price_loop 동작 중 → 현재가 요청 스킵: {code}")
            return None

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        ret = self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "opt10001_req", "opt10001", 0, SCREEN_TR_PRICE
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
        curr_str = self.dynamicCall(
            "GetCommData(QString, QString, int, QString)",
            trcode, rqname, 0, "현재가"
        ).strip()

        code = code.replace("A", "")
        try:
            price = abs(int(curr_str))
        except Exception:
            price = None

        if code and price:
            self.last_prices[code] = price
            print(f"[PRICE] {code} 현재가: {price}")

    # -----------------------------
    # 내부 매수 처리 + 포지션 기록
    # -----------------------------
    def _buy_market_amount_internal(self, code, amount, buy_cond_index=None, buy_cond_name=""):
        print(f"[BUY] _buy_market_amount_internal 진입: code={code}, amount={amount}, cond={buy_cond_index}({buy_cond_name})")

        already = self.code_accum_buy_amount.get(code, 0)
        remaining_amount = self.max_position_per_code - already
        if remaining_amount <= 0:
            print(f"[BUY] {code} 이미 {already}원 매수 → 한도 {self.max_position_per_code}원 초과, 매수 취소")
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
            print(f"[BUY] 남은 한도 {remaining_amount}원으로 매수 가능한 수량이 0: price={price}, already={already}")
            self.is_buying = False
            self._process_next_buy()
            return

        order_amount = qty * price
        name = self.get_code_name(code)
        print(f"[BUY] 계산 결과 → code={code}({name}), price={price}, qty={qty}, order_amount={order_amount}")

        rqname = "buy_by_condition"
        order_type = 1
        hoga = "03"

        # ✅ pending 기록
        self.pending_orders[code] = time.time()

        print(f"[BUY] 주문 전송: {code}({name}) 수량={qty} 시장가 (price={price})")
        ret = self.dynamicCall(
            "SendOrder(QString,QString,QString,int,QString,int,int,QString,QString)",
            [rqname, SCREEN_ORDER, self.account, int(order_type), code, int(qty), 0, hoga, ""]
        )

        if ret != 0:
            print(f"[BUY] 주문 실패 ret={ret} code={code}")
            self.pending_orders.pop(code, None)
        else:
            self.code_accum_buy_amount[code] = already + order_amount
            print(f"[BUY] 주문 전송 성공: code={code}, qty={qty}, 종목누적={self.code_accum_buy_amount[code]}원")
            self._record_open_position(code, qty, price, buy_cond_index, buy_cond_name)

        self.is_buying = False
        self._process_next_buy()

    # -----------------------------
    # 전량 시장가 매도 + 포지션 정리
    # -----------------------------
    def sell_all_market(self, code, sell_cond_index=None, sell_cond_name=""):
        name = self.get_code_name(code)
        print(f"[SELL] sell_all_market 진입: {code}({name}), cond={sell_cond_index}({sell_cond_name})")

        hold_qty = self.holdings.get(code, 0)
        if hold_qty <= 0:
            pos = self.open_positions.get(code)
            if pos is not None:
                try:
                    hold_qty = int(pos.get("qty", 0) or 0)
                except Exception:
                    hold_qty = 0
                print(f"[SELL] holdings에는 없지만 open_positions 기준 보유수량 사용: {code} qty={hold_qty}")

        if hold_qty <= 0:
            print(f"[SELL] 보유수량 0 → 매도 스킵: {code}({name})")
            return

        price = self.request_price(code)
        if price:
            print(f"[SELL] {code}({name}) 현재가(참고용): {price}")

        rqname = "sell_by_condition"
        order_type = 2
        hoga = "03"

        # ✅ pending 기록
        self.pending_orders[code] = time.time()

        print(f"[SELL] 주문 전송(전량): {code}({name}) 수량={hold_qty} 시장가")
        ret = self.dynamicCall(
            "SendOrder(QString,QString,QString,int,QString,int,int,QString,QString)",
            [rqname, SCREEN_ORDER, self.account, int(order_type), code, int(hold_qty), 0, hoga, ""]
        )

        if ret != 0:
            print(f"[SELL] 주문 실패 ret={ret} code={code}")
            self.pending_orders.pop(code, None)
        else:
            print(f"[SELL] 주문 전송 성공: code={code}({name}), qty={hold_qty}")
            self._record_close_position(code, hold_qty, price, sell_cond_index, sell_cond_name)
            self.holdings.pop(code, None)
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
    # 체잔 수신(주문/체결/잔고 반영)
    # -----------------------------
    def _on_receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(f"[CHEJAN] gubun={gubun} item_cnt={item_cnt}")

        # ✅ 체잔 오면 pending 즉시 해제 (체결/잔고/주문 어느 gubun 이든)
        try:
            c = self.dynamicCall("GetChejanData(int)", 9001).strip().replace("A", "")
            if c:
                self.pending_orders.pop(c, None)
        except Exception:
            pass

        # 잔고변경(gubun == '1')일 때 holdings 직접 업데이트
        if gubun == '1':
            try:
                code = self.dynamicCall("GetChejanData(int)", 9001).strip().replace("A", "")
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
                    print(f"[CHEJAN-HOLDINGS] {code} 보유수량={qty} → holdings={self.holdings}")

                    if qty <= 0:
                        self.pending_orders.pop(code, None)
            except Exception as e:
                print(f"[CHEJAN-ERR] holdings 업데이트 실패: {e}")

        def do_balance():
            if not self.balance_loop.isRunning():
                self.request_balance()
            else:
                print("[CHEJAN] balance_loop 동작 중 → 잔고 조회 스킵")

            # ❌ 기존 pending_orders.clear() 제거 (이게 문제 유발 가능)
            # self.pending_orders.clear()

        QTimer.singleShot(1500, do_balance)

    # -----------------------------
    # 포지션 기록/리포트용 로직
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

        buy_date = buy_dt.date() if isinstance(buy_dt, datetime.datetime) else now.date()
        sell_date = now.date()

        buy_weekday = self._weekday_kr(buy_date.weekday())
        sell_weekday = self._weekday_kr(sell_date.weekday())
        trading_days = self._business_days_between(buy_date, sell_date)

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
            "buy_date": buy_date,
            "sell_date": sell_date,
            "buy_weekday": buy_weekday,
            "sell_weekday": sell_weekday,
            "trading_days": trading_days,
        }
        self.closed_trades_today.append(trade)

        print(
            f"[POS-CLOSE] {code}({name}) 매도 기록:"
            f" buy={buy_price}, sell={sell_price}, qty={sell_qty},"
            f" profit={profit}, profit%={profit_rate:.2f},"
            f" buy_cond={buy_cond_index}({buy_cond_name}),"
            f" sell_cond={sell_cond_index}({sell_cond_name}),"
            f" holding_days={trading_days}"
        )

    @staticmethod
    def _weekday_kr(idx):
        mapping = ["월", "화", "수", "목", "금", "토", "일"]
        return mapping[idx] if 0 <= idx < 7 else ""

    @staticmethod
    def _business_days_between(start_date, end_date):
        if end_date < start_date:
            start_date, end_date = end_date, start_date
        days = 0
        cur = start_date
        while cur <= end_date:
            if cur.weekday() < 5:
                days += 1
            cur += datetime.timedelta(days=1)
        return days

    # -----------------------------
    # 매일 16:00 리포트 자동 생성 (원본 유지)
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

        if not os.path.exists(REPORT_DIR):
            os.makedirs(REPORT_DIR, exist_ok=True)

        filename = os.path.join(REPORT_DIR, f"trade_report_{target_date.strftime('%Y%m%d')}.txt")
        print(f"[REPORT] 리포트 생성: {filename}")

        stats_by_sell = {}
        stats_by_buy = {}

        for t in trades:
            profit = t["profit"]
            profit_rate = t["profit_rate"]

            s_idx = t["sell_cond_index"]
            s_name = t["sell_cond_name"]
            if s_idx is not None:
                d = stats_by_sell.setdefault(s_idx, {"name": s_name, "profits": [], "profit_rates": []})
                d["profits"].append(profit)
                d["profit_rates"].append(profit_rate)

            b_idx = t["buy_cond_index"]
            b_name = t["buy_cond_name"]
            if b_idx is not None:
                d2 = stats_by_buy.setdefault(b_idx, {"name": b_name, "profits": [], "profit_rates": []})
                d2["profits"].append(profit)
                d2["profit_rates"].append(profit_rate)

        def _stat_line(values):
            if not values:
                return "건수: 0"
            cnt = len(values)
            mn = min(values)
            mx = max(values)
            avg = sum(values) / cnt
            return f"건수: {cnt}, 최소: {mn:.2f}, 최대: {mx:.2f}, 평균: {avg:.2f}"

        with open(filename, "w", encoding="utf-8") as f:
            f.write(f"=== {target_date.strftime('%Y-%m-%d')} 매매 리포트 ===\n\n")

            f.write("4-1. 개별 매도 내역\n")
            f.write("--------------------------------------------------\n")
            for t in trades:
                f.write(f"종목 코드: {t['code']}\n")
                f.write(f"종목명: {t['name']}\n")
                f.write(f"매수 조건식 번호(index): {t['buy_cond_index']}\n")
                f.write(f"매수 조건식 이름: {t['buy_cond_name']}\n")
                f.write(f"매도 조건식 번호(index): {t['sell_cond_index']}\n")
                f.write(f"매도 조건식 이름: {t['sell_cond_name']}\n")
                f.write(f"매수가: {t['buy_price']}\n")
                f.write(f"매도가: {t['sell_price']}\n")
                f.write(f"수량: {t['qty']}\n")
                f.write(f"수익금: {t['profit']}\n")
                f.write(f"수익률: {t['profit_rate']:.2f}%\n\n")

                f.write(f"  매수 날짜: {t['buy_date']}\n")
                f.write(f"  매도 날짜: {t['sell_date']}\n")
                f.write(f"  매수 요일: {t['buy_weekday']}\n")
                f.write(f"  매도 요일: {t['sell_weekday']}\n")
                f.write(f"  매수 후 몇 거래일 만에 매도: {t['trading_days']} 거래일\n")
                f.write("--------------------------------------------------\n")

            f.write("\n4-3. 당일 매도 조건식별 통계\n")
            f.write("--------------------------------------------------\n")
            for idx, d in sorted(stats_by_sell.items(), key=lambda x: x[0]):
                f.write(f"[매도 조건식] index={idx}, 이름={d['name']}\n")
                f.write(f"  수익금 통계: {_stat_line(d['profits'])}\n")
                f.write(f"  수익률 통계(%, 소수점2): {_stat_line(d['profit_rates'])}\n")
                f.write("--------------------------------------------------\n")

            f.write("\n4-4. 당일 매도건 기준 매수 조건식별 통계\n")
            f.write("--------------------------------------------------\n")
            for idx, d in sorted(stats_by_buy.items(), key=lambda x: x[0]):
                f.write(f"[매수 조건식] index={idx}, 이름={d['name']}\n")
                f.write(f"  수익금 통계: {_stat_line(d['profits'])}\n")
                f.write(f"  수익률 통계(%, 소수점2): {_stat_line(d['profit_rates'])}\n")
                f.write("--------------------------------------------------\n")

        print(f"[REPORT] 리포트 생성 완료: {filename}")

    # -----------------------------
    # ★ 1분마다 실행되는 “고아 종목” 자동 매도 체크
    # -----------------------------
    def _check_force_sell_after_exit(self):
        # ✅ pending 만료 먼저
        self._expire_pending_orders()

        holding_codes = set(self.holdings.keys())
        open_codes = set(self.open_positions.keys())
        current_codes = holding_codes | open_codes
        if not current_codes:
            return

        valid_sell_codes = set()
        for codes in self.sell_condition_codes.values():
            valid_sell_codes.update(codes)

        print("[AUTO-SELL] 매도 조건 vs 보유 종목 1:1 매핑 점검 시작")
        print(f"[AUTO-SELL-DEBUG] current_codes(holdings+open_positions): {current_codes}")
        print(f"[AUTO-SELL-DEBUG] 매도 조건 종목 집합: {valid_sell_codes}")
        print(f"[AUTO-SELL-DEBUG] pending_codes={set(self.pending_orders.keys())}")

        for code in list(current_codes):
            if code in self.pending_orders:
                print(f"[AUTO-SELL] {code} 진행 중 주문 존재 → 스킵")
                continue

            if code not in valid_sell_codes:
                qty = self.holdings.get(code, 0)
                if qty <= 0:
                    pos = self.open_positions.get(code)
                    if pos is not None:
                        try:
                            qty = int(pos.get("qty", 0) or 0)
                        except Exception:
                            qty = 0

                name = self.get_code_name(code)
                print(f"[AUTO-SELL] 매도 조건과 1:1 매핑되지 않는 보유 종목 → 전량 매도: {code}({name}), qty={qty}")
                self.sell_all_market(code, sell_cond_index=None, sell_cond_name="AUTO_SELL_UNMAPPED")

        print("[AUTO-SELL] 1:1 매핑 점검 종료")


def main():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.login()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
