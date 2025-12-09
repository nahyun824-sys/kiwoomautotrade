import sys
import time
from collections import deque, defaultdict

from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer

# 화면번호(중복 사용 금지)
SCREEN_LOGIN = "0000"
SCREEN_CONDITION = "0001"
SCREEN_TR_PRICE = "1001"      # opt10001
SCREEN_TR_BALANCE = "1002"    # opw00018
SCREEN_ORDER = "2001"

# 매매 한도
TARGET_BUY_AMOUNT = 300000       # ★ 1종목당 목표 매수금액 = 50만원
MAX_POSITION_PER_CODE = 300000   # ★ 1종목당 최대 보유 금액 = 50만원

# ★ 변경 1: 편입/이탈 조건 이름들
# w1, w2, w3 편입(I) 시 매수
BUY_CONDITIONS = {"w1", "w2", "w3", "x1"}

# w4, w5 이탈(D) 시 매도
SELL_CONDITIONS = {"w", "x1"}


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
        self.pending_orders = set()      # 주문 진행중인 종목 {code}
        self.last_prices = {}            # {code: price}
        self.server_gubun = None

        # 조건 인덱스
        self.buy_condition_indices = set()
        self.sell_condition_indices = set()

        # 매수 큐 (한 번에 한 종목만 TR 처리)
        self.buy_queue = deque()         # [(code, amount), ...]
        self.is_buying = False

        # 종목별 누적 매수 금액 (1종목당 50만원 한도)
        self.code_accum_buy_amount = defaultdict(int)  # {code: 누적 매수 금액}
        self.max_position_per_code = MAX_POSITION_PER_CODE

        # 잔고 조회 쿨타임
        self.last_balance_req_time = 0.0

        print("[INIT] 프로그램 초기화 완료")

    # -----------------------------
    # 로그인 및 초기 준비
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

            # 잔고 먼저 조회
            self.request_balance()

            # 조건식 로드
            self.dynamicCall("GetConditionLoad()")
        else:
            print(f"[LOGIN] 로그인 실패 코드: {err_code}")
        self.login_loop.quit()

    # -----------------------------
    # 잔고 조회(opw00018)
    # -----------------------------
    def request_balance(self):
        # 잔고 루프 중복 실행 방지
        if self.balance_loop.isRunning():
            print("[BALANCE] balance_loop 동작 중 → 잔고 요청 스킵")
            return

        # 너무 자주 호출되는 것 방지 (쿨타임 3초)
        now = time.time()
        if now - self.last_balance_req_time < 3.0:
            print("[BALANCE] 최근에 조회함 → 잔고 요청 스킵")
            return
        self.last_balance_req_time = now

        print("[BALANCE] 잔고 조회 요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")  # 2: 종목별

        ret = self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "opw00018_req", "opw00018", 0, SCREEN_TR_BALANCE
        )
        if ret != 0:
            print(f"[BALANCE] TR 요청 실패 ret={ret}")
            return

        self.balance_loop.exec_()  # _on_receive_tr_data에서 종료

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

        # holdings가 비어 있으면 기존 holdings 유지
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

            # 인덱스 기반 매수/매도 조건 세트 구성
            self.buy_condition_indices = {
                idx for idx, name in conds.items() if name in BUY_CONDITIONS
            }
            # ★ 변경 2: 매도 조건도 여러 개(w4, w5) 지원
            self.sell_condition_indices = {
                idx for idx, name in conds.items() if name in SELL_CONDITIONS
            }

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
    # 조건 검색 결과 (초기 리스트) - OnReceiveTrCondition
    #   → “조건 검색에 있는 모든 종목” 50만원 매수
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

                print(f"[BUY-TRIGGER] (초기검색) 조건 편입 매수 트리거: cond_index={cond_index_int}, code={code}")
                self.enqueue_buy(code, TARGET_BUY_AMOUNT)

    # -----------------------------
    # 실시간 조건 편입/이탈 - OnReceiveRealCondition
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

    # -----------------------------
    # 조건 이벤트 공통 처리 로직
    # -----------------------------
    def _handle_condition_event(self, code, type, cond_name, cond_index_int):
        if cond_index_int is None:
            print(f"[COND-HANDLE] cond_index_int is None → 무시 (code={code})")
            return

        # 매수 트리거: 실시간 편입(I) (w1, w2, w3)
        if type == 'I' and cond_index_int in self.buy_condition_indices:
            print(f"[BUY-TRIGGER] (실시간) 조건 편입 매수 트리거: cond_index={cond_index_int}, code={code}")

            already = self.code_accum_buy_amount.get(code, 0)
            if already >= self.max_position_per_code:
                print(f"[BUY] {code} 누적 매수금액 {already}원 ≥ {self.max_position_per_code}원 → 추가 매수 금지")
                return

            if code in self.pending_orders:
                print(f"[BUY] 진행 중 주문 있어 스킵: {code}")
                return

            self.enqueue_buy(code, TARGET_BUY_AMOUNT)
            return

        # 매도 트리거: 실시간 이탈(D) (w4, w5) → 전량 시장가 매도
        if type == 'D' and cond_index_int in self.sell_condition_indices:
            print(f"[SELL-TRIGGER] (실시간) 조건 이탈 매도 트리거: cond_index={cond_index_int}, code={code}")

            if code in self.pending_orders:
                print(f"[SELL] 진행 중 주문 있어 스킵: {code}")
                return

            QTimer.singleShot(0, lambda c=code: self.sell_all_market(c))

    # -----------------------------
    # 매수 큐 관리
    # -----------------------------
    def enqueue_buy(self, code, amount):
        # 큐에 이미 있거나 진행 중이면 스킵
        if any(c == code for (c, _) in self.buy_queue) or code in self.pending_orders:
            print(f"[BUY-QUEUE] 이미 대기열 또는 진행중: {code}")
            return

        # 종목당 한도 체크
        already = self.code_accum_buy_amount.get(code, 0)
        if already >= self.max_position_per_code:
            print(f"[BUY-QUEUE] {code} 누적 {already}원 ≥ {self.max_position_per_code}원 → 큐 추가 안 함")
            return

        self.buy_queue.append((code, amount))
        print(f"[BUY-QUEUE] 큐 추가: {code}, 현재대기={len(self.buy_queue)}")
        if not self.is_buying:
            self._process_next_buy()

    def _process_next_buy(self):
        if self.is_buying:
            return
        if not self.buy_queue:
            return

        code, amount = self.buy_queue.popleft()
        self.is_buying = True
        # 너무 빠른 TR 연속 호출을 막기 위해 약간의 딜레이
        QTimer.singleShot(400, lambda c=code, a=amount: self._buy_market_amount_internal(c, a))

    # -----------------------------
    # 현재가 조회(opt10001) + 금액 제한 매수
    # -----------------------------
    def request_price(self, code):
        # 가격 루프 중복 실행 방지
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
        self.price_loop.exec_()  # _on_receive_tr_data에서 opt10001_req 처리 후 종료
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

    def _buy_market_amount_internal(self, code, amount):
        print(f"[BUY] _buy_market_amount_internal 진입: code={code}, amount={amount}")

        # 종목당 한도 체크 (2차 방어)
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

        # 남은 한도(≤50만원) 내에서만 수량 산출
        qty = remaining_amount // price
        if qty <= 0:
            print(
                f"[BUY] 남은 한도 {remaining_amount}원으로 매수 가능한 수량이 0: "
                f"price={price}, already={already}, per_code_limit={self.max_position_per_code}"
            )
            self.is_buying = False
            self._process_next_buy()
            return

        order_amount = qty * price
        print(
            f"[BUY] 계산 결과 → code={code}, price={price}, qty={qty}, "
            f"order_amount={order_amount}, before_accum={already}"
        )

        rqname = "buy_by_condition"
        order_type = 1           # 신규매수
        hoga = "03"              # 시장가
        self.pending_orders.add(code)
        print(f"[BUY] 주문 전송: {code} 수량={qty} 시장가 (price={price})")

        ret = self.dynamicCall(
            "SendOrder(QString,QString,QString,int,QString,int,int,QString,QString)",
            [
                rqname,
                SCREEN_ORDER,
                self.account,
                int(order_type),
                code,
                int(qty),
                0,
                hoga,
                ""
            ]
        )
        if ret != 0:
            print(f"[BUY] 주문 실패 ret={ret} code={code}")
            self.pending_orders.discard(code)
        else:
            # 종목별 누적 매수 금액만 관리
            self.code_accum_buy_amount[code] = already + order_amount
            print(
                f"[BUY] 주문 전송 성공: code={code}, qty={qty}, "
                f"종목누적={self.code_accum_buy_amount[code]}원 / 종목한도={self.max_position_per_code}원"
            )

        self.is_buying = False
        self._process_next_buy()

    # -----------------------------
    # 전량 시장가 매도
    # -----------------------------
    def sell_all_market(self, code):
        print(f"[SELL] sell_all_market 진입: code={code}")
        hold_qty = self.holdings.get(code, 0)
        if hold_qty <= 0:
            print(f"[SELL] 보유수량 0 → 매도 스킵: {code}")
            return

        price = self.request_price(code)
        if price:
            print(f"[SELL] {code} 현재가(참고용): {price}")

        rqname = "sell_by_condition"
        order_type = 2           # 신규매도
        hoga = "03"              # 시장가
        self.pending_orders.add(code)
        print(f"[SELL] 주문 전송(전량): {code} 수량={hold_qty} 시장가")

        ret = self.dynamicCall(
            "SendOrder(QString,QString,QString,int,QString,int,int,QString,QString)",
            [
                rqname,
                SCREEN_ORDER,
                self.account,
                int(order_type),
                code,
                int(hold_qty),
                0,
                hoga,
                ""
            ]
        )
        if ret != 0:
            print(f"[SELL] 주문 실패 ret={ret} code={code}")
            self.pending_orders.discard(code)
        else:
            print(f"[SELL] 주문 전송 성공: code={code}, qty={hold_qty}")
            # 매도 후 보유 0 처리
            self.holdings.pop(code, None)

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

        # 잔고변경(gubun == '1')일 때 holdings 직접 업데이트
        if gubun == '1':
            try:
                code = self.dynamicCall("GetChejanData(int)", 9001).strip()  # 종목코드
                code = code.replace("A", "")
                qty_str = self.dynamicCall("GetChejanData(int)", 930).strip()  # 보유수량
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
            except Exception as e:
                print(f"[CHEJAN-ERR] holdings 업데이트 실패: {e}")

        def do_balance():
            if not self.balance_loop.isRunning():
                self.request_balance()
            else:
                print("[CHEJAN] balance_loop 동작 중 → 잔고 조회 스킵")
            self.pending_orders.clear()

        QTimer.singleShot(1500, do_balance)


def main():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.login()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()