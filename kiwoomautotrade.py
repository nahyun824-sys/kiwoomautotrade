# -*- coding: utf-8 -*-
"""
Kiwoom 조건식 자동매매 (통합 완성본)

- 매수 조건: BUY_COND_NAMES
- 매도 조건: SELL_COND_NAMES (w, x2 등) <<<<< 여기 정의만 사용 (하드코딩 X)
- SELL 유효 종목 집합 = sell_cond_codes[w] ∪ sell_cond_codes[x2] ∪ ...
- 매도 조건 이탈(D) 즉시 전량 매도
- 시작 후 AUTO-SELL 유예(기본 180초): SELL 초기 편입 세트 수신 전 강제매도 방지
- 1분마다 AUTO-SELL 점검:
    (1) -10% 손절
    (2) 매도조건 1:1 매핑 실패(= SELL 유효 집합에 없으면) 전량 매도
- 조건식 구독 안정화:
    (1) SendCondition 실패(ret=0) 시 SendConditionStop 후 재시도 + 딜레이/backoff

✅ 이번 패치 핵심:
- "분할매수" 허용 (보유중이어도 30만원 한도 남으면 매수 가능)
- "미체결 매수"까지 포함한 30만원 캡 강제 (보유평가 + 미체결추정 <= MAX_POSITION_PER_CODE)
- pending 해제는 "미체결수량=0"일 때만 (실전 중복매수 방지 핵심)
"""

import sys
import time
import datetime
from collections import deque

from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer


# =========================
# 사용자 설정 (여기만 건드리면 됨)
# =========================
BUY_COND_NAMES = {"w1", "w2", "w3", "x1"}
SELL_COND_NAMES = {"w", "x2"}

TARGET_BUY_AMOUNT = 300000       # 1회 매수 시도 예산(분할매수 단위)
MAX_POSITION_PER_CODE = 300000   # ✅ 종목당 누적 보유(미체결 포함) 최대 한도
STOPLOSS_PCT = -10.0

AUTO_SELL_INTERVAL_SEC = 60
PRICE_RETRY_MAX = 3
PRICE_RETRY_SLEEP = 0.8
BALANCE_COOLDOWN_SEC = 3
AUTO_SELL_START_GRACE_SEC = 180

# 화면번호(중복 사용 금지)
SCREEN_LOGIN = "0000"
SCREEN_TR_PRICE = "1001"
SCREEN_TR_BALANCE = "1002"
SCREEN_ORDER = "2001"


# ✅ Chejan FID (키움 표준)
FID_CODE = 9001
FID_ORDER_NO = 9203
FID_ORDER_STATUS = 913     # 주문상태
FID_ORDER_QTY = 900        # 주문수량
FID_UNFILLED_QTY = 902     # 미체결수량
FID_ORDER_GUBUN = 905      # 주문구분(+매수/-매도 등)
FID_FILLED_QTY = 911       # 체결량
FID_FILLED_PRICE = 910     # 체결가


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        # 이벤트 연결
        self.OnEventConnect.connect(self._on_event_connect)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)
        self.OnReceiveConditionVer.connect(self._on_receive_condition_ver)
        self.OnReceiveTrCondition.connect(self._on_receive_tr_condition)
        self.OnReceiveRealCondition.connect(self._on_receive_real_condition)
        self.OnReceiveChejanData.connect(self._on_receive_chejan)

        # 루프
        self.login_loop = QEventLoop()
        self.condver_loop = QEventLoop()
        self.price_loop = QEventLoop()
        self.balance_loop = QEventLoop()

        # 계좌/서버
        self.account_no = None
        self.server_gubun = None

        # 조건 목록/인덱스 매핑
        self.cond_name_to_index = {}
        self.cond_index_to_name = {}

        # 보유/포지션/주문 상태
        self.holdings = {}       # code -> qty (잔고TR/체잔 반영)
        self.pos = {}            # code -> {"avg_price": int, "qty": int, "buy_cond": str, "buy_time": datetime}
        self.pending_codes = set()  # 코드 단위 "주문 진행중" 마킹

        # ✅ 주문번호 기반 pending 추적 (실전 안전)
        # order_no -> {"code","side","qty","est_price","unfilled_qty"}
        self.pending_orders = {}

        # ✅ 코드별 미체결 매수 노출금액(추정) 캐시
        # code -> sum(unfilled_qty * est_price)
        self.pending_buy_exposure = {}

        # 매수 큐
        self.buy_queue = deque()     # (code, cond_name)
        self.buy_queue_set = set()   # code 중복 방지(큐)
        self.last_buy_ts = {}        # code -> timestamp (짧은 시간 중복 트리거 방지)

        # 가격 캐시
        self.price_cache = {}
        self.price_waiting_code = None
        self.price_waiting_retry = 0

        # 잔고 쿨다운
        self._last_balance_ts = 0.0

        # SELL 조건별 세트
        self.sell_cond_codes = {name: set() for name in SELL_COND_NAMES}

        # AUTO-SELL 잠금
        self.sell_exit_lock = set()

        # AUTO-SELL 시작 유예
        self._start_ts = time.time()
        self._auto_sell_grace_printed = False

        # 타이머
        self.timer_buy = QTimer()
        self.timer_buy.timeout.connect(self._process_buy_queue)

        self.timer_auto_sell = QTimer()
        self.timer_auto_sell.timeout.connect(self._auto_sell_check)

        print("[INIT] 프로그램 초기화 완료")

    # ---------------------------
    # 공통 유틸
    # ---------------------------
    def _now(self):
        return datetime.datetime.now()

    def _ts(self):
        return time.time()

    def _safe_int(self, s):
        try:
            s = str(s).strip()
            if s == "":
                return None
            return int(s)
        except:
            return None

    def _strip_code(self, code):
        return str(code).strip()

    def _name(self, code):
        code = self._strip_code(code)
        try:
            n = self.get_master_code_name(code)
            return str(n).strip() if n else ""
        except:
            return ""

    def _fmt(self, code):
        code = self._strip_code(code)
        n = self._name(code)
        return f"{code}({n})" if n else f"{code}"

    def _get_valid_sell_union(self):
        s = set()
        for _name, _codes in self.sell_cond_codes.items():
            s |= _codes
        return s

    def _get_screen_for_cond(self, cond_index):
        return str(5000 + int(cond_index))

    def _is_auto_sell_grace(self):
        return (self._ts() - self._start_ts) < AUTO_SELL_START_GRACE_SEC

    # ---------------------------
    # ✅ 32비트 호환 dynamicCall 래퍼
    # ---------------------------
    def _dc(self, signature, args=None):
        if args is None:
            return self.dynamicCall(signature)
        return self.dynamicCall(signature, args)

    # ---------------------------
    # Kiwoom API wrapper
    # ---------------------------
    def comm_connect(self):
        print("[LOGIN] 로그인 요청")
        self._dc("CommConnect()")
        self.login_loop.exec_()

    def get_login_info(self, tag):
        return self._dc("GetLoginInfo(QString)", [tag])

    def get_condition_load(self):
        self._dc("GetConditionLoad()")
        self.condver_loop.exec_()

    def send_condition(self, screen_no, cond_name, cond_index, search=1):
        return self._dc(
            "SendCondition(QString, QString, int, int)",
            [screen_no, cond_name, int(cond_index), int(search)]
        )

    def send_condition_stop(self, screen_no, cond_name, cond_index):
        return self._dc(
            "SendConditionStop(QString, QString, int)",
            [screen_no, cond_name, int(cond_index)]
        )

    def send_order(self, rqname, screen_no, acc_no, order_type, code, qty, price, hoga, org_order_no=""):
        return self._dc(
            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
            [rqname, screen_no, acc_no, int(order_type), code, int(qty), int(price), hoga, org_order_no]
        )

    def set_input_value(self, key, value):
        self._dc("SetInputValue(QString, QString)", [key, value])

    def comm_rq_data(self, rqname, trcode, prev_next, screen_no):
        return self._dc(
            "CommRqData(QString, QString, int, QString)",
            [rqname, trcode, int(prev_next), screen_no]
        )

    def get_comm_data(self, trcode, rqname, index, item):
        return self._dc(
            "GetCommData(QString, QString, int, QString)",
            [trcode, rqname, int(index), item]
        )

    def get_repeat_cnt(self, trcode, rqname):
        return self._dc("GetRepeatCnt(QString, QString)", [trcode, rqname])

    def get_master_code_name(self, code):
        return self._dc("GetMasterCodeName(QString)", [code])

    # ---------------------------
    # 이벤트 핸들러
    # ---------------------------
    def _on_event_connect(self, err_code):
        if err_code == 0:
            print("[LOGIN] 로그인 성공")
            self.account_no = self.get_login_info("ACCNO").split(";")[0].strip()
            self.server_gubun = self.get_login_info("GetServerGubun")
            print(f"[LOGIN] 계좌번호: {self.account_no}")
            print(f"[LOGIN] 서버구분(1=모의, 0=실): {self.server_gubun}")
        else:
            print(f"[LOGIN] 로그인 실패 err_code={err_code}")
        self.login_loop.exit()

    def _on_receive_condition_ver(self, ret, msg):
        if int(ret) == 1:
            cond_list = self._dc("GetConditionNameList()")
            items = [x for x in cond_list.split(";") if x.strip()]
            for it in items:
                idx, name = it.split("^")
                idx = int(idx)
                name = name.strip()
                self.cond_name_to_index[name] = idx
                self.cond_index_to_name[idx] = name

            print("[COND] 조건 목록 로드 완료")
            print(f"[COND] BUY_CONDITIONS={BUY_COND_NAMES} -> idx={ {self.cond_name_to_index[n] for n in BUY_COND_NAMES if n in self.cond_name_to_index} }")
            print(f"[COND] SELL_CONDITIONS={SELL_COND_NAMES} -> idx={ {self.cond_name_to_index[n] for n in SELL_COND_NAMES if n in self.cond_name_to_index} }")

            all_names = sorted(BUY_COND_NAMES | SELL_COND_NAMES)
            for name in all_names:
                if name not in self.cond_name_to_index:
                    print(f"[COND-SUB] ❌ 조건식 없음 -> 스킵: {name}")
                    continue

                idx = self.cond_name_to_index[name]
                scr = self._get_screen_for_cond(idx)

                ok = 0
                for try_no in range(1, 6):
                    try:
                        self.send_condition_stop(scr, name, idx)
                    except Exception:
                        pass

                    time.sleep(0.3)
                    ok = self.send_condition(scr, name, idx, 1)
                    print(f"[COND-SUB] 구독 시도: {name} idx={idx} scr={scr} ret={ok} (try {try_no}/5)")

                    if int(ok) == 1:
                        print(f"[COND-SUB] ✅ 구독 성공: {name} idx={idx} scr={scr}")
                        break

                    time.sleep(0.8 + 0.4 * try_no)

                if int(ok) != 1:
                    print(f"[COND-SUB] ❌ 구독 최종 실패: {name} idx={idx}")

        else:
            print(f"[COND] 조건목록 로드 실패 ret={ret}, msg={msg}")

        self.condver_loop.exit()

    def _on_receive_tr_condition(self, screen_no, code_list, cond_name, cond_index, next_):
        cond_name = str(cond_name).strip()
        codes = [c.strip() for c in str(code_list).split(";") if c.strip()]

        print(f"[TRCOND] scr={screen_no} cond={cond_name} idx={cond_index} next={next_}")
        if codes:
            pretty = ";".join([self._fmt(c) for c in codes])
            print(f"[TRCOND] codes({len(codes)}): {pretty}")
        else:
            print("[TRCOND] codes(0): (empty)")

        if cond_name in BUY_COND_NAMES:
            for code in codes:
                self._enqueue_buy(code, cond_name, reason="초기검색")

        if cond_name in SELL_COND_NAMES:
            self.sell_cond_codes[cond_name] = set(codes)
            print(f"[SELL-COND-INIT] cond={cond_name} 편입세트({len(codes)}): {', '.join([self._fmt(c) for c in codes])}")

    def _on_receive_real_condition(self, code, event_type, cond_name, cond_index):
        code = self._strip_code(code)
        event_type = str(event_type).strip()
        cond_name = str(cond_name).strip()

        print(f"[COND-REAL] cond={cond_name} {'편입(I)' if event_type=='I' else '이탈(D)'}({event_type}): {self._fmt(code)}")

        if cond_name in BUY_COND_NAMES and event_type == "I":
            self._enqueue_buy(code, cond_name, reason="실시간")

        if cond_name in SELL_COND_NAMES:
            if event_type == "I":
                self.sell_cond_codes[cond_name].add(code)
            else:
                self.sell_cond_codes[cond_name].discard(code)
                self._sell_all(code, sell_reason=f"EXIT_{cond_name}", with_lock=True)

    def _on_receive_tr_data(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        rqname = str(rqname).strip()
        trcode = str(trcode).strip()

        if rqname == "opt10001_price":
            code = self.price_waiting_code
            price_raw = self.get_comm_data(trcode, rqname, 0, "현재가")
            price = self._safe_int(price_raw.replace("+", "").replace("-", "")) if price_raw else None

            self.price_cache[code] = price
            print(f"[PRICE] 응답: {self._fmt(code)} 현재가={price}")
            self.price_loop.exit()

        elif rqname == "opw00018_balance":
            cnt = self.get_repeat_cnt(trcode, rqname)
            new_holdings = {}

            for i in range(cnt):
                code = self.get_comm_data(trcode, rqname, i, "종목번호").strip()
                code = code.replace("A", "").strip()
                qty_raw = self.get_comm_data(trcode, rqname, i, "보유수량").strip()
                qty = self._safe_int(qty_raw)
                if code and qty is not None and qty > 0:
                    new_holdings[code] = qty

            self.holdings = new_holdings
            pretty = {self._fmt(k): v for k, v in self.holdings.items()}
            print(f"[BALANCE] 보유종목({len(self.holdings)}): {pretty}")
            self.balance_loop.exit()

    # ---------------------------
    # ✅ 주문/체결(체잔) 처리 (중복매수 방지 핵심)
    # ---------------------------
    def _recalc_pending_buy_exposure(self):
        expo = {}
        for o in self.pending_orders.values():
            if o.get("side") != "BUY":
                continue
            code = o.get("code")
            unfilled = o.get("unfilled_qty", 0) or 0
            est_price = o.get("est_price", 0) or 0
            if code and unfilled > 0 and est_price > 0:
                expo[code] = expo.get(code, 0) + (unfilled * est_price)
        self.pending_buy_exposure = expo

    def _on_receive_chejan(self, gubun, item_cnt, fid_list):
        gubun = str(gubun).strip()
        print(f"[CHEJAN] 수신: gubun={gubun} item_cnt={item_cnt}")

        try:
            code = self._dc("GetChejanData(int)", [FID_CODE]).strip()
            code = code.replace("A", "").strip()
            if not code:
                return

            name = self._name(code)

            if gubun == "0":
                order_no = str(self._dc("GetChejanData(int)", [FID_ORDER_NO]) or "").strip()
                status = str(self._dc("GetChejanData(int)", [FID_ORDER_STATUS]) or "").strip()
                order_gubun = str(self._dc("GetChejanData(int)", [FID_ORDER_GUBUN]) or "").strip()

                order_qty = self._safe_int(self._dc("GetChejanData(int)", [FID_ORDER_QTY]))
                unfilled_qty = self._safe_int(self._dc("GetChejanData(int)", [FID_UNFILLED_QTY]))
                filled_qty = self._safe_int(self._dc("GetChejanData(int)", [FID_FILLED_QTY]))
                filled_price = self._safe_int(self._dc("GetChejanData(int)", [FID_FILLED_PRICE]))

                # 주문구분에 "매수"/"매도"가 들어오는 경우가 많음(환경에 따라 +매수/-매도 등)
                side = "BUY" if ("매수" in order_gubun) else ("SELL" if ("매도" in order_gubun) else None)

                # pending_orders 업데이트
                if order_no:
                    if order_no not in self.pending_orders:
                        # est_price는 "주문 넣을 때 추정가"를 쓰는 게 맞는데,
                        # 체잔에서 주문가가 비거나 시장가는 0일 수 있어서,
                        # 여기선 체결가/현재가/0 중 가능한 걸로 보수적으로 잡음
                        est = filled_price if (filled_price and filled_price > 0) else (self.price_cache.get(code) or 0)
                        self.pending_orders[order_no] = {
                            "code": code,
                            "side": side or "BUY",     # 모르면 BUY로 두면 위험할 수 있으니, 아래 unfilled 없으면 pending 유지됨
                            "qty": order_qty or 0,
                            "est_price": est or 0,
                            "unfilled_qty": unfilled_qty if unfilled_qty is not None else (order_qty or 0),
                            "status": status,
                        }
                    else:
                        o = self.pending_orders[order_no]
                        if side:
                            o["side"] = side
                        if order_qty is not None:
                            o["qty"] = order_qty
                        if unfilled_qty is not None:
                            o["unfilled_qty"] = unfilled_qty
                        if status:
                            o["status"] = status
                        # 체결가가 들어오면 추정가를 체결가로 보정(분할체결 방어)
                        if filled_price and filled_price > 0:
                            o["est_price"] = filled_price

                # pending_codes 해제는 "미체결 0"일 때만
                # (미체결수량 FID가 안 들어오면 안전하게 유지)
                if unfilled_qty is not None and unfilled_qty == 0:
                    if code in self.pending_codes:
                        self.pending_codes.discard(code)
                        print(f"[CHEJAN-PENDING] 해제: {code}({name}) (unfilled=0)")
                else:
                    # 아직 미체결이 남아있으면 pending 유지
                    self.pending_codes.add(code)

                # 노출금액 재계산
                self._recalc_pending_buy_exposure()

            elif gubun == "1":
                # 잔고 반영
                qty_raw = self._dc("GetChejanData(int)", [930])
                qty = self._safe_int(qty_raw)
                if qty is not None:
                    if qty > 0:
                        self.holdings[code] = qty
                    else:
                        self.holdings.pop(code, None)
                    print(f"[CHEJAN-HOLDINGS] 반영: {code}({name}) qty={qty} -> holdings_cnt={len(self.holdings)}")

        except Exception:
            pass

    # ---------------------------
    # TR 요청 로직
    # ---------------------------
    def request_balance(self, force=False):
        now = self._ts()
        if not force and (now - self._last_balance_ts) < BALANCE_COOLDOWN_SEC:
            print("[BALANCE] 최근 조회 -> 스킵")
            return
        if self.balance_loop.isRunning():
            print("[BALANCE] balance_loop 동작중 -> 스킵")
            return

        self._last_balance_ts = now
        print("[BALANCE] 요청: opw00018 잔고조회")
        self.set_input_value("계좌번호", self.account_no)
        self.set_input_value("비밀번호", "0000")
        self.set_input_value("비밀번호입력매체구분", "00")
        self.set_input_value("조회구분", "2")
        self.comm_rq_data("opw00018_balance", "opw00018", 0, SCREEN_TR_BALANCE)
        self.balance_loop.exec_()

    def get_current_price(self, code):
        code = self._strip_code(code)

        if self.price_loop.isRunning():
            cached = self.price_cache.get(code)
            print(f"[PRICE] price_loop 동작중 -> 캐시반환: {self._fmt(code)} cached={cached}")
            return cached

        self.price_waiting_code = code
        self.price_waiting_retry = 0

        while self.price_waiting_retry < PRICE_RETRY_MAX:
            self.set_input_value("종목코드", code)
            ret = self.comm_rq_data("opt10001_price", "opt10001", 0, SCREEN_TR_PRICE)

            if int(ret) == -200:
                self.price_waiting_retry += 1
                print(f"[PRICE] 과부하(-200) -> {PRICE_RETRY_SLEEP}s 대기 후 재시도({self.price_waiting_retry}/{PRICE_RETRY_MAX}): {self._fmt(code)}")
                time.sleep(PRICE_RETRY_SLEEP)
                continue
            if int(ret) != 0:
                print(f"[PRICE] TR 요청 실패 ret={ret} code={self._fmt(code)}")
                return None

            print(f"[PRICE] TR요청 성공 -> 응답대기: {self._fmt(code)}")
            self.price_loop.exec_()
            return self.price_cache.get(code)

        return None

    # ---------------------------
    # ✅ 한도 계산 유틸 (보유 + 미체결 포함)
    # ---------------------------
    def _get_position_exposure(self, code, price):
        """
        code의 현재 노출금액(추정) = 보유수량*price + 미체결매수노출
        """
        code = self._strip_code(code)
        held_qty = self.holdings.get(code, 0) or 0
        held_value = (held_qty * price) if (price and price > 0) else 0
        pending_value = self.pending_buy_exposure.get(code, 0) or 0
        return held_value + pending_value

    # ---------------------------
    # BUY / SELL
    # ---------------------------
    def _enqueue_buy(self, code, cond_name, reason=""):
        code = self._strip_code(code)
        print(f"[BUY-TRIGGER] {reason} -> cond={cond_name}, target={self._fmt(code)}")

        # ✅ 큐 중복 방지
        if code in self.buy_queue_set:
            print(f"[BUY-QUEUE] 스킵: 이미 대기열에 존재 -> {self._fmt(code)}")
            return

        # ✅ 주문 진행중이면 스킵
        if code in self.pending_codes:
            print(f"[BUY-QUEUE] 스킵: 진행중 주문 존재 -> {self._fmt(code)}")
            return

        # ✅ 너무 빠른 재트리거 방지
        last_ts = self.last_buy_ts.get(code, 0)
        if self._ts() - last_ts < 2.0:
            print(f"[BUY-QUEUE] 스킵: 너무 빠른 재트리거 -> {self._fmt(code)}")
            return
        self.last_buy_ts[code] = self._ts()

        # ✅ 분할매수는 여기서 막지 않음 (한도는 실제 주문 직전에 계산)
        self.buy_queue.append((code, cond_name))
        self.buy_queue_set.add(code)
        print(f"[BUY-QUEUE] 추가: {self._fmt(code)} cond={cond_name} queue_len={len(self.buy_queue)}")

    def _process_buy_queue(self):
        if not self.buy_queue:
            return

        code, cond_name = self.buy_queue.popleft()
        self.buy_queue_set.discard(code)

        if code in self.pending_codes:
            print(f"[BUY-QUEUE] (deq) 스킵: 진행중 주문 존재 -> {self._fmt(code)}")
            return

        self._buy_market_split_cap(code, TARGET_BUY_AMOUNT, cond_name)

    def _buy_market_split_cap(self, code, budget, cond_name):
        """
        ✅ 분할매수 + 종목당 한도(MAX_POSITION_PER_CODE) 강제
        (보유평가 + 미체결추정 + 이번주문) <= MAX_POSITION_PER_CODE
        """
        code = self._strip_code(code)
        print(f"[BUY] 진입: target={self._fmt(code)} budget={budget} cond={cond_name}")

        price = self.get_current_price(code)
        if not price or price <= 0:
            print(f"[BUY-SKIP] 현재가 조회 실패 -> 스킵: {self._fmt(code)} price={price}")
            return

        # ✅ 현재 노출(보유+미체결) 계산
        exposure = self._get_position_exposure(code, price)
        remaining = MAX_POSITION_PER_CODE - exposure

        if remaining <= 0:
            print(f"[BUY-SKIP] 한도초과(노출={exposure} >= {MAX_POSITION_PER_CODE}) -> 스킵: {self._fmt(code)}")
            return

        use_budget = min(int(budget), int(remaining))
        qty = int(use_budget // price)

        if qty <= 0:
            print(f"[BUY-SKIP] 금액부족(1주 미만) -> 스킵: {self._fmt(code)} price={price} remaining={remaining} use_budget={use_budget}")
            return

        order_amount = qty * price

        # ✅ 최종 안전장치: 이번 주문까지 합쳐도 한도 넘으면 qty 줄이기
        if exposure + order_amount > MAX_POSITION_PER_CODE:
            max_qty = int((MAX_POSITION_PER_CODE - exposure) // price)
            if max_qty <= 0:
                print(f"[BUY-SKIP] 한도내 수량=0 -> 스킵: {self._fmt(code)} exposure={exposure} price={price}")
                return
            qty = max_qty
            order_amount = qty * price

        print(f"[BUY] 계산: {self._fmt(code)} price={price} qty={qty} order_amount={order_amount} "
              f"(exposure={exposure}, after={exposure + order_amount}/{MAX_POSITION_PER_CODE})")

        # ✅ 주문 진행중 마킹
        self.pending_codes.add(code)

        print(f"[BUY] 주문전송: {self._fmt(code)} qty={qty} 시장가")
        ret = self.send_order("buy_by_condition", SCREEN_ORDER, self.account_no, 1, code, qty, 0, "03")

        if int(ret) == 0:
            print(f"[BUY] ✅ 주문성공: {self._fmt(code)} qty={qty} used={order_amount}")

            # pos 기록(간단 추정, 잔고/체잔으로 보정됨)
            prev = self.pos.get(code)
            if not prev:
                self.pos[code] = {
                    "avg_price": price,
                    "qty": qty,
                    "buy_cond": cond_name,
                    "buy_time": self._now(),
                }
            else:
                # 분할매수: 단순 가중평균(추정)
                old_qty = prev.get("qty", 0) or 0
                old_avg = prev.get("avg_price", 0) or 0
                new_qty = old_qty + qty
                if new_qty > 0:
                    new_avg = int((old_avg * old_qty + price * qty) / new_qty)
                else:
                    new_avg = price
                prev["qty"] = new_qty
                prev["avg_price"] = new_avg
                prev["buy_cond"] = cond_name
                # buy_time은 최초 유지

            # ✅ 미체결 노출 추정값을 주문 직후에도 보수적으로 반영 (체잔 오기 전 중복 방지)
            # (시장가라 주문가를 모르니 현재가로 추정)
            self.pending_buy_exposure[code] = self.pending_buy_exposure.get(code, 0) + (qty * price)

        else:
            self.pending_codes.discard(code)
            print(f"[BUY] ❌ 주문실패 ret={ret}: {self._fmt(code)}")

    def _sell_all(self, code, sell_reason="", with_lock=False):
        code = self._strip_code(code)

        qty = self.holdings.get(code, 0)
        if qty <= 0:
            return

        if with_lock:
            if code in self.sell_exit_lock:
                return
            self.sell_exit_lock.add(code)

        if code in self.pending_codes:
            print(f"[SELL] 스킵: 진행중 주문 존재 -> {self._fmt(code)}")
            return

        self.pending_codes.add(code)
        print(f"[SELL] 진입: {self._fmt(code)} reason={sell_reason}")

        p = self.get_current_price(code)
        if p is not None:
            print(f"[SELL] 현재가(참고): {self._fmt(code)} price={p}")

        print(f"[SELL] 주문전송(전량): {self._fmt(code)} qty={qty} 시장가")
        ret = self.send_order("sell_all", SCREEN_ORDER, self.account_no, 2, code, qty, 0, "03")

        if int(ret) == 0:
            print(f"[SELL] ✅ 주문성공: {self._fmt(code)} qty={qty} reason={sell_reason}")
            self.pos.pop(code, None)
        else:
            self.pending_codes.discard(code)
            if with_lock:
                self.sell_exit_lock.discard(code)
            print(f"[SELL] ❌ 주문실패 ret={ret}: {self._fmt(code)}")

    # ---------------------------
    # AUTO-SELL
    # ---------------------------
    def _auto_sell_check(self):
        if self._is_auto_sell_grace():
            if not self._auto_sell_grace_printed:
                self._auto_sell_grace_printed = True
                print(f"[AUTO-SELL] 시작 유예 적용: {AUTO_SELL_START_GRACE_SEC}s 동안 AUTO-SELL 미동작")
            return

        current_codes = set(self.holdings.keys())
        valid_sell_codes = self._get_valid_sell_union()
        pending_codes = set(self.pending_codes)

        print("[AUTO-SELL] 점검 시작 (1) 손절, (2) 매핑실패 매도")

        # (1) 손절
        for code in list(current_codes):
            if code in pending_codes:
                continue
            if code not in self.pos:
                continue

            buy_price = self.pos[code].get("avg_price")
            if not buy_price or buy_price <= 0:
                continue

            cur = self.get_current_price(code)
            if not cur or cur <= 0:
                continue

            pct = (cur - buy_price) / buy_price * 100.0
            if pct <= STOPLOSS_PCT:
                print(f"[STOPLOSS] {STOPLOSS_PCT}% 이하 -> 전량 매도: {self._fmt(code)} pct={pct:.2f}%")
                self._sell_all(code, sell_reason="STOPLOSS", with_lock=True)

        # (2) 매도조건 매핑 실패
        for code in list(current_codes):
            if code in pending_codes:
                continue
            if code not in valid_sell_codes:
                qty = self.holdings.get(code, 0)
                if qty > 0:
                    print(f"[AUTO-SELL] 매핑 실패 -> 전량 매도: {self._fmt(code)} qty={qty}")
                    self._sell_all(code, sell_reason="AUTO_SELL_UNMAPPED", with_lock=True)

        print("[AUTO-SELL] 점검 종료")

    # ---------------------------
    # 실행 플로우
    # ---------------------------
    def run(self):
        self.comm_connect()
        self.request_balance(force=True)
        self.get_condition_load()

        self.timer_buy.start(300)
        self.timer_auto_sell.start(AUTO_SELL_INTERVAL_SEC * 1000)

        QApplication.instance().exec_()


def main():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.run()


if __name__ == "__main__":
    main()
