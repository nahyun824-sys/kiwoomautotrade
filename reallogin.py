# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop, QTimer




SCREEN_TR = "1000"
PASSWD = ""     # 여기 입력해도 되고(보안상 비추천), 비워두고 창에서 입력해도 됨
PASSWD_MEDIA = "00" # 00: 공통




def _strip(x):
    return str(x).strip()




def _to_int(x, default=0):
    s = _strip(x).replace(",", "")
    if s == "":
        return default
    try:
        return int(s)
    except:
        return default




def _to_float(x, default=0.0):
    s = _strip(x).replace(",", "")
    if s == "":
        return default
    try:
        return float(s)
    except:
        return default




class KiwoomBalanceViewer(QAxWidget):
    def __init__(self):
        super().__init__()
        print("[INIT] Kiwoom 로그인+잔고 조회 시작")


        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        if self.isNull():
            print("[ERROR] ❌ KHOpenAPI ActiveX 생성 실패")
            self.login_loop = None
            self.tr_loop = None
            return


        self.OnEventConnect.connect(self._on_event_connect)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)


        self.login_loop = QEventLoop()
        self.tr_loop = QEventLoop()


        self.account = None
        self.deposit_detail = {}
        self.holdings = []
        self.balance_summary = {}
        self._opw00018_next = "0"


    def _dc(self, signature, args=None):
        if args is None:
            return self.dynamicCall(signature)
        return self.dynamicCall(signature, args)


    def _get_comm_data(self, trcode, rqname, index, item):
        return _strip(self._dc("GetCommData(QString, QString, int, QString)", [trcode, rqname, index, item]))


    def _get_repeat_cnt(self, trcode, rqname):
        return int(self._dc("GetRepeatCnt(QString, QString)", [trcode, rqname]))


    # ===== 로그인 =====
    def login(self):
        state = self._dc("GetConnectState()")
        print(f"[LOGIN] GetConnectState={state} (1=연결,0=미연결)")
        print("[LOGIN] CommConnect() 호출")
        self._dc("CommConnect()")
        self.login_loop.exec_()


    def _on_event_connect(self, err_code):
        print(f"[EVENT] OnEventConnect err_code={err_code}")
        if int(err_code) != 0:
            print("[LOGIN] ❌ 로그인 실패/취소")
            self.login_loop.exit()
            return


        print("[LOGIN] ✅ 로그인 성공")


        accno = self._dc("GetLoginInfo(QString)", ["ACCNO"])
        acc_list = [a.strip() for a in str(accno).split(";") if a.strip()]
        print(f"[LOGIN] ACCNO(list)={acc_list}")


        if not acc_list:
            print("[ERROR] 계좌번호를 못 가져왔어")
            self.login_loop.exit()
            return


        self.account = acc_list[0]
        print(f"[LOGIN] 사용 계좌: {self.account}")
        self.login_loop.exit()


    # ===== 비번/계좌 입력창 유도 (KOA_Functions) =====
    def show_account_window(self):
        """
        키움 기본 UI(계좌/비밀번호 입력 흐름)를 띄워서 (44) 팝업을 통과시키는 목적.
        이 창에서 비밀번호 입력 후 확인하면, 이후 TR이 정상 처리되는 케이스가 많음.
        """
        print("[UI] KOA_Functions('ShowAccountWindow') 호출 (계좌/비번 입력창 유도)")
        try:
            ret = self._dc("KOA_Functions(QString, QString)", ["ShowAccountWindow", ""])
            print(f"[UI] ShowAccountWindow ret={ret}")
        except Exception as e:
            print(f"[UI] ShowAccountWindow 호출 실패: {repr(e)}")


        # 사용자가 입력할 시간을 주기 위해 잠깐 대기
        QTimer.singleShot(2500, lambda: None)


    # ===== TR: 예수금 =====
    def request_deposit(self):
        print("[TR] 예수금상세현황요청(opw00001) 시작")
        self._dc("SetInputValue(QString, QString)", ["계좌번호", self.account])
        self._dc("SetInputValue(QString, QString)", ["비밀번호", PASSWD]) # 비워둬도 됨(대신 팝업 입력)
        self._dc("SetInputValue(QString, QString)", ["비밀번호입력매체구분", PASSWD_MEDIA])
        self._dc("SetInputValue(QString, QString)", ["조회구분", "2"])


        ret = self._dc("CommRqData(QString, QString, int, QString)", ["opw00001_req", "opw00001", 0, SCREEN_TR])
        print(f"[TR] opw00001 CommRqData ret={ret}")
        if int(ret) != 0:
            return False
        self.tr_loop.exec_()
        return True


    # ===== TR: 잔고 =====
    def request_balance(self):
        print("[TR] 계좌평가잔고내역요청(opw00018) 시작")
        self.holdings = []
        self._opw00018_next = "0"
        return self._request_opw00018(prev_next=0)


    def _request_opw00018(self, prev_next=0):
        self._dc("SetInputValue(QString, QString)", ["계좌번호", self.account])
        self._dc("SetInputValue(QString, QString)", ["비밀번호", PASSWD]) # 비워둬도 됨(대신 팝업 입력)
        self._dc("SetInputValue(QString, QString)", ["비밀번호입력매체구분", PASSWD_MEDIA])
        self._dc("SetInputValue(QString, QString)", ["조회구분", "2"])


        ret = self._dc("CommRqData(QString, QString, int, QString)", ["opw00018_req", "opw00018", prev_next, SCREEN_TR])
        print(f"[TR] opw00018 CommRqData ret={ret}")
        if int(ret) != 0:
            return False
        self.tr_loop.exec_()
        return True


    # ===== TR 수신 =====
    def _on_receive_tr_data(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        rqname = _strip(rqname)
        trcode = _strip(trcode)
        prev_next = _strip(prev_next)


        if rqname == "opw00001_req":
            deposit = self._get_comm_data(trcode, rqname, 0, "예수금")
            d2_deposit = self._get_comm_data(trcode, rqname, 0, "d+2추정예수금")


            self.deposit_detail = {
                "예수금": _to_int(deposit),
                "D+2추정예수금": _to_int(d2_deposit),
            }
            print("[TR] ✅ 예수금 수신 완료")
            self.tr_loop.exit()
            return


        if rqname == "opw00018_req":
            total_buy = self._get_comm_data(trcode, rqname, 0, "총매입금액")
            total_eval = self._get_comm_data(trcode, rqname, 0, "총평가금액")
            total_pl = self._get_comm_data(trcode, rqname, 0, "총평가손익금액")
            total_yield = self._get_comm_data(trcode, rqname, 0, "총수익률(%)")


            self.balance_summary = {
                "총매입금액": _to_int(total_buy),
                "총평가금액": _to_int(total_eval),
                "총평가손익": _to_int(total_pl),
                "총수익률(%)": _to_float(total_yield),
            }


            cnt = self._get_repeat_cnt(trcode, rqname)
            for i in range(cnt):
                code = self._get_comm_data(trcode, rqname, i, "종목번호").replace("A", "").strip()
                name = self._get_comm_data(trcode, rqname, i, "종목명")
                qty = self._get_comm_data(trcode, rqname, i, "보유수량")
                buy_price = self._get_comm_data(trcode, rqname, i, "매입가")
                cur_price = self._get_comm_data(trcode, rqname, i, "현재가")
                pl = self._get_comm_data(trcode, rqname, i, "평가손익")
                yieldp = self._get_comm_data(trcode, rqname, i, "수익률(%)")


                self.holdings.append({
                    "code": code,
                    "name": name,
                    "qty": _to_int(qty),
                    "buy_price": _to_int(buy_price),
                    "cur_price": _to_int(cur_price),
                    "pl": _to_int(pl),
                    "yield(%)": _to_float(yieldp),
                })


            if prev_next == "2":
                print("[TR] 다음 페이지 있음 → 추가 요청")
                self._request_opw00018(prev_next=2)
                return


            print("[TR] ✅ 잔고 수신 완료(마지막 페이지)")
            self.tr_loop.exit()
            return


    def print_result(self):
        print("\n" + "=" * 80)
        print("[RESULT] 예수금")
        if self.deposit_detail:
            print(f"예수금: {self.deposit_detail.get('예수금', 0):,} 원")
            print(f"D+2추정예수금: {self.deposit_detail.get('D+2추정예수금', 0):,} 원")
        else:
            print("(없음)")


        print("\n[RESULT] 계좌평가 요약")
        if self.balance_summary:
            s = self.balance_summary
            print(f"총매입금액: {s['총매입금액']:,} 원")
            print(f"총평가금액: {s['총평가금액']:,} 원")
            print(f"총평가손익: {s['총평가손익']:,} 원")
            print(f"총수익률(%): {s['총수익률(%)']}")
        else:
            print("(없음)")


        print("\n[RESULT] 보유종목")
        if not self.holdings:
            print("(보유종목 없음)")
        else:
            for h in self.holdings:
                print(f"- {h['code']} {h['name']} | 수량:{h['qty']:,} | 매입:{h['buy_price']:,} | 현재:{h['cur_price']:,} | 손익:{h['pl']:,} | 수익률:{h['yield(%)']}")


        print("=" * 80 + "\n")




def main():
    app = QApplication(sys.argv)


    kiwoom = KiwoomBalanceViewer()
    if kiwoom.isNull():
        print("[END] ActiveX 생성 실패로 종료")
        QTimer.singleShot(200, app.quit)
        app.exec_()
        return


    kiwoom.login()
    if not kiwoom.account:
        print("[END] 로그인/계좌확인 실패로 종료")
        QTimer.singleShot(200, app.quit)
        app.exec_()
        return


    # ✅ 계좌/비번 입력 흐름 유도
    kiwoom.show_account_window()


    # 약간 쉬었다가 TR (입력창 확인 후 진행)
    QTimer.singleShot(1500, lambda: None)


    ok1 = kiwoom.request_deposit()
    ok2 = kiwoom.request_balance()


    kiwoom.print_result()


    QTimer.singleShot(200, app.quit)
    app.exec_()




if __name__ == "__main__":
    main()
