# -*- coding: utf-8 -*-
"""Configuration constants for the Kiwoom condition trading bot."""

# 화면번호(중복 사용 금지)
SCREEN_LOGIN = "0000"
SCREEN_CONDITION = "0001"   # 실시간 조건
SCREEN_TR_PRICE = "1001"    # opt10001
SCREEN_TR_BALANCE = "1002"  # opw00018
SCREEN_ORDER = "2001"

# 매매 한도
TARGET_BUY_AMOUNT = 300000
MAX_POSITION_PER_CODE = 300000

# 조건 이름들
BUY_CONDITIONS = {"w1", "w2", "w3", "x1"}
SELL_CONDITIONS = {"w", "x2"}

# 리포트
REPORT_DIR = "reports"

# 주문 pending 만료(초)
PENDING_EXPIRE_SEC = 180

# 매수 쿨다운(초)
BUY_COOLDOWN_SEC = 180

# 손절 기준(%)
STOPLOSS_RATE = -10.0

# opt10001 레이트리밋(초) - TR 과부하 ret=-200 완화
PRICE_REQ_INTERVAL = 0.40  # 0.35~0.60 권장
