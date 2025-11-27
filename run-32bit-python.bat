@echo off
REM 1) 가상환경 활성화 (반드시 call 사용)
call "C:\kiwoomautotrade\venv\Scripts\activate.bat"

python -c "import platform; print(platform.architecture())"

cd C:\kiwoomautotrade

python kiwoomautotrade.py

REM 2) 32비트 환경이 적용된 cmd 창을 계속 열어두기
cmd /k