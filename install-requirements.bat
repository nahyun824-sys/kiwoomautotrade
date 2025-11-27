@echo off
REM ===== 0) 이 배치파일이 있는 폴더로 이동 =====
cd /d "%~dp0"

REM ===== 1) 가상환경이 없으면 32비트 파이썬으로 생성 =====
IF NOT EXIST "venv\Scripts\python.exe" (
    echo [INFO] 32비트 파이썬으로 가상환경을 생성합니다...
    "C:\Python32\python.exe" -m venv venv
)

REM ===== 2) 가상환경 활성화 (반드시 call 사용) =====
call "%cd%\venv\Scripts\activate.bat"

REM ===== 3) 파이썬 아키텍처 및 실행 파일 경로 확인 =====
python -c "import platform, sys; print(platform.architecture(), sys.executable)"

REM ===== 4) requirements.txt 설치 (이 .bat과 같은 폴더에 있다고 가정) =====
    echo [INFO] requirements.txt를 통해 패키지를 설치합니다...
python -m pip install -r "%~dp0requirements.txt"

REM ===== 5) 결과 확인용 일시정지 =====
pause
