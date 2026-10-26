@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  [전주 교육계획 등록 시험 모드 테스트]
echo  - 등록/저장은 하지 않고 화면 입력까지만 진행됩니다.
echo ============================================================
echo.
"venv\Scripts\python.exe" test_run_jeonju.py
echo.
pause
