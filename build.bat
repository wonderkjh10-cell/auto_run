@echo off
chcp 65001 > nul
echo ====================================
echo  발주서 처리 프로그램 빌드 시작
echo ====================================
echo.

pip install pyinstaller openpyxl pandas tkinterdnd2 -q

echo.
echo [빌드 중...] 잠시 기다려주세요.
echo.

pyinstaller --onefile --windowed --name "발주서처리프로그램" main.py

echo.
echo ====================================
echo  완료! dist 폴더를 확인하세요.
echo  발주서처리프로그램.exe 가 생성되었습니다.
echo ====================================
pause
