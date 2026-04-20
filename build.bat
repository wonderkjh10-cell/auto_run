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

pyinstaller --onefile --windowed --name "발주서처리프로그램" --add-data "version.txt;." main.py

echo.
echo [압축 중...]
set /p VERSION=<version.txt
python -c "import zipfile, os; v=open('version.txt').read().strip(); z=zipfile.ZipFile(f'dist/발주서처리프로그램_v{v}.zip','w',zipfile.ZIP_DEFLATED); z.write('dist/발주서처리프로그램.exe','발주서처리프로그램.exe'); z.close(); print(f'ZIP 생성: 발주서처리프로그램_v{v}.zip')"

echo.
echo ====================================
echo  완료! dist 폴더를 확인하세요.
echo  발주서처리프로그램.exe
echo  발주서처리프로그램_v%VERSION%.zip
echo ====================================
pause
