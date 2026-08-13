@echo off
chcp 65001 > nul
echo ====================================================
echo  총무팀 자동화 허브 빌드 시작
echo ====================================================
echo.

cd /d "%~dp0"

echo [1/2] 필요 패키지 확인 중...
pip install pyinstaller tkinterdnd2 pandas openpyxl xlrd xlwt xlutils pywin32 keyring selenium --quiet
echo.

echo [2/2] PyInstaller 빌드 중...
pyinstaller --onefile --windowed ^
  --name "총무팀_자동화_허브" ^
  --add-data "mod_야근수당.py;." ^
  --add-data "mod_야근현황.py;." ^
  --add-data "mod_현장경비.py;." ^
  --add-data "mod_법률검토.py;." ^
  --add-data "law_api.py;." ^
  --add-data "law_config.py;." ^
  --add-data "야근수당_자동입력.py;." ^
  --add-data "mod_4대보험.py;." ^
  --add-data "mod_4대보험_parsers.py;." ^
  --add-data "mod_4대보험_matcher.py;." ^
  --add-data "mod_4대보험_report.py;." ^
  --add-data "mod_증명서발급.py;." ^
  --add-data "mod_계약서색인.py;." ^
  --add-data "mod_보증금현황.py;." ^
  --add-data "mod_전산기기.py;." ^
  --add-data "mod_근로내용신고.py;." ^
  --add-data "mod_근로내용신고_core.py;." ^
  --add-data "assets\근로내용확인신고_전자신고용 양식.xlsx;assets" ^
  --add-data "wehago_client.py;." ^
  --hidden-import tkinterdnd2 ^
  --hidden-import openpyxl ^
  --hidden-import xlrd ^
  --hidden-import xlwt ^
  --hidden-import xlutils ^
  --hidden-import xlutils.copy ^
  --hidden-import pandas ^
  --hidden-import win32com.client ^
  --hidden-import requests ^
  --hidden-import keyring.backends.Windows ^
  --collect-all selenium ^
  hub.py

echo.
if exist "dist\총무팀_자동화_허브.exe" (
    echo ====================================================
    echo  빌드 성공!
    echo  파일 위치: dist\총무팀_자동화_허브.exe
    echo ====================================================
    explorer dist
) else (
    echo ====================================================
    echo  빌드 실패 - 위 오류 메시지를 확인하세요
    echo ====================================================
)
pause
