@echo off
chcp 65001 >nul
rem ==== 計測ファイル一括命名ツールを起動する（ダブルクリック用） ====
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto NOPY

%PY% -c "import tkinter" >nul 2>nul || goto NOTK

%PY% -c "import win32com.client, tkinterdnd2" >nul 2>nul
if errorlevel 1 (
  echo 初回のみ必要な部品 pywin32 と tkinterdnd2 を入れます...
  %PY% -m pip install pywin32 tkinterdnd2
  if errorlevel 1 goto PIPFAIL
)

echo アプリを起動します。この黒い画面は閉じないでください。
%PY% sokutei_namer.py
if errorlevel 1 goto CRASH
exit /b 0

:NOPY
echo [エラー] Python が見つかりません。
echo   python.org から Python を入れるときに "Add python.exe to PATH" にチェックを入れてください。
pause
exit /b 1

:NOTK
echo [エラー] この Python には tkinter が入っていません。
echo   Python のインストーラーで "tcl/tk and IDLE" にチェックを入れて入れ直してください。
pause
exit /b 1

:PIPFAIL
echo [エラー] pip で部品を入れられませんでした。社内ネットワークのプロキシが原因のことが多いです。
echo   例: %PY% -m pip install --proxy http://プロキシ:ポート pywin32 tkinterdnd2
pause
exit /b 1

:CRASH
echo.
echo [エラー] アプリが止まりました。上の英語のメッセージをスクショして送ってください。
echo   同じ内容が sokutei_namer_error.log にも残っています。
pause
exit /b 1
