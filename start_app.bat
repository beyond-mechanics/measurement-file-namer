@echo off
rem ==== 計測ファイル一括命名ツールを起動する（ダブルクリック用） ====
rem 追加の部品は不要です（Python 標準機能と Windows の PowerShell だけで動きます）。
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto NOPY

%PY% -c "import tkinter" >nul 2>nul
if errorlevel 1 goto NOTK

echo アプリを起動します。使っている間、この黒い画面は閉じないでください。
%PY% sokutei_namer.py
if errorlevel 1 goto CRASH
exit /b 0

:NOPY
echo [エラー] Python が見つかりません。
echo   Python を入れるとき「Add python.exe to PATH」にチェックを入れてください。
pause
exit /b 1

:NOTK
echo [エラー] この Python には tkinter が入っていません。
echo   Python のインストーラーで「tcl/tk and IDLE」にチェックを入れて入れ直してください。
pause
exit /b 1

:CRASH
echo.
echo [エラー] アプリが止まりました。この画面をスクショして送ってください。
echo   同じ内容が sokutei_namer_error.log にも残っています。
pause
exit /b 1
