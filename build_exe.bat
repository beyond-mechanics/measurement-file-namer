@echo off
chcp 65001 >nul
rem ==== 計測ファイル一括命名ツールを exe にする（1回だけ。普段は start_app.bat で十分） ====
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo [エラー] Python が見つかりません。
  pause
  exit /b 1
)

%PY% -m pip install pyinstaller tkinterdnd2 pywin32
if errorlevel 1 (
  echo [エラー] pip で部品を入れられませんでした。プロキシ設定を確認してください。
  pause
  exit /b 1
)

%PY% -m PyInstaller --noconfirm --onefile --noconsole --collect-all tkinterdnd2 --name SokuteiNamer sokutei_namer.py
if errorlevel 1 (
  echo [エラー] exe を作れませんでした。上のメッセージをスクショして送ってください。
  pause
  exit /b 1
)

echo.
echo dist\SokuteiNamer.exe ができました。これをダブルクリックすると起動します。
start "" "dist\SokuteiNamer.exe"
pause
