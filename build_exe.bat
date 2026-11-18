@echo off
rem 計測ファイル一括命名ツールを exe 化（初回のみ pip が必要）
pip install pyinstaller tkinterdnd2
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --name SokuteiNamer sokutei_namer.py
echo.
echo dist\SokuteiNamer.exe ができました。
pause
