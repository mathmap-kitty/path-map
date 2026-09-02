@echo off
REM 每週執行一次：重抓運動績優單獨招生資訊，有變動就重建 ../sport.html
REM 由 Windows 工作排程器「path-map 體育班單招每週檢核」呼叫，也可以自己雙擊執行。
REM 每次執行的摘要會寫進 _build\sport_watch_log.md。

setlocal
set PY=C:\Users\User\AppData\Local\Programs\Python\Python313\python.exe
set DIR=%~dp0
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

cd /d "%DIR%"
"%PY%" fetch_sport.py >> "%DIR%sport_watch_stdout.log" 2>&1
endlocal
