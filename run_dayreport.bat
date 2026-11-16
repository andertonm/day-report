@echo off
REM Day Report nightly run - replaces the old batch file + dayreport.exe.
REM Task Scheduler calls this file. It never opens Excel and shows no windows/dialogs.
setlocal
cd /d "%~dp0"
REM Python installed "for all users" by the setup steps in README.txt
REM Falls back to a per-user Python install (runs fine when the task runs as the same Windows user)
set PY="C:\Program Files\Python312\python.exe"
if not exist %PY% if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist %PY% if exist "C:\Users\Mark\AppData\Local\Programs\Python\Python312\python.exe" set PY="C:\Users\Mark\AppData\Local\Programs\Python\Python312\python.exe"
if not exist %PY% set PY=py -3
if exist "%~dp0venv\Scripts\python.exe" set PY="%~dp0venv\Scripts\python.exe"
%PY% "%~dp0dayreport.py" --config "%~dp0dayreport.ini" %*
set RC=%ERRORLEVEL%
if not "%RC%"=="0" echo %date% %time% dayreport exit code %RC% >> "%~dp0last_error.txt"
exit /b %RC%
