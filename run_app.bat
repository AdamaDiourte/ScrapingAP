@echo off
setlocal
cd /d "%~dp0"

rem Lancer l'application (installe les deps si besoin, démarre le backend et ouvre le frontend)
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 start_all.py
) else (
  python start_all.py
)

rem Ouvrir l'URL HTTP servie par FastAPI (frontend monté)
start "" http://127.0.0.1:8000/

endlocal

