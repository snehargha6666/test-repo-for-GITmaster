@echo off
set "SCRIPT=%~dp0mdb_to_db.py"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%SCRIPT%" --gui
  goto end
)

where python >nul 2>nul
if %errorlevel%==0 (
  python "%SCRIPT%" --gui
  goto end
)

if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
  "%LocalAppData%\Programs\Python\Python313\python.exe" "%SCRIPT%" --gui
  goto end
)

echo Python 3 was not found. Install Python 3, then run this file again.

:end
pause
