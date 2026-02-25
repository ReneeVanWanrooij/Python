@echo off
setlocal

set "SCRIPT=C:\Users\Rvwan\Downloads\pyside6_kalender_app.py"

if not exist "%SCRIPT%" (
  echo [FOUT] Script niet gevonden: %SCRIPT%
  pause
  exit /b 1
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
  start "" pythonw "%SCRIPT%"
  exit /b 0
)

where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw "%SCRIPT%"
  exit /b 0
)

where python >nul 2>nul
if %errorlevel%==0 (
  start "" python "%SCRIPT%"
  exit /b 0
)

where py >nul 2>nul
if %errorlevel%==0 (
  start "" py "%SCRIPT%"
  exit /b 0
)

echo [FOUT] Geen Python gevonden in PATH.
echo Installeer Python of voeg python.exe toe aan PATH.
pause
exit /b 1
