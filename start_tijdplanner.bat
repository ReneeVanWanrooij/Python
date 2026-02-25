@echo off
setlocal

rem Start altijd vanuit de map waar dit .bat-bestand staat.
set "APP_DIR=%~dp0"
pushd "%APP_DIR%" >nul 2>&1
set "SCRIPT=%APP_DIR%pyside6_kalender_app.py"

if not exist "%SCRIPT%" (
    echo [FOUT] Script niet gevonden:
    echo        %SCRIPT%
    pause
    goto :done
)

set "RUNNER="
where py >nul 2>&1
if %errorlevel%==0 set "RUNNER=py"

if not defined RUNNER (
    where python >nul 2>&1
    if %errorlevel%==0 set "RUNNER=python"
)

if not defined RUNNER (
    echo [FOUT] Python niet gevonden in PATH.
    echo        Installeer Python 3 of voeg Python toe aan PATH.
    pause
    goto :done
)

echo [INFO] Start Tijdplanner...
echo [INFO] Map: %APP_DIR%
echo [INFO] Runner: %RUNNER%
echo.

call %RUNNER% "%SCRIPT%"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [FOUT] Tijdplanner stopte met exitcode %RC%.
    echo        Controleer de foutmelding hierboven.
    pause
    goto :done
)

:done
popd >nul 2>&1
endlocal
exit /b
