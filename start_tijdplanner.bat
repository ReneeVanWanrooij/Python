@echo off
setlocal

rem Start altijd vanuit de map waar dit .bat-bestand staat.
set "APP_DIR=%~dp0"
pushd "%APP_DIR%" >nul 2>&1

rem Gebruik bij voorkeur de Python Launcher op Windows.
where pyw >nul 2>&1
if %errorlevel%==0 (
    start "" pyw "%APP_DIR%pyside6_kalender_app.py"
    goto :done
)

where py >nul 2>&1
if %errorlevel%==0 (
    start "" py "%APP_DIR%pyside6_kalender_app.py"
    goto :done
)

rem Fallback: standaard python in PATH.
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "%APP_DIR%pyside6_kalender_app.py"
    goto :done
)

where python >nul 2>&1
if %errorlevel%==0 (
    start "" python "%APP_DIR%pyside6_kalender_app.py"
    goto :done
)

echo Python niet gevonden. Installeer Python 3 en probeer opnieuw.
pause

:done
popd >nul 2>&1
endlocal
exit /b
