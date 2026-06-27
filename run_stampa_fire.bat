@echo off
rem -----------------------------------------------------------------
rem  STAMPA FIRE - Launcher
rem  Rileva Python, installa dipendenze e avvia l'applicazione.
rem -----------------------------------------------------------------

set APPDIR=%~dp0
cd /d "%APPDIR%"

rem --- Ricerca Python nel PATH ---
for /f "delims=" %%i in ('where python 2^>nul') do (
    set "PYTHON=%%i"
    goto :found
)

rem --- Fallback: percorso hardcoded ---
set "PYPATH=C:\Program Files\Python312"
if exist "%PYPATH%\python.exe" (
    set "PYTHON=%PYPATH%\python.exe"
    goto :found
)

echo.
echo [ERRORE] Python non trovato.
echo Installa Python 3.9+ da https://www.python.org e riprova.
echo.
pause
exit /b 1

:found
echo Python trovato: %PYTHON%

rem --- Installazione dipendenze ---
echo Verifica dipendenze (openpyxl, pillow)...
"%PYTHON%" -m pip install openpyxl pillow --quiet

rem --- Avvio applicazione ---
echo Avvio STAMPA FIRE...
"%PYTHON%" "%APPDIR%stampa_fire.py"

rem --- Pausa solo se c'e' stato un errore ---
if %errorlevel% neq 0 (
    echo.
    echo [ERRORE] L'applicazione si e' chiusa con errore (codice: %errorlevel%).
    pause
)
