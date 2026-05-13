@echo off
setlocal

rem Start the Zotero PowerPoint picker from a stable repository-relative location.
rem User-facing output in this launcher is German by project convention.

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "REPO_ROOT=%%~fI"
set "APP_SCRIPT=%REPO_ROOT%\zotero_picker_ppt.py"
set "VENV_PYTHONW=%REPO_ROOT%\.venv\Scripts\pythonw.exe"
set "VENV_PYTHON=%REPO_ROOT%\.venv\Scripts\python.exe"

if not exist "%REPO_ROOT%" goto repo_missing
if not exist "%APP_SCRIPT%" goto script_missing

pushd "%REPO_ROOT%" >nul 2>nul
if errorlevel 1 goto repo_missing

if exist "%VENV_PYTHONW%" (
    set "PYTHON_EXE=%VENV_PYTHONW%"
    goto start_picker
)

if exist "%VENV_PYTHON%" (
    set "PYTHON_EXE=%VENV_PYTHON%"
    goto start_picker
)

where pyw.exe >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_EXE=pyw.exe"
    goto start_picker
)

where py.exe >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_EXE=py.exe"
    goto start_picker
)

goto python_missing

:start_picker
start "" "%PYTHON_EXE%" "%APP_SCRIPT%"
set "START_RESULT=%ERRORLEVEL%"
popd >nul 2>nul
if not "%START_RESULT%"=="0" goto start_failed
exit /b 0

:repo_missing
echo.
echo [zotero-ppt-picker] Fehler: Repository-Root konnte nicht gefunden werden.
echo Erwarteter Pfad: %REPO_ROOT%
goto error_exit

:script_missing
echo.
echo [zotero-ppt-picker] Fehler: zotero_picker_ppt.py wurde nicht gefunden.
echo Erwarteter Pfad: %APP_SCRIPT%
goto error_exit

:python_missing
popd >nul 2>nul
echo.
echo [zotero-ppt-picker] Fehler: Kein geeigneter Python-Starter gefunden.
echo Erwartet wird bevorzugt:
echo   %VENV_PYTHONW%
echo oder:
echo   %VENV_PYTHON%
echo Fallbacks: pyw.exe oder py.exe im PATH.
echo.
echo Bitte erstelle die virtuelle Umgebung und installiere die Abhaengigkeiten:
echo   py -m venv .venv
echo   .\.venv\Scripts\Activate.ps1
echo   pip install -r requirements.txt
goto error_exit

:start_failed
echo.
echo [zotero-ppt-picker] Fehler: Der Picker konnte nicht gestartet werden.
echo Python-Starter: %PYTHON_EXE%
echo Skript: %APP_SCRIPT%
goto error_exit

:error_exit
echo.
echo Dieses Fenster kann geschlossen werden, nachdem die Meldung gelesen wurde.
pause
exit /b 1
