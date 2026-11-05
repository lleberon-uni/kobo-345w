@echo off
cd /d "%~dp0.."
set "ETL_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%ETL_PYTHON%" (
    "%ETL_PYTHON%" -c "import openpyxl, tkinter" >nul 2>nul
    if not errorlevel 1 goto bundled
)
py -c "import openpyxl, tkinter" >nul 2>nul
if not errorlevel 1 goto installed
echo No se encontro un Python con openpyxl y tkinter instalados.
echo Instale las dependencias con: py -m pip install -r ETL\requirements.txt
pause
exit /b 1
:bundled
"%ETL_PYTHON%" ETL\aplicacion.py
goto done
:installed
py ETL\aplicacion.py
:done
if errorlevel 1 pause
