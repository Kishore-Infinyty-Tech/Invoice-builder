@echo off
setlocal
cd /d "%~dp0"
title Invoice Studio
echo.
echo   Invoice Studio
echo   ----------------------------------------------

set "PY="
for %%C in (python py) do (
  if not defined PY (
    %%C -c "import sys;sys.exit(sys.version_info<(3,8))" >nul 2>nul && set "PY=%%C"
  )
)

if not defined PY (
  echo.
  echo   Python is needed once on this computer, and it was not found.
  echo.
  echo     1. Open  https://www.python.org/downloads/
  echo     2. Download and install Python 3  ^(tick "Add python.exe to PATH"^)
  echo     3. Double-click this file again
  echo.
  pause
  exit /b 1
)

%PY% -c "import openpyxl, PIL" >nul 2>nul
if errorlevel 1 (
  echo   First start: installing two small helper packages...
  %PY% -m pip install --disable-pip-version-check -q -r requirements.txt
  if errorlevel 1 (
    echo.
    echo   Could not install them. Check the internet connection and try again.
    echo.
    pause
    exit /b 1
  )
)

echo   Starting... your browser will open by itself.
echo   Keep this window open while you work. Close it to stop the tool.
echo.
%PY% app.py
echo.
echo   The tool has stopped.
pause
