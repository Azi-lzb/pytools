@echo off
setlocal
cd /d "%~dp0"

set "PY_EXE=%LocalAppData%\Programs\Python\Python310\python.exe"

if exist "%PY_EXE%" (
  "%PY_EXE%" -m PyInstaller ^
  --onefile ^
  --name pytools ^
  --noupx ^
  --clean ^
  --noconfirm ^
  --paths . ^
  --exclude-module matplotlib ^
  --exclude-module IPython ^
  --exclude-module jupyter ^
  --exclude-module notebook ^
  --exclude-module pytest ^
  --exclude-module setuptools ^
  --exclude-module tests ^
  pytools_entry.py
) else (
  py -3 -m PyInstaller ^
  --onefile ^
  --name pytools ^
  --noupx ^
  --clean ^
  --noconfirm ^
  --paths . ^
  --exclude-module matplotlib ^
  --exclude-module IPython ^
  --exclude-module jupyter ^
  --exclude-module notebook ^
  --exclude-module pytest ^
  --exclude-module setuptools ^
  --exclude-module tests ^
  pytools_entry.py
)

if not exist "dist\pytools.exe" (
  echo [WARN] dist\pytools.exe not found
  exit /b 1
)

set "TS_FILE=%TEMP%\pytools_build_ts.txt"
set "TS="
powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss" > "%TS_FILE%"
if exist "%TS_FILE%" (
  set /p TS=<"%TS_FILE%"
  del /q "%TS_FILE%" >nul 2>&1
)

if defined TS (
  copy /y "dist\pytools.exe" "dist\pytools_%TS%.exe" >nul
  echo [OK] built dist\pytools.exe
  echo [OK] built dist\pytools_%TS%.exe
) else (
  echo [OK] built dist\pytools.exe
  echo [WARN] timestamp suffix copy skipped
)

exit /b 0
