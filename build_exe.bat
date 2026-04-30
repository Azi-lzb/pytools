@echo off
REM Minimal one-file build for pytools.
REM Output: dist\pytools.exe
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
