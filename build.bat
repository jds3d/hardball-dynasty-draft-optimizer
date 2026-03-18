@echo off
REM Build the Hardball Dynasty Draft Optimizer GUI executable.
REM Requires: .venv activated with pip install -r requirements.txt pyinstaller

echo Building HardballDraftOptimizer.exe in repo root ...
set PYTHON_EXE=python
if exist ".venv\Scripts\python.exe" set PYTHON_EXE=.venv\Scripts\python.exe
%PYTHON_EXE% -m PyInstaller --noconfirm --distpath . hardball_draft.spec
if %ERRORLEVEL% neq 0 exit /b %ERRORLEVEL%
echo.
echo Done. Executable: HardballDraftOptimizer.exe
echo Place credentials.env, config.json, and your Excel template next to the exe (or in the same folder).
echo outputs\ will be created next to the exe when you run Fetch.
