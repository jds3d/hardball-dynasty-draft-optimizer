@echo off
REM Build the Hardball Dynasty Draft Optimizer GUI executable.
REM Requires: .venv activated with pip install -r requirements.txt pyinstaller

echo Building HardballDraftOptimizer.exe in project root ...
set PYTHON_EXE=python
if exist ".venv\Scripts\python.exe" set PYTHON_EXE=.venv\Scripts\python.exe
%PYTHON_EXE% -m PyInstaller --noconfirm --distpath . hardball_draft.spec
if %ERRORLEVEL% neq 0 exit /b %ERRORLEVEL%
echo.
echo Done. Executable: HardballDraftOptimizer.exe (same folder as this project)
echo credentials.env, config.json, and the Excel template should already be here; see README.
echo outputs\ will be created next to the exe when you run Fetch.
