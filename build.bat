@echo off
REM Build the Hardball Dynasty Draft Optimizer GUI executable.
REM Requires: .venv activated with pip install -r requirements.txt pyinstaller

echo Building HardballDraftOptimizer.exe ...
pyinstaller --noconfirm hardball_draft.spec
if %ERRORLEVEL% neq 0 exit /b %ERRORLEVEL%
echo.
echo Done. Executable: dist\HardballDraftOptimizer.exe
echo Place credentials.env, config.json, and your Excel template next to the exe (or in the same folder).
echo outputs\ will be created next to the exe when you run Fetch.
