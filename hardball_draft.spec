# PyInstaller spec for Hardball Dynasty Draft Optimizer (GUI executable).
# Build: pyinstaller hardball_draft.spec
# Output: dist/HardballDraftOptimizer.exe (or dist\HardballDraftOptimizer.exe on Windows)

import sys

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# algorithm.json is bundled so the exe works without extra files; user can override by placing algorithm.json next to the exe.
added_files = [('algorithm.json', '.')]

# Selenium and webdriver_manager use lazy/dynamic imports; collect all submodules so the exe finds them at runtime.
try:
    _selenium_hidden = collect_submodules('selenium')
    _webdriver_mgr_hidden = collect_submodules('webdriver_manager')
except Exception:
    _selenium_hidden = _webdriver_mgr_hidden = []
# Explicit fallback so the exe always gets these (often missed by dynamic import)
_selenium_fallback = [
    'selenium.webdriver.chrome.webdriver',
    'selenium.webdriver.chrome.service',
    'selenium.webdriver.chrome.options',
    'selenium.webdriver.chrome.remote_connection',
]
_webdriver_mgr_fallback = ['webdriver_manager.chrome', 'webdriver_manager.core']
_all_hidden = (
    ['credentials', 'excel_draft', 'web_draft', 'app_dir', 'openpyxl', 'pandas']
    + list(_selenium_hidden) + list(_webdriver_mgr_hidden)
    + _selenium_fallback + _webdriver_mgr_fallback
)

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=_all_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='HardballDraftOptimizer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Set to 'icon.ico' if you add an icon
)
