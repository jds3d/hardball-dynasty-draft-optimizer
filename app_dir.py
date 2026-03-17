"""
Resolve the application directory: project root when running as script,
directory containing the executable when run as a frozen (PyInstaller) exe.
Config, credentials, template, and outputs are expected in this directory
(or outputs/ subdir).
"""
import sys
from pathlib import Path


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_algorithm_file() -> Path:
    """Path to algorithm.json. When frozen, prefer file next to exe; else use bundled copy."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        user_algo = exe_dir / "algorithm.json"
        if user_algo.exists():
            return user_algo
        return Path(sys._MEIPASS) / "algorithm.json"
    return get_app_dir() / "algorithm.json"
