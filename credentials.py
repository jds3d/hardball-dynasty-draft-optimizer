"""
Load HBD login from environment variables or a local file (never commit credentials).
"""
import os
from pathlib import Path

# File in project root; must be in .gitignore
CREDENTIALS_FILE = Path(__file__).resolve().parent / "credentials.env"


def get_hbd_credentials() -> tuple[str, str] | None:
    """
    Return (username, password) or None if not set.
    Reads from HBD_USERNAME / HBD_PASSWORD env vars first, then from credentials.env.
    """
    username = os.environ.get("HBD_USERNAME", "").strip()
    password = os.environ.get("HBD_PASSWORD", "").strip()
    if username and password:
        return (username, password)
    if not CREDENTIALS_FILE.exists():
        return None
    try:
        with open(CREDENTIALS_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, _, value = line.partition("=")
                    key, value = key.strip().upper(), value.strip().strip('"').strip("'")
                    if key == "USERNAME":
                        username = value
                    elif key == "PASSWORD":
                        password = value
        if username and password:
            return (username, password)
    except Exception:
        pass
    return None
