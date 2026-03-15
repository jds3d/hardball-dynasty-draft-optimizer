"""
Load HBD login credentials from credentials.env and game config from config.json.
"""
import json
import os
from pathlib import Path

CREDENTIALS_FILE = Path(__file__).resolve().parent / "credentials.env"
CONFIG_FILE = Path(__file__).resolve().parent / "config.json"


def _load_env_file() -> dict[str, str]:
    """Parse credentials.env into a dict of uppercase keys."""
    result: dict[str, str] = {}
    if not CREDENTIALS_FILE.exists():
        return result
    try:
        with open(CREDENTIALS_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, _, value = line.partition("=")
                    result[key.strip().upper()] = value.strip().strip('"').strip("'")
    except Exception:
        pass
    return result


def _load_config_file() -> dict:
    """Load config.json. Returns empty dict if missing or invalid."""
    if not CONFIG_FILE.exists():
        return {}
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def get_hbd_credentials() -> tuple[str, str] | None:
    """
    Return (username, password) or None if not set.
    Reads from HBD_USERNAME / HBD_PASSWORD env vars first, then from credentials.env.
    """
    username = os.environ.get("HBD_USERNAME", "").strip()
    password = os.environ.get("HBD_PASSWORD", "").strip()
    if username and password:
        return (username, password)
    env = _load_env_file()
    username = env.get("USERNAME", "")
    password = env.get("PASSWORD", "")
    if username and password:
        return (username, password)
    return None


def get_headless() -> bool:
    """Return whether to run the browser in headless mode."""
    env = _load_env_file()
    raw = os.environ.get("HEADLESS", "").strip() or env.get("HEADLESS", "")
    return raw.lower() in ("true", "1", "yes")


def get_scouting_config() -> dict[str, float]:
    """Return scouting budget and trust formula config from config.json."""
    cfg = _load_config_file().get("scouting", {})
    return {
        "college": float(cfg.get("college", 0.0)),
        "high_school": float(cfg.get("high_school", 0.0)),
        "min_trust": float(cfg.get("min_trust", 0.10)),
        "curve": float(cfg.get("curve", 0.17)),
    }


def get_signability_config() -> dict[str, float]:
    """Return signability penalty multipliers from config.json."""
    cfg = _load_config_file().get("signability", {})
    return {
        "will_sign": float(cfg.get("will_sign", 1.0)),
        "first_round": float(cfg.get("first_round", 0.90)),
        "first_round_threshold": float(cfg.get("first_round_threshold", 70.0)),
        "first_five": float(cfg.get("first_five", 0.80)),
        "first_five_threshold": float(cfg.get("first_five_threshold", 60.0)),
        "may_sign": float(cfg.get("may_sign", 0.60)),
        "undecided": float(cfg.get("undecided", 0.40)),
        "probably_wont": float(cfg.get("probably_wont", 0.05)),
        "unknown": float(cfg.get("unknown", 0.0)),
        "fallback": float(cfg.get("fallback", 0.50)),
    }
