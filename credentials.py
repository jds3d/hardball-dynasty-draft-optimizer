"""
Load HBD login and scouting config from environment variables or a local file (never commit credentials).
"""
import os
from pathlib import Path

# File in project root; must be in .gitignore
CREDENTIALS_FILE = Path(__file__).resolve().parent / "credentials.env"


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


def get_scouting_config() -> dict[str, float]:
    """
    Return scouting budget and formula config from env vars or credentials.env.
    """
    env = _load_env_file()
    def _val(env_key: str, default: float) -> float:
        raw = os.environ.get(env_key, "").strip() or env.get(env_key, "")
        try:
            return float(raw)
        except (ValueError, TypeError):
            return default

    return {
        "college": _val("SCOUTING_COLLEGE", 0.0),
        "high_school": _val("SCOUTING_HIGH_SCHOOL", 0.0),
        "min_trust": _val("SCOUTING_MIN_TRUST", 0.10),
        "curve": _val("SCOUTING_CURVE", 0.17),
    }


def get_signability_config() -> dict[str, float]:
    """
    Return signability penalty multipliers from env vars or credentials.env.
    Each key maps a signability text category to its multiplier (0.0–1.0).
    Also includes overall thresholds for conditional penalties.
    """
    env = _load_env_file()
    def _val(env_key: str, default: float) -> float:
        raw = os.environ.get(env_key, "").strip() or env.get(env_key, "")
        try:
            return float(raw)
        except (ValueError, TypeError):
            return default

    return {
        "will_sign": _val("SIGN_WILL_SIGN", 1.0),
        "first_round": _val("SIGN_FIRST_ROUND", 0.90),
        "first_round_threshold": _val("SIGN_FIRST_ROUND_THRESHOLD", 70.0),
        "first_five": _val("SIGN_FIRST_FIVE", 0.80),
        "first_five_threshold": _val("SIGN_FIRST_FIVE_THRESHOLD", 60.0),
        "may_sign": _val("SIGN_MAY_SIGN", 0.60),
        "undecided": _val("SIGN_UNDECIDED", 0.40),
        "probably_wont": _val("SIGN_PROBABLY_WONT", 0.05),
        "unknown": _val("SIGN_UNKNOWN", 0.0),
        "fallback": _val("SIGN_FALLBACK", 0.50),
    }


def get_headless() -> bool:
    """Return whether to run the browser in headless mode."""
    env = _load_env_file()
    raw = os.environ.get("HEADLESS", "").strip() or env.get("HEADLESS", "")
    return raw.lower() in ("true", "1", "yes")
