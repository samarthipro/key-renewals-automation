"""
Shared configuration for the Applied Epic automation scripts.

Reads settings from a local .env file (see .env.example). The .env file is
git-ignored and must never be committed, because it contains your Applied
Epic username and password.
"""

from pathlib import Path
import os

try:
    from dotenv import load_dotenv
except ImportError:
    print("ERROR: Install python-dotenv using: pip install python-dotenv", flush=True)
    raise

# Load variables from a .env file located next to this script.
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def _require(name):
    value = os.getenv(name)
    if not value:
        raise RuntimeError(
            f"Missing required setting '{name}'. "
            f"Copy .env.example to .env and fill in your values."
        )
    return value


# Portal URL and credentials.
EPIC_URL = os.getenv("EPIC_URL", "https://liber23.appliedepic.com/#/")
EPIC_USERNAME = _require("EPIC_USERNAME")
EPIC_PASSWORD = _require("EPIC_PASSWORD")

# Where the authenticated session is stored (relative paths resolve to this
# script's folder so runs are consistent regardless of the working directory).
_storage_name = os.getenv("EPIC_STORAGE_STATE", "storage_state.json")
STORAGE_STATE = Path(_storage_name)
if not STORAGE_STATE.is_absolute():
    STORAGE_STATE = BASE_DIR / STORAGE_STATE

# Where downloaded documents are saved.
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", BASE_DIR / "downloads"))

# The client/account to search for in Epic's "Locate" screen.
CLIENT_NAME = os.getenv("CLIENT_NAME", "LOTUS HR")

# Behaviour toggles.
LOGIN_TIMEOUT_SECONDS = int(os.getenv("LOGIN_TIMEOUT_SECONDS", "300"))
HEADLESS = os.getenv("HEADLESS", "false").strip().lower() in ("1", "true", "yes")
