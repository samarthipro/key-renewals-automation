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

# Where Playwright saves cookies/local storage as a portable backup.
_storage_name = os.getenv("EPIC_STORAGE_STATE", "storage_state.json")
STORAGE_STATE = Path(_storage_name)
if not STORAGE_STATE.is_absolute():
    STORAGE_STATE = BASE_DIR / STORAGE_STATE

# Applied Epic also uses browser-profile data and sessionStorage, which are not
# fully covered by Playwright's standard storage_state.json. Both login and
# fetch scripts reuse these paths to preserve the complete authenticated state.
_profile_name = os.getenv("EPIC_PROFILE_DIR", ".epic-browser-profile")
PROFILE_DIR = Path(_profile_name)
if not PROFILE_DIR.is_absolute():
    PROFILE_DIR = BASE_DIR / PROFILE_DIR

_session_name = os.getenv("EPIC_SESSION_STORAGE", "epic_session_storage.json")
SESSION_STORAGE = Path(_session_name)
if not SESSION_STORAGE.is_absolute():
    SESSION_STORAGE = BASE_DIR / SESSION_STORAGE

# Where downloaded documents are saved.
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", BASE_DIR / "downloads"))

# The client/account to search for in Epic's "Locate" screen.
CLIENT_NAME = os.getenv("CLIENT_NAME", "LOTUS HR")

# The lookup code of the account result row to open (from the Locate results).
LOOKUP_CODE = os.getenv("LOOKUP_CODE", "LOTUHRI-01")

# The policy code to open under the client's Policies list.
POLICY_CODE = os.getenv("POLICY_CODE", "KEYPEO00688800")

# Exact attachment name. Blank means every visible name ending in "Policy".
ATTACHMENT_NAME = os.getenv("ATTACHMENT_NAME", "")

# Use installed Google Chrome for Applied Epic compatibility. Set this blank
# only if Chrome is unavailable and Playwright's bundled Chromium is required.
BROWSER_CHANNEL = os.getenv("EPIC_BROWSER_CHANNEL", "chrome").strip() or None

# External normal-Chrome/CDP worker configuration.
CDP_PORT = int(os.getenv("EPIC_CDP_PORT", "9222"))
CHROME_PATH = os.getenv("EPIC_CHROME_PATH", "").strip()

# Fallback page coordinates for Epic's blue-toolbar Locate control. These are
# Playwright viewport coordinates (not Windows screen coordinates) and can be
# adjusted locally if the toolbar layout differs.
LOCATE_CLICK_X = int(os.getenv("EPIC_LOCATE_CLICK_X", "282"))
LOCATE_CLICK_Y = int(os.getenv("EPIC_LOCATE_CLICK_Y", "94"))

# Behaviour toggles.
LOGIN_TIMEOUT_SECONDS = int(os.getenv("LOGIN_TIMEOUT_SECONDS", "300"))
HEADLESS = os.getenv("HEADLESS", "false").strip().lower() in ("1", "true", "yes")
