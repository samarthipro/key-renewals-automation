"""
Log in to Applied Epic once and preserve the authenticated browser state.

Applied Epic/Okta does not keep all authentication data in ordinary cookies.
This script therefore preserves three layers after a successful login:
  1. A persistent Chromium profile (cookies, IndexedDB, service-worker data).
  2. Playwright storage_state.json (cookies, localStorage, IndexedDB backup).
  3. sessionStorage values in epic_session_storage.json.

Run this script again whenever Epic/Okta expires the session:
    python epic_login.py
"""

from datetime import datetime
import json
import sys
import time

import epic_config as config

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print(
        "ERROR: Install Playwright:\n"
        "    pip install playwright\n"
        "    python -m playwright install chromium",
        flush=True,
    )
    raise


def log(message):
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def try_click_login_button(page):
    """Best-effort click of the Applied Epic splash-screen Login button."""
    selectors = [
        "button:has-text('Login')",
        "a:has-text('Login')",
        "input[value='Login']",
        "text=Login",
    ]
    for selector in selectors:
        try:
            button = page.locator(selector).first
            if button.is_visible(timeout=2000):
                button.click()
                log("Clicked the Applied Epic Login button.")
                return True
        except Exception:
            continue
    log("Could not auto-click Login; click it manually in the browser.")
    return False


def try_fill_login(page):
    """Best-effort credential fill; manual login remains available."""
    username_selectors = [
        "input[type='email']",
        "input[name='username']",
        "input[name='email']",
        "input[id*='user' i]",
        "input[autocomplete='username']",
    ]
    password_selectors = [
        "input[type='password']",
        "input[name='password']",
        "input[id*='pass' i]",
        "input[autocomplete='current-password']",
    ]

    for selector in username_selectors:
        try:
            field = page.locator(selector).first
            if field.is_visible(timeout=1000):
                field.fill(config.EPIC_USERNAME)
                log("Pre-filled the username.")
                break
        except Exception:
            continue

    for selector in password_selectors:
        try:
            field = page.locator(selector).first
            if field.is_visible(timeout=1000):
                field.fill(config.EPIC_PASSWORD)
                log("Pre-filled the password.")
                break
        except Exception:
            continue


def concurrent_login_warning_visible(page):
    """Return True when Epic asks whether to take over another login."""
    try:
        yes_buttons = page.get_by_role("button", name="Yes", exact=True)
        return any(
            yes_buttons.nth(index).is_visible(timeout=300)
            for index in range(yes_buttons.count())
        )
    except Exception:
        return False


def find_logged_in_page(context):
    """Return an unlocked page containing Epic's Locate toolbar, or None."""
    for page in list(context.pages):
        try:
            # A visible Yes means another-device takeover is awaiting the
            # user's decision. Do not automate that decision or start work yet.
            if concurrent_login_warning_visible(page):
                continue
            locate = page.get_by_title("Locate (F2)").locator("i")
            if locate.is_visible(timeout=800):
                return page
        except Exception:
            continue
    return None


def save_session_storage(context):
    """Save sessionStorage for every currently open HTTP(S) origin."""
    saved = {}
    for page in list(context.pages):
        try:
            if not page.url.startswith(("http://", "https://")):
                continue
            origin = page.evaluate("window.location.origin")
            values = page.evaluate(
                """() => {
                    const result = {};
                    for (let i = 0; i < sessionStorage.length; i++) {
                        const key = sessionStorage.key(i);
                        result[key] = sessionStorage.getItem(key);
                    }
                    return result;
                }"""
            )
            if values:
                saved[origin] = values
        except Exception:
            continue

    config.SESSION_STORAGE.write_text(
        json.dumps(saved, indent=2), encoding="utf-8"
    )
    log(
        f"Saved sessionStorage for {len(saved)} origin(s): "
        f"{config.SESSION_STORAGE}"
    )


def save_storage_state(context):
    """Save cookies/localStorage and IndexedDB where supported."""
    try:
        context.storage_state(path=str(config.STORAGE_STATE), indexed_db=True)
    except TypeError:
        # Compatibility with older Playwright versions.
        context.storage_state(path=str(config.STORAGE_STATE))
    log(f"Saved Playwright storage state: {config.STORAGE_STATE}")


def main():
    config.PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    log("Launching the persistent Applied Epic browser profile...")

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(config.PROFILE_DIR),
            headless=False,
            channel=config.BROWSER_CHANNEL,
            args=["--start-maximized"],
            no_viewport=True,
            accept_downloads=True,
        )

        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_navigation_timeout(120000)
        page.set_default_timeout(30000)
        log(f"Opening {config.EPIC_URL}")
        # Applied Epic is a slow SPA and may not fire DOMContentLoaded within
        # Playwright's 30-second default. "commit" confirms the server response;
        # the logged-in toolbar polling below handles the remaining UI load.
        page.goto(config.EPIC_URL, wait_until="commit", timeout=120000)
        page.wait_for_timeout(5000)

        # If the persistent profile is already authenticated, do not trigger a
        # new login. Otherwise click Login and let the user finish Okta/MFA.
        logged_in_page = find_logged_in_page(context)
        if logged_in_page is None:
            try_click_login_button(page)
            page.wait_for_timeout(4000)
            sign_in_page = context.pages[-1]
            if sign_in_page is not page:
                log("Switched to the Okta sign-in popup.")
            try_fill_login(sign_in_page)

            print("\n" + "=" * 70)
            print(" COMPLETE LOGIN IN THE BROWSER:")
            print("   1. Enter/confirm your username and password.")
            print("   2. Complete the authorized Okta push/code verification.")
            print("   3. Wait for the Epic home screen with the Locate toolbar.")
            print("\n Do not press ENTER here; detection is automatic.")
            print(
                f" The script will wait up to {config.LOGIN_TIMEOUT_SECONDS} seconds."
            )
            print("=" * 70 + "\n")

            deadline = time.time() + config.LOGIN_TIMEOUT_SECONDS
            takeover_warning_announced = False
            while time.time() < deadline:
                # If this account is active elsewhere, Epic asks the user
                # whether to take over. Leave that decision to the user.
                for candidate_page in list(context.pages):
                    if concurrent_login_warning_visible(candidate_page):
                        candidate_page.bring_to_front()
                        if not takeover_warning_announced:
                            log(
                                "Epic reports this account is active elsewhere. "
                                "Choose Yes or No manually in the browser."
                            )
                            takeover_warning_announced = True
                logged_in_page = find_logged_in_page(context)
                if logged_in_page is not None:
                    break
                log("Waiting for the Epic home screen...")
                time.sleep(3)

        if logged_in_page is None:
            log("TIMED OUT. No authenticated state was saved.")
            context.close()
            sys.exit(1)

        log("Confirmed: logged in and no takeover decision is pending.")
        logged_in_page.bring_to_front()
        logged_in_page.wait_for_timeout(3000)

        save_session_storage(context)
        save_storage_state(context)
        log(f"Persistent browser profile saved at: {config.PROFILE_DIR}")
        log("You can now run: python epic_fetch_documents.py")
        context.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
