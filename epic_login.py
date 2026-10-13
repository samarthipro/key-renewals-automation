"""
STEP 1 (run once): Log in to Applied Epic and save the authenticated session.

What it does
------------
1. Opens Applied Epic in a VISIBLE browser window.
2. Pre-fills your username/password from .env (you can also type them yourself).
3. Waits for YOU to finish logging in and approve the push notification on
   your phone.
4. Once you are on the Epic home screen, it saves the browser session
   (cookies + local storage) to storage_state.json.

After this, epic_fetch_documents.py reuses that saved session so Epic treats
you as already logged in and does NOT ask for the push notification again -
until the session expires. When it expires, just run this script again.

Why a saved session instead of "saving the link"
-------------------------------------------------
A URL alone does not keep you logged in. The login lives in the browser
session (cookies/tokens). storage_state.json captures that session so it can
be replayed on later runs.

Run:
    python epic_login.py
"""

from datetime import datetime
import sys

import epic_config as config

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("ERROR: Install Playwright:\n"
          "    pip install playwright\n"
          "    python -m playwright install chromium", flush=True)
    raise


def log(message):
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def try_click_login_button(page):
    """
    The first Applied Epic page is a splash screen with a single 'Login'
    button (no credential fields yet). Click it to reach the sign-in page.

    Best effort: if the button is not found, you can click it yourself in the
    browser window.
    """
    login_button_selectors = [
        "button:has-text('Login')",
        "a:has-text('Login')",
        "input[value='Login']",
        "text=Login",
    ]
    for selector in login_button_selectors:
        try:
            button = page.locator(selector).first
            if button.is_visible(timeout=2000):
                button.click()
                log(f"Clicked the Login button using selector: {selector}")
                return True
        except Exception:
            continue
    log("Could not auto-click 'Login' - please click it yourself in the browser.")
    return False


def try_fill_login(page):
    """
    Best-effort attempt to pre-fill the username/password fields.

    Applied Epic login markup can vary. If these selectors do not match your
    tenant, it is harmless - you can simply type the credentials yourself in
    the opened browser window. Update the selectors below if you want full
    auto-fill.
    """
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
            if field.is_visible(timeout=1500):
                field.fill(config.EPIC_USERNAME)
                log(f"Pre-filled username using selector: {selector}")
                break
        except Exception:
            continue

    for selector in password_selectors:
        try:
            field = page.locator(selector).first
            if field.is_visible(timeout=1500):
                field.fill(config.EPIC_PASSWORD)
                log(f"Pre-filled password using selector: {selector}")
                break
        except Exception:
            continue


def main():
    log("Launching browser for Applied Epic login...")
    with sync_playwright() as p:
        # Always visible here: you need to see the page and approve the push.
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        log(f"Opening {config.EPIC_URL}")
        page.goto(config.EPIC_URL, wait_until="domcontentloaded")

        # The first page is a splash screen with a 'Login' button.
        page.wait_for_timeout(3000)
        try_click_login_button(page)

        # After clicking Login, the sign-in page (and its credential fields)
        # may load, sometimes in a new tab/window. Give it time, then try to
        # fill the username/password on whichever page is now active.
        page.wait_for_timeout(4000)

        # If clicking Login opened a new tab, switch to the newest page.
        if len(context.pages) > 1:
            page = context.pages[-1]
            log("Switched to the new sign-in tab.")

        try_fill_login(page)

        print("\n" + "=" * 70)
        print(" ACTION REQUIRED IN THE BROWSER WINDOW:")
        print("   1. Click 'Login' if it is still showing.")
        print("   2. Enter your credentials and approve the PUSH / Okta prompt.")
        print("   3. Keep going until you reach the Applied Epic HOME screen")
        print("      (the one with the 'Locate' button in the top toolbar).")
        print("\n DO NOT touch this terminal. The script will AUTOMATICALLY")
        print(" detect when you are logged in and then save the session.")
        print(f" It will wait up to {config.LOGIN_TIMEOUT_SECONDS} seconds.")
        print("=" * 70 + "\n")

        # Automatically poll until the HOME screen (Locate toolbar) appears on
        # ANY tab in the context. This replaces the fragile 'press ENTER' step,
        # which could fire early and save a logged-OUT session.
        import time
        deadline = time.time() + config.LOGIN_TIMEOUT_SECONDS
        logged_in = False
        main_page = context.pages[0]

        while time.time() < deadline:
            for pg in list(context.pages):
                try:
                    if pg.get_by_title("Locate (F2)").locator("i").is_visible(timeout=800):
                        main_page = pg
                        logged_in = True
                        break
                except Exception:
                    continue
            if logged_in:
                break
            log("Waiting for the Epic home screen (Locate toolbar)...")
            try:
                main_page.wait_for_timeout(3000)
            except Exception:
                time.sleep(3)

        if not logged_in:
            log("TIMED OUT waiting for the Epic home screen. Session NOT saved.")
            log("Re-run and make sure you reach the home screen (Locate toolbar)")
            log("before the timeout. You can increase LOGIN_TIMEOUT_SECONDS in .env.")
            browser.close()
            sys.exit(1)

        log("Confirmed: logged in (Locate toolbar is visible).")

        # Let the SPA finish writing auth state to localStorage before saving.
        try:
            main_page.bring_to_front()
            main_page.wait_for_timeout(3000)
        except Exception:
            pass

        # Save the authenticated session (cookies + localStorage) from the
        # context, which covers all tabs/pages.
        context.storage_state(path=str(config.STORAGE_STATE))
        log(f"Session saved to: {config.STORAGE_STATE}")
        log("You can now run: python epic_fetch_documents.py")

        browser.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
