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
        print("   1. Confirm/submit your login (credentials may be pre-filled).")
        print("   2. Approve the PUSH NOTIFICATION on your phone when prompted.")
        print("   3. Wait until you reach the Applied Epic home screen.")
        print(f"\n You have {config.LOGIN_TIMEOUT_SECONDS} seconds.")
        print(" When you SEE that you are fully logged in, come back here")
        print(" and press ENTER to save the session.")
        print("=" * 70 + "\n")

        # Wait for the user to confirm they are logged in. We also keep the
        # browser responsive during the wait.
        try:
            input("Press ENTER once you are logged in and on the Epic home screen... ")
        except KeyboardInterrupt:
            log("Cancelled by user. No session saved.")
            browser.close()
            sys.exit(1)

        # Bring focus back to the MAIN Epic tab (login may have used a popup).
        main_page = context.pages[0]
        try:
            main_page.bring_to_front()
            # Reload so the SPA picks up the freshly stored auth token, then
            # give it time to land on the home screen and write localStorage.
            main_page.reload(wait_until="domcontentloaded")
            main_page.wait_for_timeout(5000)
        except Exception:
            pass

        # Verify we are actually logged in before saving (Locate toolbar shows
        # only when authenticated). Warn - but still save - if we cannot see it.
        logged_in = False
        try:
            logged_in = main_page.get_by_title("Locate (F2)").locator("i").is_visible(timeout=8000)
        except Exception:
            logged_in = False

        if logged_in:
            log("Confirmed: logged in (Locate toolbar is visible).")
        else:
            log("WARNING: could not confirm the logged-in home screen.")
            log("Saving anyway, but if the fetch script shows the login page,")
            log("re-run this login and make sure you reach the Epic home screen first.")

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
