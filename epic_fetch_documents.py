"""
Fetch policy documents from Applied Epic using the authenticated browser state.

Authentication is restored from the persistent Chromium profile created by
`epic_login.py`, with an explicit sessionStorage restoration because Applied
Epic/Okta does not keep all login state in ordinary Playwright storage state.

Navigation flow:
    Locate -> client -> Policies -> policy -> Access -> Attachments -> Download

Run:
    python epic_fetch_documents.py
"""

from datetime import datetime
import json
import re
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


def install_session_storage(context):
    """Restore saved sessionStorage before Applied Epic loads."""
    if not config.SESSION_STORAGE.exists():
        raise FileNotFoundError(
            f"Missing {config.SESSION_STORAGE}. Run 'python epic_login.py' "
            "again to capture the complete authenticated state."
        )

    state = json.loads(config.SESSION_STORAGE.read_text(encoding="utf-8"))
    encoded_state = json.dumps(state).replace("</", "<\\/")
    context.add_init_script(
        script=f"""
        (() => {{
            const allState = {encoded_state};
            const values = allState[window.location.origin];
            if (!values) return;
            for (const [key, value] of Object.entries(values)) {{
                window.sessionStorage.setItem(key, value);
            }}
        }})();
        """
    )
    log(f"Prepared sessionStorage for {len(state)} origin(s).")


def session_is_valid(page):
    """Require a stable Locate toolbar before treating the session as valid."""
    deadline = time.time() + 30
    while time.time() < deadline:
        try:
            login_button = page.get_by_role("button", name="Login")
            if login_button.is_visible(timeout=500):
                return False
        except Exception:
            pass

        try:
            locate = page.get_by_title("Locate (F2)").locator("i")
            if locate.is_visible(timeout=800):
                # Ensure this was not a transient home screen immediately before
                # an authentication redirect.
                page.wait_for_timeout(3000)
                return locate.is_visible(timeout=1000)
        except Exception:
            pass

        page.wait_for_timeout(1000)

    return False


def save_download(download, index):
    """Save a download into DOWNLOAD_DIR without overwriting an older file."""
    suggested = download.suggested_filename or f"document_{index}"
    target = config.DOWNLOAD_DIR / suggested
    counter = 1
    while target.exists():
        target = config.DOWNLOAD_DIR / f"{target.stem}_{counter}{target.suffix}"
        counter += 1
    download.save_as(str(target))
    log(f"Downloaded: {target.name}")
    return target


def locate_client(page):
    """Locate the configured client and open its matching account row."""
    client = config.CLIENT_NAME
    log(f"Locating client: {client!r}")

    page.get_by_title("Locate (F2)").locator("i").click(timeout=15000)
    page.wait_for_timeout(1500)

    search_box = page.get_by_role("textbox").first
    search_box.click()
    search_box.fill(client)
    search_box.press("Enter")
    log("Submitted the client search.")
    page.wait_for_timeout(3000)

    lookup_code = config.LOOKUP_CODE
    opened = False
    if lookup_code:
        try:
            row = page.locator("div").filter(
                has_text=re.compile(rf"^{re.escape(lookup_code)}$")
            ).nth(1)
            row.click(timeout=15000)
            log(f"Opened account by lookup code: {lookup_code}")
            opened = True
        except Exception as error:
            log(f"Could not open by lookup code ({error}); trying client name.")

    if not opened:
        page.get_by_text(client).first.click(timeout=15000)
        log(f"Opened account by name: {client}")

    page.wait_for_timeout(3000)


def open_policy_attachments(page):
    """Open Policies, choose the configured policy, then open Attachments."""
    page.locator("a").filter(has_text="Policies").click(timeout=15000)
    log("Opened the Policies section.")
    page.wait_for_timeout(3000)

    policy_code = config.POLICY_CODE
    page.get_by_text(policy_code).click(timeout=15000)
    log(f"Opened policy: {policy_code}")
    page.wait_for_timeout(2000)

    page.get_by_title("Access (None)").locator("i").click(timeout=15000)
    page.wait_for_timeout(1000)
    page.get_by_text("Attachments").first.click(timeout=15000)
    log("Opened the policy's Attachments.")
    page.wait_for_timeout(3000)


def download_attachments(page):
    """Download one configured attachment, or attempt every listed row."""
    downloaded = []
    target_name = config.ATTACHMENT_NAME.strip()

    def download_one(name):
        try:
            page.get_by_text(name).first.dblclick(timeout=15000)
            page.wait_for_timeout(1500)
            with page.expect_download(timeout=30000) as download_info:
                page.get_by_role("button", name="Download").click()
            downloaded.append(save_download(download_info.value, len(downloaded) + 1))
            return True
        except Exception as error:
            log(f"Could not download {name!r}: {error}")
            return False

    if target_name:
        log(f"Downloading specific attachment: {target_name!r}")
        download_one(target_name)
        return downloaded

    log("No ATTACHMENT_NAME set; attempting all detected attachment rows.")
    try:
        rows = page.get_by_role("row")
        count = rows.count()
    except Exception:
        count = 0

    names = []
    for index in range(count):
        try:
            text = rows.nth(index).inner_text(timeout=1000).strip()
            first_line = text.splitlines()[0].strip() if text else ""
            if first_line and first_line not in names:
                names.append(first_line)
        except Exception:
            continue

    if not names:
        log("No attachment rows detected. Set ATTACHMENT_NAME in .env.")
    for name in names:
        download_one(name)
    return downloaded


def save_current_session_storage(context):
    """Refresh saved sessionStorage after a successful run."""
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
    config.SESSION_STORAGE.write_text(json.dumps(saved, indent=2), encoding="utf-8")


def navigate_and_download(page):
    locate_client(page)
    open_policy_attachments(page)
    downloaded = download_attachments(page)
    if not downloaded:
        log("No documents were downloaded.")
    return downloaded


def main():
    if not config.PROFILE_DIR.exists() or not config.SESSION_STORAGE.exists():
        raise FileNotFoundError(
            "Complete authenticated browser state was not found. "
            "Run 'python epic_login.py' with the updated scripts first."
        )

    config.DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Documents will be saved to: {config.DOWNLOAD_DIR}")

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(config.PROFILE_DIR),
            headless=config.HEADLESS,
            accept_downloads=True,
        )
        install_session_storage(context)

        page = context.pages[0] if context.pages else context.new_page()
        log("Opening Applied Epic with the complete saved browser state...")
        page.goto(config.EPIC_URL, wait_until="domcontentloaded")

        if not session_is_valid(page):
            context.close()
            raise RuntimeError(
                "Applied Epic still shows the Login page. The Okta/Epic session "
                "has expired or was not fully captured. Close all automation "
                "browser windows, run 'python epic_login.py', and then run this "
                "script again."
            )

        log("Confirmed authenticated home screen. Navigating to documents...")
        downloaded = navigate_and_download(page)
        save_current_session_storage(context)

        try:
            context.storage_state(path=str(config.STORAGE_STATE), indexed_db=True)
        except TypeError:
            context.storage_state(path=str(config.STORAGE_STATE))

        log(f"Done. {len(downloaded)} document(s) downloaded.")
        context.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
