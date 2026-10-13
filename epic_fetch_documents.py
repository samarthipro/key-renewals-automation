"""
STEP 2: Fetch policy documents from Applied Epic using the saved session.

What it does
------------
1. Loads the authenticated session saved by epic_login.py (storage_state.json),
   so Epic treats you as already logged in - no push notification needed.
2. Opens Epic and verifies the session is still valid (not expired).
3. Follows the recorded navigation flow:
      Locate -> search client -> open account -> Policies -> open policy
      -> Access -> Attachments -> open attachment -> Download
4. Saves downloaded files into DOWNLOAD_DIR.

If the saved session has EXPIRED, the script tells you to re-run epic_login.py.

Configurable values (via .env):
    CLIENT_NAME     - client to search for         (e.g. LOTUS HR)
    LOOKUP_CODE     - account result row to open    (e.g. LOTUHRI-01)
    POLICY_CODE     - policy to open                (e.g. KEYPEO00688800)
    ATTACHMENT_NAME - specific attachment to grab; blank = download all

Run:
    python epic_fetch_documents.py
"""

from datetime import datetime
import re
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


def session_is_valid(page):
    """Return False if Epic bounced us back to a login screen (expired session)."""
    login_markers = [
        "input[type='password']",
        "button:has-text('Login')",
    ]
    for selector in login_markers:
        try:
            if page.locator(selector).first.is_visible(timeout=2000):
                return False
        except Exception:
            continue
    return True


def save_download(download, index):
    """Save a Playwright download into DOWNLOAD_DIR, preserving its name."""
    suggested = download.suggested_filename or f"document_{index}"
    target = config.DOWNLOAD_DIR / suggested

    # Avoid overwriting: add a counter if a file with the same name exists.
    counter = 1
    while target.exists():
        target = config.DOWNLOAD_DIR / f"{target.stem}_{counter}{target.suffix}"
        counter += 1

    download.save_as(str(target))
    log(f"Downloaded: {target.name}")
    return target


def locate_client(page):
    """
    Locate -> search client by name -> open the matching account row.

    Selectors come from a recorded Playwright session of the real flow.
    """
    client = config.CLIENT_NAME
    log(f"Locating client: {client!r}")

    # Open the Locate screen (toolbar button titled 'Locate (F2)').
    page.get_by_title("Locate (F2)").locator("i").click()
    page.wait_for_timeout(1500)

    # Type the client name into the first textbox and press Enter.
    search_box = page.get_by_role("textbox").first
    search_box.click()
    search_box.fill(client)
    search_box.press("Enter")
    log("Submitted the client search.")
    page.wait_for_timeout(3000)

    # Open the matching result row - prefer the lookup code, else the name.
    lookup_code = config.LOOKUP_CODE
    opened = False
    if lookup_code:
        try:
            row = page.locator("div").filter(
                has_text=re.compile(rf"^{re.escape(lookup_code)}$")
            ).nth(1)
            row.click()
            log(f"Opened account by lookup code: {lookup_code}")
            opened = True
        except Exception as exc:
            log(f"Could not open by lookup code ({exc}); trying client name.")

    if not opened:
        page.get_by_text(client).first.click()
        log(f"Opened account by name: {client}")

    page.wait_for_timeout(3000)


def open_policy_attachments(page):
    """
    Policies -> open policy -> Access -> Attachments.

    Reaches the attachment list for the configured policy.
    """
    # Click the 'Policies' link in the account sidebar.
    page.locator("a").filter(has_text="Policies").click()
    log("Opened the Policies section.")
    page.wait_for_timeout(3000)

    # Open the specific policy row by its code.
    policy_code = config.POLICY_CODE
    page.get_by_text(policy_code).click()
    log(f"Opened policy: {policy_code}")
    page.wait_for_timeout(2000)

    # Open the 'Access' menu and go to 'Attachments'.
    page.get_by_title("Access (None)").locator("i").click()
    page.wait_for_timeout(1000)
    page.get_by_text("Attachments").first.click()
    log("Opened the policy's Attachments.")
    page.wait_for_timeout(3000)


def download_attachments(page):
    """
    Download attachment(s) from the policy's Attachments list.

    If ATTACHMENT_NAME is set, download just that one; otherwise download
    every attachment row found. Each download uses: open the attachment,
    then click the 'Download' button.
    """
    downloaded = []
    target_name = config.ATTACHMENT_NAME.strip()

    def download_one(name):
        """Open an attachment by name and click Download; save the file."""
        try:
            page.get_by_text(name).first.dblclick()
            page.wait_for_timeout(1500)
            with page.expect_download(timeout=30000) as dl_info:
                page.get_by_role("button", name="Download").click()
            saved = save_download(dl_info.value, len(downloaded) + 1)
            downloaded.append(saved)
            return True
        except Exception as exc:
            log(f"Could not download {name!r}: {exc}")
            return False

    if target_name:
        log(f"Downloading specific attachment: {target_name!r}")
        download_one(target_name)
    else:
        # Download every attachment row. We collect visible row names first.
        log("No specific attachment set - downloading all listed attachments.")
        try:
            rows = page.get_by_role("row")
            count = rows.count()
        except Exception:
            count = 0

        names = []
        for i in range(count):
            try:
                text = rows.nth(i).inner_text(timeout=1000).strip()
                # Keep rows that look like a document name (skip headers/blank).
                first_line = text.splitlines()[0].strip() if text else ""
                if first_line and first_line not in names:
                    names.append(first_line)
            except Exception:
                continue

        if not names:
            log("No attachment rows detected. Run with HEADLESS=false to "
                "inspect, or set ATTACHMENT_NAME in .env to target one file.")
        for name in names:
            download_one(name)

    return downloaded


def navigate_and_download(page):
    """Full flow: locate client -> open policy attachments -> download."""
    locate_client(page)
    open_policy_attachments(page)
    downloaded = download_attachments(page)

    if not downloaded:
        log("No documents were downloaded.")
        log("Run with HEADLESS=false to watch the flow and adjust selectors,")
        log("or set ATTACHMENT_NAME in .env to target a specific document.")

    return downloaded


def main():
    if not config.STORAGE_STATE.exists():
        raise FileNotFoundError(
            f"No saved session found at {config.STORAGE_STATE}.\n"
            f"Run 'python epic_login.py' first to log in and save your session."
        )

    config.DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Documents will be saved to: {config.DOWNLOAD_DIR}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=config.HEADLESS)
        context = browser.new_context(
            storage_state=str(config.STORAGE_STATE),
            accept_downloads=True,
        )
        page = context.new_page()

        log("Opening Applied Epic with saved session...")
        page.goto(config.EPIC_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(4000)  # let the SPA settle

        if not session_is_valid(page):
            browser.close()
            raise RuntimeError(
                "Saved session appears to have EXPIRED (Epic is asking you to "
                "log in again).\nRun 'python epic_login.py' to refresh it, "
                "then re-run this script."
            )

        log("Session is valid. Navigating to documents...")
        downloaded = navigate_and_download(page)

        # Refresh the saved session so its validity window keeps rolling.
        try:
            context.storage_state(path=str(config.STORAGE_STATE))
        except Exception:
            pass

        log(f"Done. {len(downloaded)} document(s) downloaded.")
        browser.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
