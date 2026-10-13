"""
STEP 2: Fetch documents from Applied Epic using the saved session.

What it does
------------
1. Loads the authenticated session saved by epic_login.py (storage_state.json),
   so Epic treats you as already logged in - no push notification needed.
2. Opens Epic, verifies the session is still valid (not expired).
3. Navigates to the document location(s) you configure in navigate_and_download()
   and downloads the files into DOWNLOAD_DIR.

If the saved session has EXPIRED, the script tells you to re-run epic_login.py.

IMPORTANT - this part needs YOUR tenant's specifics
---------------------------------------------------
Every Applied Epic account organises documents differently (by client /
account / policy / attachments). The navigate_and_download() function below is
where you tell Playwright how to reach YOUR documents. It ships with a generic
"click every download link on the page" fallback plus clearly marked TODOs.
To tailor it precisely, run with HEADLESS=false, watch the browser, and adjust
the selectors. (Tip: use `python -m playwright codegen` to record clicks.)

Run:
    python epic_fetch_documents.py
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


def session_is_valid(page):
    """
    Heuristic check: are we logged in, or were we bounced to a login screen?

    We look for a visible password field or a 'Sign in' control as a sign the
    session expired. Adjust these if your tenant's login page differs.
    """
    login_markers = [
        "input[type='password']",
        "text=/sign in/i",
        "text=/log in/i",
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
    Use Epic's 'Locate' screen to search for the client and open the account.

    Mirrors the manual flow:
      1. Click 'Locate' in the top toolbar.
      2. In the Locate Criteria panel, type the client name into
         'Account/Business name contains'.
      3. Click the blue 'Locate' button.
      4. Click the matching result row to open the account.

    Applied Epic's markup varies, so each step tries several selectors and
    falls back gracefully. Run with HEADLESS=false to watch it work, and
    refine selectors using the codegen recorder if a step misses.
    """
    client = config.CLIENT_NAME
    log(f"Locating client: {client!r}")

    # --- Step 1: open the Locate screen ---
    for selector in ("text=Locate", "a:has-text('Locate')", "button:has-text('Locate')"):
        try:
            el = page.locator(selector).first
            if el.is_visible(timeout=3000):
                el.click()
                log(f"Clicked 'Locate' ({selector}).")
                break
        except Exception:
            continue
    page.wait_for_timeout(2000)

    # --- Step 2: type the client name into the 'contains' field ---
    search_selectors = [
        "input[name*='contains' i]",
        "input[id*='contains' i]",
        "input[placeholder*='name' i]",
        "input[type='text']",
    ]
    filled = False
    for selector in search_selectors:
        try:
            field = page.locator(selector).first
            if field.is_visible(timeout=2000):
                field.fill(client)
                log(f"Entered client name using selector: {selector}")
                filled = True
                break
        except Exception:
            continue
    if not filled:
        log("Could not find the search box automatically - "
            "fill it in the browser (run with HEADLESS=false) if needed.")

    # --- Step 3: click the blue 'Locate' search button ---
    for selector in (
        "button:has-text('Locate')",
        "input[value='Locate']",
        "a:has-text('Locate')",
    ):
        try:
            el = page.locator(selector).last
            if el.is_visible(timeout=2000):
                el.click()
                log("Clicked the 'Locate' search button.")
                break
        except Exception:
            continue
    page.wait_for_timeout(3000)

    # --- Step 4: open the matching result row ---
    for selector in (
        f"text={client}",
        "table tbody tr",
        "[role='row']",
    ):
        try:
            row = page.locator(selector).first
            if row.is_visible(timeout=3000):
                row.dblclick()
                log(f"Opened the account result ({selector}).")
                break
        except Exception:
            continue
    page.wait_for_timeout(3000)


def open_documents_section(page):
    """
    Open the sidebar section that holds the documents (e.g. 'Attachments').

    TODO: confirm which section you want - the sidebar in your screenshot has
    Attachments, Proofs of Insurance, Policies, etc. Set SECTION below.
    """
    SECTION = "Attachments"  # change to the section you need
    for selector in (f"text={SECTION}", f"a:has-text('{SECTION}')"):
        try:
            el = page.locator(selector).first
            if el.is_visible(timeout=3000):
                el.click()
                log(f"Opened the '{SECTION}' section.")
                break
        except Exception:
            continue
    page.wait_for_timeout(3000)


def navigate_and_download(page):
    """
    Navigate Epic and download the client's documents.

    Flow: Locate the client -> open the documents section -> download files.
    The download step uses a generic 'click anything that downloads' approach
    until we tailor it from your recorded clicks (playwright codegen).
    """
    locate_client(page)
    open_documents_section(page)

    downloaded = []

    # --- Generic fallback: click elements that look like document downloads ---
    candidate_selectors = [
        "a[href$='.pdf']",
        "a[download]",
        "a:has-text('Download')",
        "button:has-text('Download')",
        "[title*='download' i]",
    ]

    seen = set()
    for selector in candidate_selectors:
        try:
            locator = page.locator(selector)
            count = locator.count()
        except Exception:
            continue

        for i in range(count):
            element = locator.nth(i)
            try:
                label = (element.inner_text(timeout=1000) or "").strip()[:60]
            except Exception:
                label = selector

            key = f"{selector}#{i}:{label}"
            if key in seen:
                continue
            seen.add(key)

            try:
                with page.expect_download(timeout=15000) as dl_info:
                    element.click()
                download = dl_info.value
                downloaded.append(save_download(download, len(downloaded) + 1))
            except PWTimeout:
                # Click did not produce a download; skip quietly.
                continue
            except Exception as exc:
                log(f"Skipped a candidate ({label!r}): {exc}")
                continue

    if not downloaded:
        log("No downloads were captured on the current page.")
        log("Edit navigate_and_download() to reach your documents, or run")
        log("with HEADLESS=false to inspect the page and adjust selectors.")

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

        log(f"Opening Applied Epic with saved session...")
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
