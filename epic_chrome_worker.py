"""
Applied Epic worker attached to an independently running Google Chrome.

Why this exists
---------------
On this Epic tenant, toolbar controls were visible but remained non-interactive
inside Chrome launched directly by Playwright. This script instead:
  1. Starts ordinary installed Google Chrome with a dedicated local profile and
     a localhost Chrome DevTools Protocol (CDP) port.
  2. Attaches Playwright to that live Chrome process.
  3. Completes authorized Epic/Okta login when required.
  4. Keeps the browser alive and repeatedly accepts client/policy jobs.

Entering EXIT disconnects the worker but intentionally leaves Chrome running.
Run this script again to reconnect to that same live browser without navigating
or refreshing it. Close Chrome manually when you want to end the live session.
If Chrome is closed, your organization's Okta policy may require login again.

Run:
    python epic_chrome_worker.py
"""

from datetime import datetime
import os
from pathlib import Path
import subprocess
import sys
import time
from urllib.error import URLError
from urllib.request import urlopen

import epic_config as config
from epic_fetch_documents import (
    authenticate_in_current_context,
    log,
    run_client_worker,
    session_is_valid,
)

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print(
        "ERROR: Install Playwright:\n"
        "    pip install playwright\n"
        "    python -m playwright install chrome",
        flush=True,
    )
    raise


CDP_ENDPOINT = f"http://127.0.0.1:{config.CDP_PORT}"


def chrome_candidates():
    """Return likely Google Chrome executable paths on Windows."""
    if config.CHROME_PATH:
        yield Path(config.CHROME_PATH)

    for environment_name, suffix in (
        ("PROGRAMFILES", Path("Google/Chrome/Application/chrome.exe")),
        ("PROGRAMFILES(X86)", Path("Google/Chrome/Application/chrome.exe")),
        ("LOCALAPPDATA", Path("Google/Chrome/Application/chrome.exe")),
    ):
        base = os.getenv(environment_name)
        if base:
            yield Path(base) / suffix


def find_chrome():
    """Find installed Google Chrome or raise with configuration guidance."""
    for candidate in chrome_candidates():
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(
        "Google Chrome was not found. Set EPIC_CHROME_PATH in .env, for "
        "example: C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    )


def cdp_is_ready():
    """Return whether Chrome's localhost debugging endpoint is available."""
    try:
        with urlopen(f"{CDP_ENDPOINT}/json/version", timeout=1) as response:
            return response.status == 200
    except (OSError, URLError):
        return False


def start_external_chrome():
    """Start normal Chrome independently and wait for its CDP endpoint."""
    chrome = find_chrome()
    config.PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    command = [
        str(chrome),
        f"--remote-debugging-port={config.CDP_PORT}",
        "--remote-debugging-address=127.0.0.1",
        f"--user-data-dir={config.PROFILE_DIR}",
        "--start-maximized",
        "--no-first-run",
        "--no-default-browser-check",
        config.EPIC_URL,
    ]

    log(f"Starting installed Chrome: {chrome}")
    creation_flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creation_flags,
    )

    deadline = time.time() + 30
    while time.time() < deadline:
        if cdp_is_ready():
            log(f"Chrome debugging endpoint is ready at {CDP_ENDPOINT}.")
            return
        time.sleep(0.5)

    raise RuntimeError(
        "Chrome opened but its local debugging endpoint did not become ready. "
        "Close that Chrome window and confirm the dedicated profile is not "
        "already locked by another process."
    )


def get_epic_page(context):
    """Reuse the existing Epic tab without refreshing it whenever possible."""
    pages = list(context.pages)
    for page in reversed(pages):
        if "appliedepic.com" in page.url:
            return page
    return pages[0] if pages else context.new_page()


def main():
    config.DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Documents will be saved to: {config.DOWNLOAD_DIR}")

    if cdp_is_ready():
        log("Found the existing live Chrome worker; reconnecting to it.")
    else:
        start_external_chrome()

    # Do not call browser.close() or context.close(): Chrome is independently
    # owned and must remain alive after this Python worker disconnects.
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_ENDPOINT)
        if not browser.contexts:
            raise RuntimeError("Connected to Chrome, but no browser context exists.")
        context = browser.contexts[0]
        page = get_epic_page(context)
        page.set_default_navigation_timeout(120000)
        page.set_default_timeout(30000)
        page.bring_to_front()

        if not page.url.startswith("http"):
            log("Opening Applied Epic in the attached Chrome tab...")
            page.goto(config.EPIC_URL, wait_until="commit", timeout=120000)
        else:
            log(f"Using the existing browser page without refresh: {page.url}")

        if session_is_valid(page):
            log("The attached Chrome session is already authenticated.")
        else:
            page = authenticate_in_current_context(context, page)

        log("Chrome is authenticated. Starting the persistent client worker...")
        run_client_worker(context, page)
        log(
            "Worker disconnected. Chrome was intentionally left open; rerun "
            "this script to reconnect, or close Chrome manually to end it."
        )


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
