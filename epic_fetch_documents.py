"""
Fetch policy documents from Applied Epic using the authenticated browser state.

Authentication behavior:
- The script first attempts to restore the saved browser state.
- Some Applied Epic/Okta tenants reject reconstructed state after Chromium
  exits. If that happens, this script opens the authorized login flow and
  waits for you to complete Okta, then continues extraction immediately in
  the SAME browser process. Do not run epic_login.py separately first.

Navigation flow for every requested client:
    Locate -> client -> Policies -> policy -> Access -> Attachments -> Download

After authentication, the script stays running and prompts for a client name
and policy number. It keeps the same live browser/session for every job and
closes only when you enter EXIT or manually close the browser.

Run:
    python epic_fetch_documents.py
"""

from datetime import datetime
import json
import re
import sys
import time

import epic_config as config
from epic_login import (
    click_post_login_confirmation_if_visible,
    complete_post_login_confirmation,
    find_logged_in_page,
    save_session_storage,
    save_storage_state,
    try_click_login_button,
    try_fill_login,
)

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


def install_saved_browser_state(context):
    """Restore cookies, localStorage, and sessionStorage before Epic loads."""
    if not config.STORAGE_STATE.exists():
        raise FileNotFoundError(
            f"Missing {config.STORAGE_STATE}. Run 'python epic_login.py' again."
        )
    if not config.SESSION_STORAGE.exists():
        raise FileNotFoundError(
            f"Missing {config.SESSION_STORAGE}. Run 'python epic_login.py' again."
        )

    storage_state = json.loads(config.STORAGE_STATE.read_text(encoding="utf-8"))
    session_state = json.loads(
        config.SESSION_STORAGE.read_text(encoding="utf-8")
    )

    # Persistent Chromium profiles do not reliably retain session-only cookies
    # after the browser process exits. Import Playwright's saved cookies into
    # the newly opened persistent context explicitly.
    cookies = storage_state.get("cookies", [])
    if cookies:
        context.add_cookies(cookies)

    local_state = {}
    for origin_data in storage_state.get("origins", []):
        origin = origin_data.get("origin")
        if not origin:
            continue
        local_state[origin] = {
            item["name"]: item["value"]
            for item in origin_data.get("localStorage", [])
        }

    encoded_local = json.dumps(local_state).replace("</", "<\\/")
    encoded_session = json.dumps(session_state).replace("</", "<\\/")
    context.add_init_script(
        script=f"""
        (() => {{
            const allLocalState = {encoded_local};
            const allSessionState = {encoded_session};
            const origin = window.location.origin;

            const localValues = allLocalState[origin] || {{}};
            for (const [key, value] of Object.entries(localValues)) {{
                window.localStorage.setItem(key, value);
            }}

            const sessionValues = allSessionState[origin] || {{}};
            for (const [key, value] of Object.entries(sessionValues)) {{
                window.sessionStorage.setItem(key, value);
            }}
        }})();
        """
    )
    log(
        "Restored saved browser state: "
        f"{len(cookies)} cookie(s), "
        f"localStorage for {len(local_state)} origin(s), "
        f"sessionStorage for {len(session_state)} origin(s)."
    )


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


def authenticate_in_current_context(context, page):
    """
    Complete Epic/Okta login and return the authenticated Epic page.

    This deliberately keeps the same Chromium process alive. Some Applied Epic
    tenants reject a reconstructed session after the browser exits, even when
    cookies and web storage are restored.
    """
    log("Saved authentication was rejected by Applied Epic.")
    log("Starting login in this same browser; extraction will continue afterward.")

    try_click_login_button(page)
    page.wait_for_timeout(4000)

    sign_in_page = context.pages[-1]
    if sign_in_page is not page:
        log("Switched to the Okta sign-in popup.")
    try_fill_login(sign_in_page)

    print("\n" + "=" * 70)
    print(" COMPLETE LOGIN IN THE OPEN BROWSER:")
    print("   1. Enter/confirm your username and password.")
    print("   2. Complete the authorized Okta push/code verification.")
    print("   3. Wait for the Epic home screen with the Locate toolbar.")
    print("\n Do not close the browser or press ENTER here.")
    print(" Extraction starts automatically after successful login.")
    print("=" * 70 + "\n")

    deadline = time.time() + config.LOGIN_TIMEOUT_SECONDS
    authenticated_page = None
    while time.time() < deadline:
        # Epic may show its Yes confirmation before Locate is detectable.
        for candidate_page in list(context.pages):
            if click_post_login_confirmation_if_visible(candidate_page):
                candidate_page.wait_for_timeout(3000)
        authenticated_page = find_logged_in_page(context)
        if authenticated_page is not None:
            break
        log("Waiting for the Epic home screen...")
        time.sleep(3)

    if authenticated_page is None:
        raise RuntimeError(
            "Timed out waiting for the authenticated Epic home screen. "
            "Re-run the script and complete Okta before the configured timeout."
        )

    authenticated_page.bring_to_front()
    complete_post_login_confirmation(authenticated_page)
    authenticated_page.wait_for_timeout(3000)
    log("Confirmed authenticated and unlocked Epic home screen.")

    # Keep backups for diagnostics/future reuse attempts. The extraction below
    # does not close or replace this authenticated browser process.
    save_session_storage(context)
    save_storage_state(context)
    return authenticated_page


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


def mouse_click_locator(page, locator, description, x_ratio=0.5, y_ratio=0.5):
    """Click a visible locator using Playwright's low-level mouse events."""
    locator.scroll_into_view_if_needed(timeout=5000)
    box = locator.bounding_box()
    if box is None or box["width"] <= 0 or box["height"] <= 0:
        raise RuntimeError(f"{description} has no clickable bounding box.")

    x = box["x"] + box["width"] * x_ratio
    y = box["y"] + box["height"] * y_ratio
    log(
        f"Mouse-clicking {description} at ({x:.0f}, {y:.0f}); "
        f"box={box}."
    )
    page.mouse.move(x, y)
    page.wait_for_timeout(250)
    page.mouse.down()
    page.wait_for_timeout(100)
    page.mouse.up()


def mouse_click_first_visible(page, locator, description):
    """Mouse-click the first visible match and return True when found."""
    count = locator.count()
    for index in range(count):
        candidate = locator.nth(index)
        try:
            if candidate.is_visible(timeout=300):
                mouse_click_locator(page, candidate, f"{description}[{index}]")
                return True
        except Exception:
            continue
    return False


def find_visible_textbox(page):
    """Return Epic Locate's visible Account/Business-name textbox."""
    exact_xpaths = [
        # Selectors captured from Applied Epic; program index varies by screen.
        "/html/body/root/app/program[2]/div/div[2]/div[2]/div/div/screen/div/div/div/proxy[1]/div/asi-frame/div/div[2]/proxy[5]/div/asi-string-edit/div/input",
        "/html/body/root/app/program[1]/div/div[2]/div[2]/div/div/screen/div/div/div/proxy[1]/div/asi-frame/div/div[2]/proxy[5]/div/asi-string-edit/div/input",
    ]
    for xpath in exact_xpaths:
        try:
            box = page.locator(f"xpath={xpath}")
            if box.is_visible(timeout=500):
                return box
        except Exception:
            continue

    # Semantic fallback if Applied changes the program/proxy indexes.
    boxes = page.get_by_role("textbox")
    try:
        count = boxes.count()
    except Exception:
        return None
    for index in range(count):
        try:
            box = boxes.nth(index)
            if box.is_visible(timeout=300):
                return box
        except Exception:
            continue
    return None


def wait_for_locate_textbox(page, timeout_seconds=7):
    """Wait for Locate's Account/Business search textbox."""
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        box = find_visible_textbox(page)
        if box is not None:
            return box
        page.wait_for_timeout(500)
    return None


def click_visible_locate_text(page):
    """Click a visible exact 'Locate' toolbar label."""
    candidates = page.get_by_text("Locate", exact=True)
    for index in range(candidates.count()):
        candidate = candidates.nth(index)
        try:
            if candidate.is_visible(timeout=300):
                candidate.click(timeout=5000)
                return
        except Exception:
            continue
    raise RuntimeError("No visible Locate text control was found.")


def click_active_program_locate(page):
    """Click Locate inside Epic's currently visible program tree."""
    active_program = page.locator("root app program:visible").last
    control = active_program.locator('[title="Locate (F2)"]').first
    icon = control.locator("i").first
    if icon.is_visible(timeout=3000):
        icon.click(timeout=10000)
    else:
        control.click(timeout=10000)


def javascript_click_active_locate(page):
    """Invoke the visible Locate control's DOM click handler directly."""
    control = (
        page.locator("root app program:visible")
        .last.locator('[title="Locate (F2)"]')
        .first
    )
    control.evaluate("element => element.click()")


def coordinate_click_active_locate(page):
    """Mouse-click the first visible title-matched Locate control."""
    controls = page.locator('[title="Locate (F2)"]')
    if not mouse_click_first_visible(page, controls, "Locate title control"):
        raise RuntimeError("No visible title-matched Locate control was found.")


def calibrated_click_locate(page):
    """Use the screenshot-calibrated toolbar coordinates as a last fallback."""
    x = config.LOCATE_CLICK_X
    y = config.LOCATE_CLICK_Y
    log(f"Mouse-clicking configured Locate coordinates ({x}, {y}).")
    page.mouse.move(x, y)
    page.wait_for_timeout(250)
    page.mouse.down()
    page.wait_for_timeout(100)
    page.mouse.up()


def dispatch_f2(page):
    """Send F2 through Playwright and dispatch the corresponding DOM events."""
    page.bring_to_front()
    page.keyboard.press("F2")
    page.evaluate(
        """() => {
            const options = {
                key: 'F2', code: 'F2', keyCode: 113, which: 113,
                bubbles: true, cancelable: true
            };
            for (const target of [document.activeElement, document.body, document, window]) {
                if (!target) continue;
                target.dispatchEvent(new KeyboardEvent('keydown', options));
                target.dispatchEvent(new KeyboardEvent('keyup', options));
            }
        }"""
    )


def open_locate(page):
    """Open Locate and verify its textbox, targeting the active program."""
    title_control = page.get_by_title("Locate (F2)")
    locate_xpath_program_2 = (
        "/html/body/root/app/program[2]/div/toolbar/div/div[2]/div/"
        "div[2]/div[1]"
    )
    locate_xpath_program_1 = (
        "/html/body/root/app/program[1]/div/toolbar/div/div[2]/div/"
        "div[2]/div[1]"
    )
    attempts = [
        ("visible Epic program Locate control", lambda: click_active_program_locate(page)),
        ("visible Locate DOM handler", lambda: javascript_click_active_locate(page)),
        ("visible Locate screen coordinates", lambda: coordinate_click_active_locate(page)),
        ("configured Locate page coordinates", lambda: calibrated_click_locate(page)),
        (
            "provided Locate XPath (program 2)",
            lambda: page.locator(f"xpath={locate_xpath_program_2}").click(
                timeout=10000, force=True
            ),
        ),
        (
            "provided Locate XPath (program 1)",
            lambda: page.locator(f"xpath={locate_xpath_program_1}").click(
                timeout=10000, force=True
            ),
        ),
        (
            "Locate toolbar icon",
            lambda: title_control.locator("i").click(timeout=10000, force=True),
        ),
        (
            "Locate toolbar container",
            lambda: title_control.click(timeout=10000, force=True),
        ),
        ("F2 keyboard and DOM events", lambda: dispatch_f2(page)),
        ("visible Locate label", lambda: click_visible_locate_text(page)),
    ]

    for description, action in attempts:
        try:
            log(f"Trying to open Locate using: {description}...")
            action()
            search_box = wait_for_locate_textbox(page)
            if search_box is not None:
                log(f"Locate opened using: {description}.")
                return search_box
        except Exception as error:
            log(f"Locate attempt failed ({description}): {error}")

    # Do not terminate the authenticated worker during selector tuning. Report
    # useful DOM counts and allow one manual Locate click while polling for the
    # same search textbox; all later steps remain automated.
    try:
        visible_programs = page.locator("root app program:visible").count()
        titled_controls = page.locator('[title="Locate (F2)"]').count()
        log(
            f"Locate diagnostics: url={page.url!r}, "
            f"visible_programs={visible_programs}, "
            f"title_controls={titled_controls}."
        )
    except Exception:
        pass

    log("Automatic Locate clicks did not open the screen.")
    log("Click Locate once in the open browser; waiting up to 60 seconds...")
    search_box = wait_for_locate_textbox(page, timeout_seconds=60)
    if search_box is not None:
        log("Detected the Locate screen after the manual click; continuing.")
        return search_box

    raise RuntimeError(
        "Locate did not open automatically or during the 60-second manual "
        "fallback. The authenticated worker will remain available for another "
        "job."
    )


def submit_locate_search(page, search_box):
    """Click the supplied blue Locate button, falling back to Enter."""
    button_xpaths = [
        "/html/body/root/app/program[2]/div/div[2]/div[2]/div[1]/div/screen/div/div/div/proxy[1]/div/asi-frame/div/div[2]/proxy[12]/div/asi-button/button",
        "/html/body/root/app/program[1]/div/div[2]/div[2]/div[1]/div/screen/div/div/div/proxy[1]/div/asi-frame/div/div[2]/proxy[12]/div/asi-button/button",
    ]
    for xpath in button_xpaths:
        try:
            button = page.locator(f"xpath={xpath}")
            if button.is_visible(timeout=1000):
                mouse_click_locator(page, button, "blue Locate search button")
                log("Clicked the blue Locate search button using its XPath.")
                return
        except Exception:
            continue

    # Semantic button fallback, then keyboard fallback.
    buttons = page.get_by_role("button", name="Locate")
    for index in range(buttons.count()):
        try:
            button = buttons.nth(index)
            if button.is_visible(timeout=300):
                mouse_click_locator(page, button, "Locate button by role")
                log("Clicked the blue Locate search button by role.")
                return
        except Exception:
            continue

    search_box.press("Enter")
    log("Locate button was not found; submitted search with Enter.")


def locate_client(page, client, lookup_code=""):
    """Locate a client and open its account result."""
    log(f"Locating client: {client!r}")

    # Do not assume a click succeeded. open_locate() verifies that the Locate
    # textbox actually appeared before this function continues.
    search_box = open_locate(page)
    search_box.click()
    search_box.fill(client)
    submit_locate_search(page, search_box)
    log("Submitted the client search.")
    page.wait_for_timeout(3000)

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
        # For clients without a configured lookup code, open the last visible
        # text match; Epic's result row follows the Locate criteria controls.
        matches = page.get_by_text(
            re.compile(re.escape(client), re.IGNORECASE), exact=False
        )
        clicked = False
        for index in range(matches.count() - 1, -1, -1):
            candidate = matches.nth(index)
            try:
                if candidate.is_visible(timeout=300):
                    candidate.click(timeout=15000)
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            raise RuntimeError(f"No visible account result matched {client!r}.")
        log(f"Opened the first matching account result for: {client}")

    page.wait_for_timeout(3000)


def click_policies(page):
    """Click Policies using the supplied sidebar XPath with fallbacks."""
    policy_xpaths = [
        "/html/body/root/app/program[2]/div/div[2]/div[1]/sidebar/div/div/sidebar-button[6]/div/a",
        "/html/body/root/app/program[1]/div/div[2]/div[1]/sidebar/div/div/sidebar-button[6]/div/a",
    ]
    for xpath in policy_xpaths:
        try:
            control = page.locator(f"xpath={xpath}")
            if control.is_visible(timeout=1000):
                mouse_click_locator(page, control, "Policies sidebar control")
                log("Clicked Policies using the supplied sidebar XPath.")
                return
        except Exception:
            continue

    fallback = page.locator("a").filter(has_text="Policies")
    if not mouse_click_first_visible(page, fallback, "Policies visible label"):
        raise RuntimeError("No visible Policies control was found.")
    log("Clicked Policies using its visible label.")


def open_policy_attachments(page, policy_code):
    """Open Policies, choose the requested policy, then open Attachments."""
    click_policies(page)
    page.wait_for_timeout(3000)

    policy_matches = page.get_by_text(policy_code, exact=False)
    if not mouse_click_first_visible(page, policy_matches, f"policy {policy_code}"):
        raise RuntimeError(f"Policy {policy_code!r} was not visible in Policies.")
    log(f"Opened policy: {policy_code}")
    page.wait_for_timeout(2000)

    access = page.get_by_title("Access (None)")
    if not mouse_click_first_visible(page, access, "Access toolbar control"):
        raise RuntimeError("The Access toolbar control was not visible.")
    page.wait_for_timeout(1000)

    attachments = page.get_by_text("Attachments", exact=True)
    if not mouse_click_first_visible(page, attachments, "Attachments menu item"):
        raise RuntimeError("The Attachments menu item was not visible.")
    log("Opened the policy's Attachments.")
    page.wait_for_timeout(3000)


def download_attachments(page, target_name=""):
    """Download one requested attachment, or attempt every listed row."""
    downloaded = []
    target_name = target_name.strip()

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


def navigate_and_download(
    page,
    client,
    policy_code,
    lookup_code="",
    attachment_name="",
):
    """Run one client/policy download job in the current live session."""
    locate_client(page, client, lookup_code)
    open_policy_attachments(page, policy_code)
    downloaded = download_attachments(page, attachment_name)
    if not downloaded:
        log("No documents were downloaded.")
    return downloaded


def run_client_worker(context, page):
    """Keep the authenticated browser alive and process jobs until EXIT."""
    total_jobs = 0
    total_downloads = 0

    print("\n" + "=" * 70)
    print(" APPLIED EPIC DOCUMENT WORKER IS ACTIVE")
    print(" The same authenticated browser will stay open for every client.")
    print(" Type EXIT as the client name when you want to close it.")
    print("=" * 70)

    while True:
        if page.is_closed():
            log("The browser was closed. Stopping the worker.")
            break

        try:
            client = input(
                f"\nClient name [{config.CLIENT_NAME}] or EXIT: "
            ).strip()
        except (EOFError, KeyboardInterrupt):
            client = "EXIT"

        if client.casefold() in {"exit", "quit", "q"}:
            break
        if not client:
            client = config.CLIENT_NAME

        default_policy = (
            config.POLICY_CODE
            if client.casefold() == config.CLIENT_NAME.casefold()
            else ""
        )
        policy_prompt = f"Policy number [{default_policy}]: " if default_policy else "Policy number: "
        try:
            policy_code = input(policy_prompt).strip() or default_policy
        except (EOFError, KeyboardInterrupt):
            break

        if not policy_code:
            log("A policy number is required. No job was started.")
            continue

        # The configured lookup/attachment values belong to the configured
        # default client. Other clients are selected by their result name and
        # all their attachment rows are attempted.
        is_default_client = client.casefold() == config.CLIENT_NAME.casefold()
        lookup_code = config.LOOKUP_CODE if is_default_client else ""
        attachment_name = config.ATTACHMENT_NAME if is_default_client else ""

        try:
            # Close any leftover viewer/menu from the previous job.
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)

            if not session_is_valid(page):
                page = authenticate_in_current_context(context, page)

            log(f"Starting job for client={client!r}, policy={policy_code!r}")
            downloaded = navigate_and_download(
                page,
                client,
                policy_code,
                lookup_code=lookup_code,
                attachment_name=attachment_name,
            )
            total_jobs += 1
            total_downloads += len(downloaded)
            log(
                f"Job finished: {len(downloaded)} document(s). "
                "The browser remains active for another client."
            )
            save_current_session_storage(context)
        except Exception as error:
            log(f"JOB ERROR for {client!r} / {policy_code!r}: {error}")
            log("The browser remains open. You may enter another client or EXIT.")

    log(
        f"Worker finished: {total_jobs} successful job(s), "
        f"{total_downloads} document(s) downloaded."
    )


def main():
    config.PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    config.DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    log(f"Documents will be saved to: {config.DOWNLOAD_DIR}")

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(config.PROFILE_DIR),
            headless=False,  # visible because interactive Okta may be required
            channel=config.BROWSER_CHANNEL,
            args=["--start-maximized"],
            no_viewport=True,
            accept_downloads=True,
        )

        if config.STORAGE_STATE.exists() and config.SESSION_STORAGE.exists():
            install_saved_browser_state(context)
        else:
            log("No complete saved state found; interactive login will be used.")

        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_navigation_timeout(120000)
        page.set_default_timeout(30000)
        log("Opening Applied Epic...")
        page.goto(config.EPIC_URL, wait_until="commit", timeout=120000)

        if session_is_valid(page):
            log("Reused authenticated home screen from saved browser state.")
            complete_post_login_confirmation(page, timeout_seconds=5)
        else:
            # Applied Epic/Okta rejected the reconstructed state. Authenticate
            # and continue extraction without ending this browser process.
            page = authenticate_in_current_context(context, page)

        log("Authentication is ready. Starting the persistent client worker...")
        run_client_worker(context, page)
        save_current_session_storage(context)

        try:
            context.storage_state(path=str(config.STORAGE_STATE), indexed_db=True)
        except TypeError:
            context.storage_state(path=str(config.STORAGE_STATE))

        log("Closing the Applied Epic browser worker.")
        context.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
