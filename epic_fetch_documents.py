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
    concurrent_login_warning_visible,
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
    takeover_warning_announced = False
    while time.time() < deadline:
        # A Yes button here means this account is active on another device.
        # Bring it forward but leave the takeover decision to the user.
        for candidate_page in list(context.pages):
            if concurrent_login_warning_visible(candidate_page):
                candidate_page.bring_to_front()
                if not takeover_warning_announced:
                    log(
                        "Epic reports this account is active elsewhere. "
                        "Choose Yes or No manually in the browser."
                    )
                    takeover_warning_announced = True
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
    authenticated_page.wait_for_timeout(3000)
    log("Confirmed authenticated home screen with no takeover decision pending.")

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


def wait_and_open_policy(page, policy_code, timeout_seconds=20):
    """Wait for the Policies grid to render, then open the exact policy."""
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        # Prefer an exact policy-number cell so another page label or a policy
        # with a longer, similar number cannot be selected accidentally.
        exact_matches = page.get_by_text(policy_code, exact=True)
        if mouse_click_first_visible(
            page, exact_matches, f"policy {policy_code}"
        ):
            return True
        page.wait_for_timeout(500)
    return False


def open_policy_attachments(page, policy_code):
    """Open Policies, choose the requested policy, then open Attachments."""
    click_policies(page)

    # The first visit commonly paints the Policies shell before its rows. A
    # fixed three-second delay made the initial job fail while a repeated job
    # succeeded against the now-warmed grid.
    if not wait_and_open_policy(page, policy_code):
        raise RuntimeError(
            f"Policy {policy_code!r} did not appear within 20 seconds."
        )
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


def find_description_filter_input(page):
    """Find the first filter input in the active Attachments program."""
    # Epic keeps older programs in the DOM. Prefer the last visible program,
    # which is the active screen, so a still-visible input from an older
    # Attachments instance is not returned first.
    programs = page.locator("root app program:visible")
    try:
        for index in range(programs.count() - 1, -1, -1):
            controls = programs.nth(index).locator(
                "filter-fields-form asi-combo-box input"
            )
            if controls.count() < 1:
                continue
            control = controls.first
            if control.is_visible(timeout=500) and control.bounding_box():
                return control
    except Exception:
        pass

    # Absolute-path fallbacks for Epic versions where the active program's
    # custom elements are not discoverable through the structural selector.
    xpaths = [
        "/html/body/root/app/program[3]/div/div[3]/div[2]/div/div/screen/div/div/div/proxy/div/asi-frame/div/div[2]/proxy[6]/div/asi-panel/div/proxy[3]/div/asi-virtual-list-view/div/filter-fields/div/filter-fields-form/div/asi-combo-box[1]/div/input",
        "/html/body/root/app/program[2]/div/div[3]/div[2]/div/div/screen/div/div/div/proxy/div/asi-frame/div/div[2]/proxy[6]/div/asi-panel/div/proxy[3]/div/asi-virtual-list-view/div/filter-fields/div/filter-fields-form/div/asi-combo-box[1]/div/input",
        "/html/body/root/app/program[1]/div/div[3]/div[2]/div/div/screen/div/div/div/proxy/div/asi-frame/div/div[2]/proxy[6]/div/asi-panel/div/proxy[3]/div/asi-virtual-list-view/div/filter-fields/div/filter-fields-form/div/asi-combo-box[1]/div/input",
    ]
    for xpath in xpaths:
        try:
            control = page.locator(f"xpath={xpath}")
            if control.is_visible(timeout=500) and control.bounding_box():
                return control
        except Exception:
            continue
    return None


def visible_search_inputs(page):
    """Find left-to-right inputs on the active Attachments search row."""
    scope = page
    label_box = None

    # Search newest visible programs first; Epic retains older screens in the
    # DOM and page-wide selectors can otherwise mix controls from two screens.
    programs = page.locator("root app program:visible")
    try:
        program_indexes = range(programs.count() - 1, -1, -1)
    except Exception:
        program_indexes = ()

    for program_index in program_indexes:
        program = programs.nth(program_index)
        labels = program.get_by_text("Search where", exact=True)
        for label_index in range(labels.count()):
            try:
                label = labels.nth(label_index)
                if label.is_visible(timeout=300):
                    candidate_box = label.bounding_box()
                    if candidate_box:
                        scope = program
                        label_box = candidate_box
                        break
            except Exception:
                continue
        if label_box is not None:
            break

    # Compatibility fallback when Epic does not expose a program wrapper.
    if label_box is None:
        labels = page.get_by_text("Search where", exact=True)
        for index in range(labels.count()):
            try:
                label = labels.nth(index)
                if label.is_visible(timeout=300):
                    label_box = label.bounding_box()
                    if label_box:
                        break
            except Exception:
                continue

    if label_box is None:
        raise RuntimeError("The Attachments 'Search where' toolbar was not visible.")

    label_center_y = label_box["y"] + label_box["height"] / 2
    controls = []
    inputs = scope.locator("input")
    for index in range(inputs.count()):
        candidate = inputs.nth(index)
        try:
            if not candidate.is_visible(timeout=200):
                continue
            box = candidate.bounding_box()
            if not box:
                continue
            center_y = box["y"] + box["height"] / 2
            if (
                abs(center_y - label_center_y) <= 25
                and box["x"] > label_box["x"]
            ):
                controls.append((box["x"], candidate))
        except Exception:
            continue

    controls.sort(key=lambda item: item[0])
    return [candidate for _, candidate in controls]


def wait_for_attachment_toolbar(page, timeout_seconds=20):
    """Wait until Epic exposes three stable attachment-search inputs."""
    deadline = time.time() + timeout_seconds
    previous_boxes = None
    stable_observations = 0

    while time.time() < deadline:
        try:
            controls = visible_search_inputs(page)
            description = find_description_filter_input(page)
            if len(controls) >= 3 and description is not None:
                boxes = [description.bounding_box()]
                boxes.extend(control.bounding_box() for control in controls[:3])
                if all(box is not None for box in boxes):
                    signature = tuple(
                        (
                            round(box["x"]),
                            round(box["y"]),
                            round(box["width"]),
                            round(box["height"]),
                        )
                        for box in boxes
                    )
                    if signature == previous_boxes:
                        stable_observations += 1
                    else:
                        previous_boxes = signature
                        stable_observations = 1
                    if stable_observations >= 2:
                        return
        except Exception:
            previous_boxes = None
            stable_observations = 0
        page.wait_for_timeout(500)

    raise RuntimeError(
        "The Attachments search toolbar did not stabilize within 20 seconds."
    )


def choose_visible_dropdown_option(page, control_box, option, toolbar_bottom):
    """Click an option inside the dropdown aligned with its original input."""

    matches = page.get_by_text(option, exact=True)
    candidates = []
    minimum_x = control_box["x"] - 20
    maximum_x = control_box["x"] + control_box["width"] + 100

    for index in range(matches.count()):
        candidate = matches.nth(index)
        try:
            if not candidate.is_visible(timeout=200):
                continue
            box = candidate.bounding_box()
            if not box:
                continue

            # Reject the grid's Description column header. The actual option is
            # rendered below the toolbar and horizontally aligned with the
            # first combo input/dropdown popup.
            if (
                box["y"] > toolbar_bottom + 35
                and minimum_x <= box["x"] <= maximum_x
            ):
                candidates.append((box["y"], candidate))
        except Exception:
            continue

    if not candidates:
        return False

    _, selected = min(candidates, key=lambda item: item[0])
    mouse_click_locator(page, selected, f"actual {option} dropdown option")
    return True


def click_adjacent_dropdown_arrow(page, control, description):
    """Click the separate arrow immediately right of an Epic combo input."""
    box = control.bounding_box()
    if box is None:
        raise RuntimeError(f"{description} input has no bounding box.")

    # The screenshot shows a ~20px arrow button directly beside the input.
    x = box["x"] + box["width"] + 10
    y = box["y"] + box["height"] / 2
    log(f"Mouse-clicking {description} arrow at ({x:.0f}, {y:.0f}).")
    page.mouse.move(x, y)
    page.wait_for_timeout(200)
    page.mouse.down()
    page.wait_for_timeout(100)
    page.mouse.up()
    # Opening an Epic combo can replace its input node immediately. Return the
    # pre-click geometry so option selection never asks the stale node for a
    # new bounding box.
    return box


def dropdown_has_value(control, option):
    """Return whether a combo input visibly contains the requested value."""
    try:
        current = control.input_value(timeout=1000).strip()
        return current.casefold() == option.casefold()
    except Exception:
        return False


def type_dropdown_value(page, control, option):
    """Type an option directly and verify that the combo displays it."""
    try:
        control.click(force=True)
        control.fill(option)
        page.keyboard.press("Enter")
        page.wait_for_timeout(700)
        if dropdown_has_value(control, option):
            return True
    except Exception:
        pass

    # Fallback for custom Epic inputs that intercept Playwright.fill(). Use the
    # native value setter while still dispatching input/change events so Epic's
    # component model receives the update.
    try:
        control.evaluate(
            """(element, value) => {
                const setter = Object.getOwnPropertyDescriptor(
                    HTMLInputElement.prototype, 'value'
                ).set;
                setter.call(element, value);
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
                element.dispatchEvent(new KeyboardEvent('keydown', {
                    key: 'Enter', code: 'Enter', bubbles: true
                }));
                element.dispatchEvent(new KeyboardEvent('keyup', {
                    key: 'Enter', code: 'Enter', bubbles: true
                }));
            }""",
            option,
        )
        page.wait_for_timeout(700)
        return dropdown_has_value(control, option)
    except Exception:
        return False


def select_attachment_dropdown(page, option):
    """Select an attachment filter, retrying after Epic replaces its input."""

    def fresh_control():
        if option.casefold() == "description":
            return find_description_filter_input(page)
        controls = visible_search_inputs(page)
        return controls[1] if len(controls) > 1 else None

    # Epic commonly replaces the combo input as its menu opens. Retry with a
    # newly located input, and use the geometry captured before the click when
    # finding the option instead of querying the detached input afterward.
    for attempt in range(1, 4):
        current_control = fresh_control()
        if current_control is None:
            page.wait_for_timeout(700)
            continue
        if dropdown_has_value(current_control, option):
            log(f"Attachment filter dropdown already selected: {option}.")
            return

        try:
            control_box = click_adjacent_dropdown_arrow(
                page, current_control, option
            )
            page.wait_for_timeout(700)
            current_toolbar_bottom = control_box["y"] + control_box["height"]
            if choose_visible_dropdown_option(
                page, control_box, option, current_toolbar_bottom
            ):
                page.wait_for_timeout(700)
                current_control = fresh_control()
                if current_control is not None and dropdown_has_value(
                    current_control, option
                ):
                    log(f"Selected and verified attachment filter: {option}.")
                    return
                log(
                    f"Epic did not retain {option} after attempt {attempt}; "
                    "reacquiring the filter control."
                )
            else:
                log(
                    f"The {option} option was not visible after attempt "
                    f"{attempt}; reacquiring the filter control."
                )
        except Exception as error:
            log(
                f"Could not use the {option} dropdown on attempt "
                f"{attempt}: {error}"
            )

        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        page.wait_for_timeout(700)

    if option.casefold() == "description":
        try:
            current_control = fresh_control()
            if current_control is None:
                raise RuntimeError("Description input disappeared after re-render.")
            click_adjacent_dropdown_arrow(page, current_control, option)
            page.wait_for_timeout(500)
            page.keyboard.press("Home")
            for _ in range(4):
                page.keyboard.press("ArrowDown")
            page.keyboard.press("Enter")
            page.wait_for_timeout(1000)

            current_control = fresh_control()
            if current_control is not None and dropdown_has_value(
                current_control, option
            ):
                log(
                    "Selected and verified Description by dropdown "
                    "keyboard navigation."
                )
                return
        except Exception as error:
            log(f"Description keyboard dropdown selection failed: {error}")

    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    current_control = fresh_control()
    if (
        option.casefold() != "description"
        and current_control is not None
        and type_dropdown_value(page, current_control, option)
    ):
        log(f"Typed and verified attachment filter: {option}.")
        return

    raise RuntimeError(
        f"Could not commit {option!r} as an actual attachment dropdown option."
    )


def apply_policy_attachment_filter(page):
    """Set Description / Contains / policy and click Find."""
    # The Attachments page can display its grid before the custom search
    # controls finish rendering. Require stable control geometry first.
    wait_for_attachment_toolbar(page)
    controls = visible_search_inputs(page)
    if len(controls) < 3:
        raise RuntimeError(
            "Expected three inputs on the Attachments search row "
            f"but found {len(controls)}."
        )

    description_control = find_description_filter_input(page)
    if description_control is None:
        raise RuntimeError("The exact first Description filter input was not visible.")

    select_attachment_dropdown(page, "Description")

    # Re-query because selecting the first dropdown can re-render the toolbar.
    controls = visible_search_inputs(page)
    select_attachment_dropdown(page, "Contains")

    # Final verification: never run Find while Description is blank. Retry each
    # field once after all combo re-renders, then require both visible values.
    controls = visible_search_inputs(page)
    description_control = find_description_filter_input(page)
    if description_control is None or not dropdown_has_value(
        description_control, "Description"
    ):
        log("First filter became blank; forcing Description again.")
        description_control = find_description_filter_input(page)
        if description_control is None:
            raise RuntimeError("Description filter disappeared after re-render.")
        select_attachment_dropdown(page, "Description")

    controls = visible_search_inputs(page)
    if not dropdown_has_value(controls[1], "Contains"):
        log("Second filter became blank; forcing Contains again.")
        select_attachment_dropdown(page, "Contains")

    controls = visible_search_inputs(page)
    description_control = find_description_filter_input(page)
    if description_control is None or not dropdown_has_value(
        description_control, "Description"
    ):
        raise RuntimeError("First attachment filter is still blank, not Description.")
    if not dropdown_has_value(controls[1], "Contains"):
        raise RuntimeError("Second attachment filter is still blank, not Contains.")

    search_value = controls[2]
    search_value.click(force=True)
    search_value.fill("policy")
    log("Entered 'policy' in the attachment filter.")

    find_controls = page.get_by_text("Find", exact=True)
    if not mouse_click_first_visible(page, find_controls, "Attachments Find"):
        # Accessible-role fallback for implementations where the label is nested.
        find_buttons = page.get_by_role("button", name="Find", exact=True)
        if not mouse_click_first_visible(page, find_buttons, "Attachments Find button"):
            search_value.press("Enter")
            log("Find control was not detected; submitted with Enter.")
        else:
            log("Clicked Find in the attachment filter.")
    else:
        log("Clicked Find in the attachment filter.")

    page.wait_for_timeout(3000)


def find_visible_exact_text(page, text):
    """Return the smallest visible exact-text match for a document name."""
    matches = page.get_by_text(text, exact=True)
    visible = []
    for index in range(matches.count()):
        candidate = matches.nth(index)
        try:
            if candidate.is_visible(timeout=300):
                box = candidate.bounding_box()
                if box:
                    visible.append((box["width"] * box["height"], candidate))
        except Exception:
            continue
    return min(visible, key=lambda item: item[0])[1] if visible else None


def double_click_document(page, document):
    """Open an attachment using a trusted low-level mouse double-click."""
    box = document.bounding_box()
    if box is None:
        raise RuntimeError("Document row has no clickable bounding box.")
    x = box["x"] + box["width"] / 2
    y = box["y"] + box["height"] / 2
    log(f"Double-clicking document at ({x:.0f}, {y:.0f}).")
    page.mouse.move(x, y)
    page.mouse.click(x, y, click_count=2, delay=120)


def wait_for_download_button(context, original_page, timeout_seconds=20):
    """Find Download in either an Epic modal or a newly opened popup page."""
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        pages = list(reversed(context.pages))
        if original_page not in pages:
            pages.append(original_page)
        for candidate_page in pages:
            if candidate_page.is_closed():
                continue
            try:
                buttons = candidate_page.get_by_role(
                    "button", name="Download", exact=True
                )
                for index in range(buttons.count()):
                    button = buttons.nth(index)
                    if button.is_visible(timeout=300):
                        return candidate_page, button
            except Exception:
                continue
        original_page.wait_for_timeout(500)
    return None, None


def policy_document_names(page):
    """Return unique visible attachment names whose text ends with Policy."""
    pattern = re.compile(r".+\bPolicy\s*$", re.IGNORECASE)
    matches = page.get_by_text(pattern)
    names = []
    for index in range(matches.count()):
        candidate = matches.nth(index)
        try:
            if not candidate.is_visible(timeout=300):
                continue
            text = " ".join(candidate.inner_text(timeout=1000).split())
            if pattern.fullmatch(text) and text not in names:
                names.append(text)
        except Exception:
            continue
    return names


def download_attachments(page, target_name=""):
    """Filter attachments, then download a named or Policy-ending document."""
    downloaded = []
    target_name = target_name.strip()

    apply_policy_attachment_filter(page)
    names = [target_name] if target_name else policy_document_names(page)

    if target_name:
        log(f"Searching for the configured attachment: {target_name!r}")
    else:
        log("Searching for attachment names ending in 'Policy'.")

    if not names:
        log("No visible attachment name ending in 'Policy' was found.")
        return downloaded

    log(f"Matched {len(names)} policy document(s): {names}")

    for name in names:
        try:
            document = find_visible_exact_text(page, name)
            if document is None:
                raise RuntimeError("The document name is no longer visible.")

            pages_before = set(page.context.pages)
            double_click_document(page, document)

            dialog_page, download_button = wait_for_download_button(
                page.context, page
            )
            if download_button is None:
                raise RuntimeError(
                    "Double-click succeeded, but no visible Download button "
                    "appeared in the modal or popup."
                )

            with dialog_page.expect_download(timeout=60000) as download_info:
                mouse_click_locator(
                    dialog_page,
                    download_button,
                    "Download dialog button",
                )

            saved = save_download(download_info.value, len(downloaded) + 1)
            downloaded.append(saved)

            # Close only a popup created for this document. An in-page Epic
            # modal normally closes itself after Download; Escape is a fallback.
            new_pages = [
                candidate
                for candidate in page.context.pages
                if candidate not in pages_before and candidate is not page
            ]
            for popup in new_pages:
                try:
                    popup.close()
                except Exception:
                    pass
            if dialog_page is page:
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
            page.wait_for_timeout(1500)
        except Exception as error:
            log(f"Could not download {name!r}: {error}")

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
