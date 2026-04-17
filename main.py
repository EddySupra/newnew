from seleniumbase import SB
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from google.oauth2.service_account import Credentials
import pyautogui
import random
import time
import gspread


USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "1tXA9LKaQWmE97NsR29U6MyPtJA5qylZvHXMZPVa38oo"
WORKSHEET_NAME = "Sheet1"

STATUS_COL = 11   # K
NOTE_COL = 12     # L


# ----------------------------
# Timing helpers
# ----------------------------

def random_pause(a=0.08, b=0.22):
    time.sleep(random.uniform(a, b))


def micro_pause():
    time.sleep(random.uniform(0.02, 0.08))


def short_pause():
    time.sleep(random.uniform(0.10, 0.35))


def medium_pause():
    time.sleep(random.uniform(0.40, 1.10))


def long_pause():
    time.sleep(random.uniform(1.20, 2.80))


def wait_before_next_lead(min_seconds=60, max_seconds=120):
    total_wait = random.uniform(min_seconds, max_seconds)
    print(f"Waiting {total_wait:.1f} seconds before starting next lead...")

    end_time = time.time() + total_wait
    while time.time() < end_time:
        chunk = random.uniform(2.0, 6.0)
        remaining = end_time - time.time()
        time.sleep(min(chunk, remaining))


# ----------------------------
# Sheet setup
# ----------------------------

def connect_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes
    )
    client = gspread.authorize(creds)
    workbook = client.open_by_key(SPREADSHEET_ID)
    sheet = workbook.worksheet(WORKSHEET_NAME)
    return sheet


# ----------------------------
# Wait helpers
# ----------------------------

def wait_for_dom_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def wait_for_css(driver, css, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css))
    )


def wait_visible_css(driver, css, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, css))
    )


def wait_clickable_css(driver, css, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css))
    )


def wait_clickable_xpath(driver, xpath, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )


# ----------------------------
# Browser JS helpers
# Executes against live rendered DOM
# ----------------------------

def get_rendered_html(driver):
    return driver.execute_script("return document.documentElement.outerHTML;")


def get_rendered_text(driver, css):
    return driver.execute_script("""
        const el = document.querySelector(arguments[0]);
        return el ? el.innerText : null;
    """, css)


def get_rendered_value(driver, css):
    return driver.execute_script("""
        const el = document.querySelector(arguments[0]);
        return el ? el.value : null;
    """, css)


def click_in_browser_js(driver, css):
    return driver.execute_script("""
        const el = document.querySelector(arguments[0]);
        if (!el) return false;
        el.scrollIntoView({block: 'center', inline: 'center'});
        el.click();
        return true;
    """, css)


def dispatch_rich_input_events(driver, element):
    driver.execute_script("""
        const el = arguments[0];
        el.dispatchEvent(new Event('focus', { bubbles: true }));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
    """, element)


def js_set_input_value(driver, element, value):
    driver.execute_script("""
        const el = arguments[0];
        const val = arguments[1];

        el.focus();
        el.value = '';
        el.dispatchEvent(new Event('input', { bubbles: true }));

        el.value = val;
        el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'a' }));
        el.dispatchEvent(new KeyboardEvent('keypress', { bubbles: true, key: 'a' }));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'a' }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
    """, element, str(value))


def wait_for_value(driver, element, expected_value, timeout=5):
    expected_value = str(expected_value).strip()
    WebDriverWait(driver, timeout).until(
        lambda d: (element.get_attribute("value") or "").strip() == expected_value
    )


# ----------------------------
# Human-like interaction helpers
# ----------------------------

def human_hover_and_click(driver, element, pre_click_pause=(0.08, 0.25)):
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        element
    )
    random_pause(0.12, 0.34)

    actions = ActionChains(driver)
    actions.move_to_element(element)
    actions.pause(random.uniform(*pre_click_pause))
    actions.click()
    actions.perform()


def get_element_screen_center(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        element
    )
    random_pause(0.16, 0.36)

    data = driver.execute_script("""
        const el = arguments[0];
        const rect = el.getBoundingClientRect();

        return {
            x: rect.left + (rect.width / 2),
            y: rect.top + (rect.height / 2),
            outerHeight: window.outerHeight,
            innerHeight: window.innerHeight,
            screenX: window.screenX,
            screenY: window.screenY
        };
    """, element)

    browser_toolbar_height = data["outerHeight"] - data["innerHeight"]
    screen_x = data["screenX"] + data["x"]
    screen_y = data["screenY"] + browser_toolbar_height + data["y"]

    return int(screen_x), int(screen_y)


def real_mouse_click_element(driver, element, duration_range=(0.30, 0.85)):
    x, y = get_element_screen_center(driver, element)
    x += random.randint(-4, 4)
    y += random.randint(-4, 4)

    duration = random.uniform(*duration_range)
    pyautogui.moveTo(x, y, duration=duration, tween=pyautogui.easeInOutQuad)
    random_pause(0.05, 0.16)
    pyautogui.click()


def human_focus_element(driver, element, use_real_mouse=False):
    if use_real_mouse:
        real_mouse_click_element(driver, element)
    else:
        human_hover_and_click(driver, element)
    random_pause(0.06, 0.18)


def human_clear_element(element, ctrl_a=True):
    if ctrl_a:
        element.send_keys(Keys.CONTROL, "a")
        random_pause(0.04, 0.12)
        element.send_keys(Keys.BACKSPACE)
    else:
        current_value = element.get_attribute("value") or ""
        for _ in current_value:
            element.send_keys(Keys.BACKSPACE)
            random_pause(0.02, 0.06)

    random_pause(0.05, 0.16)


def human_type_element(element, text, min_delay=0.04, max_delay=0.12, typo_chance=0.0):
    letters = "abcdefghijklmnopqrstuvwxyz"

    for ch in str(text):
        if typo_chance > 0 and ch.isalpha() and random.random() < typo_chance:
            wrong = random.choice(letters.replace(ch.lower(), ""))
            if ch.isupper():
                wrong = wrong.upper()
            element.send_keys(wrong)
            time.sleep(random.uniform(min_delay, max_delay))
            element.send_keys(Keys.BACKSPACE)
            time.sleep(random.uniform(min_delay, max_delay))

        element.send_keys(ch)
        time.sleep(random.uniform(min_delay, max_delay))


def finish_field_with_tab(element):
    random_pause(0.04, 0.12)
    element.send_keys(Keys.TAB)
    random_pause(0.08, 0.22)


# ----------------------------
# Hybrid click helper
# ----------------------------

def click_element_hybrid(driver, element=None, css=None, xpath=None, timeout=20, use_real_mouse=True):
    if element is None:
        if css:
            element = wait_clickable_css(driver, css, timeout=timeout)
        elif xpath:
            element = wait_clickable_xpath(driver, xpath, timeout=timeout)
        else:
            raise ValueError("Provide element, css, or xpath")

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        element
    )
    random_pause(0.10, 0.28)

    try:
        human_focus_element(driver, element, use_real_mouse=use_real_mouse)
        return element
    except Exception:
        pass

    try:
        human_hover_and_click(driver, element)
        return element
    except Exception:
        pass

    try:
        element.click()
        random_pause(0.06, 0.18)
        return element
    except Exception:
        pass

    driver.execute_script("arguments[0].click();", element)
    random_pause(0.08, 0.22)
    return element


def safe_action_click_xpath_with_fallback(driver, xpath, timeout=20):
    elem = wait_clickable_xpath(driver, xpath, timeout=timeout)
    click_element_hybrid(driver, element=elem, use_real_mouse=True)
    return elem


# ----------------------------
# Input helper
# ----------------------------

def strong_type_css_human_first(driver, css, text, timeout=30, use_real_mouse=False):
    target_text = str(text).strip()
    elem = wait_visible_css(driver, css, timeout=timeout)

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        elem
    )
    random_pause(0.10, 0.30)

    try:
        human_focus_element(driver, elem, use_real_mouse=use_real_mouse)
    except Exception:
        try:
            human_hover_and_click(driver, elem)
        except Exception:
            driver.execute_script("arguments[0].focus();", elem)
            random_pause(0.04, 0.12)

    human_clear_element(elem, ctrl_a=True)
    random_pause(0.04, 0.12)

    human_type_element(
        elem,
        target_text,
        min_delay=random.uniform(0.035, 0.055),
        max_delay=random.uniform(0.085, 0.145),
        typo_chance=random.uniform(0.0, 0.02),
    )

    random_pause(0.04, 0.12)

    try:
        dispatch_rich_input_events(driver, elem)
    except Exception:
        pass

    actual = (elem.get_attribute("value") or "").strip()
    if actual != target_text:
        js_set_input_value(driver, elem, target_text)
        wait_for_value(driver, elem, target_text, timeout=5)

    finish_field_with_tab(elem)
    random_pause(0.08, 0.20)


def strong_select_state(driver, state_value, timeout=30, use_real_mouse=False):
    state_value = str(state_value).strip().upper()

    select_elem = wait_clickable_css(driver, "#state", timeout=timeout)
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        select_elem
    )
    random_pause(0.18, 0.42)

    try:
        human_focus_element(driver, select_elem, use_real_mouse=use_real_mouse)
    except Exception:
        try:
            human_hover_and_click(driver, select_elem)
        except Exception:
            driver.execute_script("arguments[0].click();", select_elem)

    random_pause(0.20, 0.55)

    sel = Select(select_elem)
    try:
        sel.select_by_value(state_value)
    except Exception:
        try:
            sel.select_by_visible_text(state_value)
        except Exception:
            raise Exception(f"Could not select state: {state_value}")

    random_pause(0.15, 0.35)

    driver.execute_script("""
        const el = arguments[0];
        el.dispatchEvent(new Event('focus', { bubbles: true }));
        el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowDown' }));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowDown' }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
    """, select_elem)

    random_pause(0.20, 0.45)

    selected_value = (select_elem.get_attribute("value") or "").strip().upper()
    if selected_value != state_value:
        raise Exception(f"State selection did not stick. Wanted {state_value}, got {selected_value}")


# ----------------------------
# Validation / debug helpers
# ----------------------------

def get_invalid_fields(driver):
    return driver.find_elements(
        By.CSS_SELECTOR,
        "#firstName.ng-invalid, "
        "#lastName.ng-invalid, "
        "#dobMonth.ng-invalid, "
        "#dobDay.ng-invalid, "
        "#dobYear.ng-invalid, "
        "#ssnText.ng-invalid, "
        "#streetNumberName.ng-invalid, "
        "#apt.ng-invalid, "
        "#city.ng-invalid, "
        "#state.ng-invalid, "
        "#zipcode.ng-invalid"
    )


def print_invalid_fields(driver):
    invalids = get_invalid_fields(driver)
    print(f"Invalid field count: {len(invalids)}")
    for el in invalids:
        try:
            print(
                "Invalid:",
                el.get_attribute("id"),
                "| name:", el.get_attribute("name"),
                "| class:", el.get_attribute("class"),
                "| value:", el.get_attribute("value")
            )
        except StaleElementReferenceException:
            print("Invalid field became stale before debug print")


def debug_gov_program_buttons(driver):
    print("----- GOV PROGRAM BUTTON DEBUG -----")
    buttons = driver.find_elements(By.CSS_SELECTOR, "button.nvga-next.indi-button--primary")
    for btn in buttons:
        try:
            print({
                "id": btn.get_attribute("id"),
                "text": btn.text,
                "class": btn.get_attribute("class"),
                "disabled": btn.get_attribute("disabled"),
                "aria_hidden": btn.get_attribute("aria-hidden"),
                "enabled": btn.is_enabled(),
            })
        except StaleElementReferenceException:
            print("Button became stale during debug")
    print("-----------------------------------")


def debug_post_click_state(driver, before_url):
    print("----- POST CLICK DEBUG -----")
    print("URL before click:", before_url)
    print("URL now:", driver.current_url)

    try:
        btn = driver.find_element(By.CSS_SELECTOR, "#serviceProviderNextButton")
        print("Button still present: YES")
        print("Button disabled attr:", btn.get_attribute("disabled"))
        print("Button aria-hidden:", btn.get_attribute("aria-hidden"))
        print("Button class:", btn.get_attribute("class"))
        print("Button enabled:", btn.is_enabled())
    except Exception:
        print("Button still present: NO")

    try:
        error_btn = driver.find_element(By.CSS_SELECTOR, "#infoNextErrorButton")
        print("Error button aria-hidden:", error_btn.get_attribute("aria-hidden"))
        print("Error button class:", error_btn.get_attribute("class"))
    except Exception:
        print("Error button not found")

    print_invalid_fields(driver)

    try:
        body_text = driver.execute_script("return document.body.innerText.slice(0, 3000);")
        print("Body text preview:")
        print(body_text)
    except Exception:
        print("Could not read body text")

    print("----------------------------")


# ----------------------------
# Page-specific logic
# ----------------------------

def wait_until_service_provider_next_enabled(driver, timeout=10):
    def _ready(d):
        btn = d.find_element(By.CSS_SELECTOR, "#serviceProviderNextButton")
        disabled_attr = btn.get_attribute("disabled")
        aria_hidden = (btn.get_attribute("aria-hidden") or "").strip().lower()
        classes = btn.get_attribute("class") or ""
        return (
            aria_hidden == "false"
            and disabled_attr is None
            and btn.is_enabled()
            and "ng-hide" not in classes
        )
    WebDriverWait(driver, timeout).until(_ready)


def click_checkbox_like_human(driver, css, timeout=20):
    checkbox = wait_visible_css(driver, css, timeout=timeout)

    try:
        click_element_hybrid(driver, element=checkbox, use_real_mouse=True)
    except Exception:
        try:
            clicked = click_in_browser_js(driver, css)
            if not clicked:
                raise Exception("JS click returned false")
        except Exception:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                checkbox
            )
            random_pause(0.18, 0.40)
            driver.execute_script("arguments[0].click();", checkbox)

    random_pause(0.40, 0.95)


def wait_until_gov_program_next_enabled(driver, timeout=10):
    def _ready(d):
        buttons = d.find_elements(By.CSS_SELECTOR, "button.nvga-next.indi-button--primary")
        for btn in buttons:
            btn_id = (btn.get_attribute("id") or "").strip()
            aria_hidden = (btn.get_attribute("aria-hidden") or "").strip().lower()
            classes = btn.get_attribute("class") or ""
            disabled_attr = btn.get_attribute("disabled")
            text = (btn.text or "").strip().lower()

            if btn_id == "govProgramNextErrorButton":
                continue
            if text != "next":
                continue
            if aria_hidden == "true":
                continue
            if "ng-hide" in classes:
                continue
            if disabled_attr is not None:
                continue
            if not btn.is_enabled():
                continue

            return btn

        return False

    return WebDriverWait(driver, timeout).until(_ready)


def wait_for_page_advance(driver, before_url, timeout=8):
    def _advanced(d):
        if d.current_url != before_url:
            return True

        try:
            btn = d.find_element(By.CSS_SELECTOR, "#serviceProviderNextButton")
            aria_hidden = (btn.get_attribute("aria-hidden") or "").strip().lower()
            classes = btn.get_attribute("class") or ""
            if aria_hidden == "true" or "ng-hide" in classes:
                return True
        except Exception:
            return True

        next_page_selectors = [
            "#eligSnapSpan",
            "#nextSuccessButton9",
            "#address-1",
            "#address1",
            "#applicationReview",
            "#tribalId",
        ]

        for css in next_page_selectors:
            try:
                elems = d.find_elements(By.CSS_SELECTOR, css)
                if elems:
                    return True
            except Exception:
                pass

        return False

    WebDriverWait(driver, timeout).until(_advanced)


def wait_for_page_advance_after_gov_program(driver, before_url, timeout=8):
    def _advanced(d):
        if d.current_url != before_url:
            return True

        next_page_selectors = [
            "#nextSuccessButton9",
            "#applicationReview",
            "#tribalId",
            "#address-1",
            "#address1",
        ]

        for css in next_page_selectors:
            try:
                elems = d.find_elements(By.CSS_SELECTOR, css)
                if elems:
                    return True
            except Exception:
                pass

        try:
            snap = d.find_element(By.CSS_SELECTOR, "#eligSnapSpan")
            snap_classes = snap.get_attribute("class") or ""
            if "checked" not in snap_classes.lower():
                return True
        except Exception:
            return True

        return False

    WebDriverWait(driver, timeout).until(_advanced)


def click_gov_program_next(driver, timeout=20):
    btn = wait_until_gov_program_next_enabled(driver, timeout=timeout)
    before_url = driver.current_url

    print("Clicking gov program Next...")

    try:
        click_element_hybrid(driver, element=btn, use_real_mouse=True)
    except Exception:
        btn_id = btn.get_attribute("id")
        if btn_id:
            clicked = click_in_browser_js(driver, f"#{btn_id}")
            if not clicked:
                driver.execute_script("arguments[0].click();", btn)
        else:
            driver.execute_script("arguments[0].click();", btn)

    random_pause(0.80, 1.80)

    try:
        wait_for_page_advance_after_gov_program(driver, before_url, timeout=8)
        return True
    except TimeoutException:
        print("Gov program Next was clicked but next page did not appear.")
        debug_gov_program_buttons(driver)
        return False


def click_primary_button_like_human(driver, css, timeout=20):
    btn = wait_clickable_css(driver, css, timeout=timeout)

    try:
        click_element_hybrid(driver, element=btn, use_real_mouse=True)
    except Exception:
        clicked = click_in_browser_js(driver, css)
        if not clicked:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                btn
            )
            random_pause(0.18, 0.40)
            driver.execute_script("arguments[0].click();", btn)

    random_pause(0.80, 1.80)


def open_account_homepage(driver, timeout=15):
    wait = WebDriverWait(driver, timeout)

    try:
        my_account_btn = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[contains(@class,'btn-single') and .//span[contains(normalize-space(),'My account')]]"
            ))
        )

        click_element_hybrid(driver, element=my_account_btn, use_real_mouse=True)
        random_pause(0.45, 1.10)

        wait.until(
            lambda d: (
                "in" in (d.find_element(By.ID, "SPCollapse").get_attribute("class") or "")
                or (d.find_element(By.ID, "SPCollapse").get_attribute("aria-expanded") or "").lower() == "true"
            )
        )

        account_home = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[@id='SPCollapse' and contains(@class,'collapse')]"
                "//a[@ng-click='navigatePendingApplications();' and contains(normalize-space(),'Account homepage')]"
            ))
        )

        click_element_hybrid(driver, element=account_home, use_real_mouse=False)

        print("Clicked My account -> Account homepage")
        return True

    except TimeoutException:
        print("Could not find or click My account / Account homepage")
        return False


def recover_to_homepage(driver):
    print("Recovery: refreshing page...")
    try:
        driver.refresh()
    except Exception:
        print("Refresh hit timeout, continuing...")

    wait_for_dom_ready(driver, timeout=30)

    wait_for_dom_ready(driver, timeout=30)
    random_pause(1.5, 3.0)

    print("Recovery: opening My account dropdown...")
    wait_clickable_css(driver, "button[data-target='#SPCollapse']", timeout=30).click()

    random_pause(0.8, 1.6)

    print("Recovery: clicking Account homepage...")
    wait_clickable_xpath(
        driver,
        "//a[contains(text(), 'Account homepage')]",
        timeout=30
    ).click()

    wait_for_dom_ready(driver, timeout=30)
    random_pause(1.5, 3.0)

    print("Recovery: closing My account dropdown...")
    wait_clickable_css(driver, "button[data-target='#SPCollapse']", timeout=30).click()

    print("Recovery: waiting before next lead...")
    random_pause(20, 60)

# ----------------------------
# Data helpers
# ----------------------------

def wait_for_progress_or_force_recover(driver, timeout=20):
    print("Waiting for page progress (strict timeout)...")

    start = time.time()
    initial_marker = driver.execute_script(
        "return document.body.innerText.slice(0, 500);"
    )

    while time.time() - start < timeout:
        try:
            current_marker = driver.execute_script(
                "return document.body.innerText.slice(0, 500);"
            )

            if current_marker != initial_marker:
                print("Page progressed successfully.")
                return True

        except Exception:
            pass

        time.sleep(0.5)

    # ⛔ HARD TIMEOUT → ALWAYS RECOVER
    print("Hard timeout reached. Forcing recovery...")

    refresh_attempts = random.randint(1, 3)

    for attempt in range(refresh_attempts):
        try:
            print(f"Forced refresh {attempt+1}...")
            driver.refresh()
            wait_for_dom_ready(driver, timeout=30)
            time.sleep(random.uniform(1.5, 3.0))
        except Exception as e:
            print(f"Refresh error: {e}")

    print("Navigating back to homepage after forced refresh...")
    recover_to_homepage(driver)

    return False

def update_status(sheet, row_number, status, note=""):
    try:
        sheet.update_cell(row_number, STATUS_COL, status)
        sheet.update_cell(row_number, NOTE_COL, note)
    except Exception as e:
        print(f"Could not update status for row {row_number}: {e}")


def should_skip_row(row):
    status = (row.get("Status") or "").strip().lower()
    return status in {"done", "submitted", "skip"}


def row_from_values(values):
    values = values + [""] * (12 - len(values))
    return {
        "First Name": values[0],
        "Last Name": values[1],
        "DOB": values[2],
        "SSN4": values[3],
        "address1": values[4],
        "address2": values[5],
        "City": values[6],
        "State": values[7],
        "Zip": values[8],
        "Status": values[10],
        "Note": values[11],
    }


def normalize_ssn4(ssn4_value):
    return str(ssn4_value).strip().zfill(4)


def split_dob(dob_value):
    dob_value = str(dob_value).strip()
    parts = dob_value.split("/")

    if len(parts) != 3:
        raise ValueError(f"Invalid DOB format: {dob_value}. Expected MM/DD/YYYY")

    month = parts[0].zfill(2)
    day = parts[1].zfill(2)
    year = parts[2].strip()

    if len(year) != 4:
        raise ValueError(f"Invalid DOB year: {dob_value}")

    return month, day, year

import time
import random
from selenium.common.exceptions import StaleElementReferenceException

def wait_for_loader_or_timeout(driver, timeout=20):
    print("Watching for Loading screen...")

    start = time.time()
    loading_seen = False
    last_seen_time = None

    while time.time() - start < timeout:
        try:
            loading_elements = driver.find_elements(
                By.XPATH,
                "//*[contains(text(), 'Loading')]"
            )

            visible_loading = any(el.is_displayed() for el in loading_elements)

            if visible_loading:
                if not loading_seen:
                    print("Loading screen detected...")
                    loading_seen = True

                last_seen_time = time.time()

            else:
                if loading_seen:
                    # loader appeared AND disappeared → success
                    print("Loading finished normally.")
                    return True

        except StaleElementReferenceException:
            pass
        except Exception:
            pass

        time.sleep(0.5)

    # ⛔ TIMEOUT → STUCK LOADING
    print("Loading stuck > timeout. Forcing recovery...")

    refresh_attempts = random.randint(1, 3)

    for i in range(refresh_attempts):
        try:
            print(f"Refresh attempt {i+1}")

            try:
                driver.refresh()
            except Exception:
                print("Refresh timeout, continuing...")

            wait_for_dom_ready(driver, timeout=30)
            random_pause(1.5, 3.0)

        except Exception as e:
            print(f"Refresh error: {e}")

    recover_to_homepage(driver)
    return False

# ----------------------------
# Main row processor
# ----------------------------

def fill_form_from_row(driver, row):
    wait_for_dom_ready(driver, timeout=30)

    def human_behavior():
        # small idle / scroll / thinking behavior
        if random.random() < 0.25:
            driver.execute_script(
                "window.scrollBy(0, arguments[0]);",
                random.randint(-150, 150)
            )
        if random.random() < 0.30:
            time.sleep(random.uniform(0.4, 1.6))

    def maybe_retype(css, value):
        if random.random() < 0.12:
            elem = wait_visible_css(driver, css)
            human_focus_element(driver, elem)
            human_clear_element(elem)
            human_type_element(elem, value)

    def human_type(css, value, use_mouse=False):
        strong_type_css_human_first(
            driver,
            css,
            value,
            timeout=30,
            use_real_mouse=use_mouse
        )
        if random.random() < 0.25:
            time.sleep(random.uniform(0.3, 1.2))  # thinking pause
        maybe_retype(css, value)

    def human_click(css):
        elem = wait_clickable_css(driver, css, timeout=30)

        # hesitation click
        if random.random() < 0.15:
            human_hover_and_click(driver, elem)
            time.sleep(random.uniform(0.2, 0.6))

        click_element_hybrid(
            driver,
            element=elem,
            use_real_mouse=(random.random() < 0.3)
        )

        if random.random() < 0.3:
            time.sleep(random.uniform(0.5, 1.5))

    # ----------------------------
    # Start application
    # ----------------------------
    safe_action_click_xpath_with_fallback(
        driver,
        "//a[@ng-click=\"c.startNewApplication('lifeline')\" and contains(normalize-space(), 'Start Lifeline Application')]",
        timeout=30
    )

    random_pause(1.0, 2.5)
    wait_for_dom_ready(driver, timeout=30)

    # ----------------------------
    # Prepare data
    # ----------------------------
    first_name = row.get("First Name", "").strip()
    last_name = row.get("Last Name", "").strip()
    dob_raw = row.get("DOB", "").strip()
    ssn4 = normalize_ssn4(row.get("SSN4", ""))
    address1 = row.get("address1", "").strip()
    address2 = row.get("address2", "").strip()
    city = row.get("City", "").strip()
    state = row.get("State", "").strip()
    zip_code = row.get("Zip", "").strip()

    dob_month, dob_day, dob_year = split_dob(dob_raw)

    # ----------------------------
    # Name (slightly randomized order)
    # ----------------------------
    name_fields = [
        ("#firstName", first_name),
        ("#lastName", last_name),
    ]
    random.shuffle(name_fields)

    for css, value in name_fields:
        human_type(css, value, use_mouse=(random.random() < 0.4))
        human_behavior()

    # ----------------------------
    # DOB (sometimes pause between fields)
    # ----------------------------
    human_type("#dobMonth", dob_month)
    if random.random() < 0.3:
        time.sleep(random.uniform(0.5, 1.5))

    human_type("#dobDay", dob_day)
    human_type("#dobYear", dob_year)

    human_behavior()

    # ----------------------------
    # SSN
    # ----------------------------
    human_type("#ssnText", ssn4)

    # ----------------------------
    # Address block (more "careful" behavior)
    # ----------------------------
    human_type("#streetNumberName", address1, use_mouse=True)

    if address2:
        if random.random() < 0.6:  # sometimes skip then come back
            human_type("#apt", address2)

    human_behavior()

    human_type("#city", city)

    # State selection (humans pause here often)
    if random.random() < 0.5:
        time.sleep(random.uniform(0.6, 2.0))

    strong_select_state(
        driver,
        state,
        timeout=30,
        use_real_mouse=(random.random() < 0.4)
    )

    human_type("#zipcode", zip_code)

    # occasional correction behavior
    if random.random() < 0.15:
        human_type("#zipcode", zip_code)

    human_behavior()

    # ----------------------------
    # Pre-submit hesitation
    # ----------------------------
    if random.random() < 0.4:
        time.sleep(random.uniform(1.5, 3.5))

    print("Checking for invalid fields...")
    print_invalid_fields(driver)

    wait_until_service_provider_next_enabled(driver, timeout=10)

    before_url = driver.current_url

    print("Clicking first Next...")
    human_click("#serviceProviderNextButton")

    try:
        wait_for_page_advance(driver, before_url, timeout=8)
    except TimeoutException:
        debug_post_click_state(driver, before_url)
        return False, "Next click did not advance"

    # ----------------------------
    # SNAP checkbox (with reading delay)
    # ----------------------------
    human_behavior()
    click_checkbox_like_human(driver, "#eligSnapSpan", timeout=20)

    # simulate reading eligibility text
    time.sleep(random.uniform(1.5, 3.5))

    # ----------------------------
    # Gov program Next
    # ----------------------------
    if not click_gov_program_next(driver, timeout=20):
        return False, "Gov program next failed"

    # ----------------------------
    # Agreement checkbox
    # ----------------------------
    human_behavior()
    time.sleep(random.uniform(8.4, 10.9))
    click_checkbox_like_human(driver, "span.indi-form__checkbox-icon", timeout=20)

    # simulate reading terms
    time.sleep(random.uniform(3.0, 6.5))

    # ----------------------------
    # Submit
    # ----------------------------
    human_click("#nextSuccessButton9")

    # loading watchdog
    ok = wait_for_loader_or_timeout(driver, timeout=20)
    if not ok:
        return False, "Loading stuck → recovered"

    # post-submit idle (user waiting / reading)
    time.sleep(random.uniform(4.0, 8.0))

    # ----------------------------
    # Navigate back to homepage
    # ----------------------------
    open_account_homepage(driver)

    return True, "Completed with humanized flow"


# ----------------------------
# Orchestrator
# ----------------------------

def process_rows():
    sheet = connect_sheet()
    all_rows = sheet.get_all_values()
    data_rows = all_rows[1:]

    with SB(
        uc=True,
        test=True,
        incognito=False,
        agent=USER_AGENT,
        maximize=True
    ) as sb:
        driver = sb.driver

        driver.set_page_load_timeout(25)

        try:
            print("Chrome is open in SeleniumBase UC mode.")
            print("Manually navigate to the form page in the browser.")
            input("When you are on the correct page, press Enter here to start processing rows...")

            for row_number, values in enumerate(data_rows, start=2):
                row = row_from_values(values)

                if should_skip_row(row):
                    print(f"Skipping row {row_number}")
                    continue

                print(f"Processing row {row_number}: {row['First Name']} {row['Last Name']}")

                try:
                    ok, note = fill_form_from_row(driver, row)

                    if ok:
                        update_status(sheet, row_number, "DONE", note)
                        print(f"Row {row_number} completed")
                    else:
                        update_status(sheet, row_number, "FAILED", note)
                        print(f"Row {row_number} failed: {note}")

                        try:
                            recover_to_homepage(driver)
                        except Exception as e:
                            print(f"Recovery failed: {e}")

                except Exception as e:
                    error_msg = f"{type(e).__name__}: {str(e)[:200]}"
                    update_status(sheet, row_number, "ERROR", error_msg)
                    print(f"Row {row_number} error: {error_msg}")

                    try:
                        recover_to_homepage(driver)
                    except Exception as rec_err:
                        print(f"Recovery failed after exception: {rec_err}")

                wait_before_next_lead(20, 50)

        finally:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    process_rows()

