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
SPREADSHEET_ID = "1baMJ0sLCQphUdZzETLsqcYOqlgK5P5WK4bRCiP1WLjA"
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


# ----------------------------
# Data helpers
# ----------------------------

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


# ----------------------------
# Main row processor
# ----------------------------

def fill_form_from_row(driver, row):
    wait_for_dom_ready(driver, timeout=30)

    safe_action_click_xpath_with_fallback(
        driver,
        "//a[@ng-click=\"c.startNewApplication('lifeline')\" and contains(normalize-space(), 'Start Lifeline Application')]",
        timeout=30
    )
    random_pause(0.90, 1.80)
    wait_for_dom_ready(driver, timeout=30)

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

    print("Typing first name...")
    strong_type_css_human_first(driver, "#firstName", first_name, timeout=30, use_real_mouse=True)

    print("Typing last name...")
    strong_type_css_human_first(driver, "#lastName", last_name, timeout=30, use_real_mouse=False)

    print("Typing DOB month...")
    strong_type_css_human_first(driver, "#dobMonth", dob_month, timeout=30, use_real_mouse=False)

    print("Typing DOB day...")
    strong_type_css_human_first(driver, "#dobDay", dob_day, timeout=30, use_real_mouse=False)

    print("Typing DOB year...")
    strong_type_css_human_first(driver, "#dobYear", dob_year, timeout=30, use_real_mouse=False)

    print("Typing SSN4...")
    strong_type_css_human_first(driver, "#ssnText", ssn4, timeout=30, use_real_mouse=False)

    print("Typing address1...")
    strong_type_css_human_first(driver, "#streetNumberName", address1, timeout=30, use_real_mouse=True)

    if address2:
        print("Typing address2...")
        strong_type_css_human_first(driver, "#apt", address2, timeout=30, use_real_mouse=False)

    print("Typing city...")
    strong_type_css_human_first(driver, "#city", city, timeout=30, use_real_mouse=False)

    print("Selecting state...")
    strong_select_state(driver, state, timeout=30, use_real_mouse=True)

    print("Typing zip...")
    strong_type_css_human_first(driver, "#zipcode", zip_code, timeout=30, use_real_mouse=False)

    random_pause(0.70, 1.40)

    print("Rendered ZIP value:", get_rendered_value(driver, "#zipcode"))
    print("Rendered button text:", get_rendered_text(driver, "#serviceProviderNextButton"))

    print("Checking for invalid fields before clicking Next...")
    print_invalid_fields(driver)

    print("Waiting for service provider Next to become enabled...")
    wait_until_service_provider_next_enabled(driver, timeout=10)

    next_btn = wait_clickable_css(driver, "#serviceProviderNextButton", timeout=30)
    before_url = driver.current_url

    print("Clicking service provider Next...")
    clicked = False

    try:
        click_element_hybrid(driver, element=next_btn, use_real_mouse=True)
        clicked = True
    except Exception:
        try:
            clicked = click_in_browser_js(driver, "#serviceProviderNextButton")
        except Exception:
            clicked = False

    if not clicked:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            next_btn
        )
        random_pause(0.18, 0.40)
        driver.execute_script("arguments[0].click();", next_btn)

    random_pause(0.90, 1.90)

    try:
        print("Waiting for page to actually advance...")
        wait_for_page_advance(driver, before_url, timeout=8)
    except TimeoutException:
        debug_post_click_state(driver, before_url)
        try:
            rendered_html = get_rendered_html(driver)
            print("Rendered HTML preview:")
            print(rendered_html[:3000])
        except Exception:
            print("Could not capture rendered HTML")
        return False, "Clicked first Next but page did not advance"

    print("Clicking SNAP checkbox...")
    click_checkbox_like_human(driver, "#eligSnapSpan", timeout=20)

    random_pause(0.70, 1.40)

    print("Waiting for gov program Next to become enabled...")
    gov_ok = click_gov_program_next(driver, timeout=20)
    if not gov_ok:
        return False, "Clicked SNAP but gov program Next did not advance"

    print("Clicking agreement checkbox...")
    click_checkbox_like_human(driver, "span.indi-form__checkbox-icon", timeout=20)

    random_pause(12.5, 16.8)

    print("Clicking submit...")
    click_primary_button_like_human(driver, "#nextSuccessButton9", timeout=20)

    random_pause(4.0, 9.0)

    print("Opening account homepage...")
    open_account_homepage(driver)

    return True, "Advanced through SNAP and gov program Next successfully"


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
        incognito=True,
        agent=USER_AGENT,
        maximize=True
    ) as sb:
        driver = sb.driver

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

                    if ok 
                        update_status(sheet, row_number, "DONE", note)
                        print(f"Row {row_number} completed")
                    else:
                        update_status(sheet, row_number, "FAILED", note)
                        print(f"Row {row_number} failed: {note}")

                except Exception as e:
                    error_msg = f"{type(e).__name__}: {str(e)[:200]}"
                    update_status(sheet, row_number, "ERROR", error_msg)
                    print(f"Row {row_number} error: {error_msg}")

                wait_before_next_lead(60, 120)

        finally:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    process_rows()
