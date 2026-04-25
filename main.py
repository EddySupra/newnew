"""
solver.py — Automated Lifeline application filler using SeleniumBase.

Flow per lead:
  1. Navigate to getinternet.gov application
  2. Fill personal info form (name, DOB, SSN4, address)
  3. Solve reCAPTCHA (audio fallback via speech recognition)
  4. Create an account (username derived from email)
  5. Check eligibility (SNAP checkbox)
  6. Agree to terms and submit
  7. Sign out and wait before the next lead
"""

# Standard library
import asyncio
import os
import random
import time

# Third-party
import aiohttp
import gspread
import speech_recognition as sr
from google.oauth2.service_account import Credentials
from pydub import AudioSegment
from seleniumbase import SB
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)
from webdriver_manager.chrome import ChromeDriverManager


USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "1r1zErE49Fz3G_w2eej5__NPePnbeW384JGPxPAo6iDA"
WORKSHEET_NAME = "Sheet1"

STATUS_COL = 11   # K
NOTE_COL = 12     # L




class RecaptchaSolver:
    def __init__(self, driver):
        self.driver = driver

    async def download_audio(self, url, path):
        headers = {
            "User-Agent": USER_AGENT,
            "Referer": "https://www.google.com/",
            "Origin": "https://www.google.com",
        }
        async with aiohttp.ClientSession(headers=headers) as session:
            async with session.get(url) as response:
                if response.status != 200:
                    raise Exception(f"Audio download failed: HTTP {response.status}")
                with open(path, 'wb') as f:
                    f.write(await response.read())
        print("Downloaded audio asynchronously.")

    def solveCaptcha(self):
        try:
            # Switch to the CAPTCHA iframe
            iframe_inner = WebDriverWait(self.driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@title, 'reCAPTCHA')]"))
            )
            # Click on the CAPTCHA box
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'recaptcha-anchor'))
            ).click()
            random_pause(1, 3)
            # Check if the CAPTCHA is solved
            time.sleep(1)  # Allow some time for the state to update
            if self.isSolved():
                print("CAPTCHA solved by clicking.")
                self.driver.switch_to.default_content()  # Switch back to main content
                return
            # If not solved, attempt audio CAPTCHA solving
            self.solveAudioCaptcha()

        except Exception as e:
            print(f"An error occurred while solving CAPTCHA: {e}")
            self.driver.switch_to.default_content()  # Ensure we switch back in case of error
            raise

    def solveAudioCaptcha(self):
        try:
            self.driver.switch_to.default_content()
            
            # Switch to the audio CAPTCHA iframe
            iframe_audio = WebDriverWait(self.driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//iframe[@title="recaptcha challenge expires in two minutes"]'))
            )
            random_pause(.5,.7)
            # Click on the audio button
            audio_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'recaptcha-audio-button'))
            )
            audio_button.click()
            random_pause(.6,1)
            # Get the audio source URL
            audio_source = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'audio-source'))
            ).get_attribute('src')
            print(f"Audio source URL: {audio_source}")

            # Download the audio to the temp folder asynchronously
            temp_dir = os.getenv("TEMP") if os.name == "nt" else "/tmp/"
            path_to_mp3 = os.path.normpath(os.path.join(temp_dir, f"{random.randrange(1, 1000)}.mp3"))
            path_to_wav = os.path.normpath(os.path.join(temp_dir, f"{random.randrange(1, 1000)}.wav"))
            random_pause(.2,.7)
            asyncio.run(self.download_audio(audio_source, path_to_mp3))
            random_pause(.2,.7)
            # Convert mp3 to wav
            sound = AudioSegment.from_mp3(path_to_mp3)
            sound.export(path_to_wav, format="wav")
            print("Converted MP3 to WAV.")
            
            # Recognize the audio
            recognizer = sr.Recognizer()
            with sr.AudioFile(path_to_wav) as source:
                audio = recognizer.record(source)
            captcha_text = recognizer.recognize_google(audio).lower()
            print(f"Recognized CAPTCHA text: {captcha_text}")
            
            # Enter the CAPTCHA text
            audio_response = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, 'audio-response'))
            )
            audio_response.send_keys(captcha_text)
            audio_response.send_keys(Keys.ENTER)
            print("Entered and submitted CAPTCHA text.")
           
            # Wait for CAPTCHA to be processed
            time.sleep(0.8)  # Increase this if necessary

            # Verify CAPTCHA is solved
            if self.isSolved():
                print("Audio CAPTCHA solved.")
            else:
                print("Failed to solve audio CAPTCHA.")
                raise Exception("Failed to solve CAPTCHA")

        except Exception as e: 
            print(f"An error occurred while solving audio CAPTCHA: {e}")
            self.driver.switch_to.default_content()  # Ensure we switch back in case of error
            raise

        finally:
            # Always switch back to the main content
            self.driver.switch_to.default_content()

    def isSolved(self):
        try:
            # Switch back to the default content
            self.driver.switch_to.default_content()

            # Switch to the reCAPTCHA iframe
            iframe_check = self.driver.find_element(By.XPATH, "//iframe[contains(@title, 'reCAPTCHA')]")
            self.driver.switch_to.frame(iframe_check)

            # Find the checkbox element and check its aria-checked attribute
            checkbox = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'recaptcha-anchor'))
            )
            aria_checked = checkbox.get_attribute("aria-checked")

            # Return True if the aria-checked attribute is "true" or the checkbox has the 'recaptcha-checkbox-checked' class
            return aria_checked == "true" or 'recaptcha-checkbox-checked' in checkbox.get_attribute("class")

        except Exception as e:
            print(f"An error occurred while checking if CAPTCHA is solved: {e}")
            return False

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

def angular_human_type(driver, element, text,
                       min_delay=0.05, max_delay=0.14,
                       pause_chance=0.05):

    text = str(text)

    driver.execute_script("""
        arguments[0].scrollIntoView({block:'center'});
    """, element)

    # stable focus (no re-render flicker)
    driver.execute_script("""
        const el = arguments[0];
        el.focus();
        el.click();
    """, element)

    time.sleep(random.uniform(0.10, 0.20))

    # clear in Angular-safe way
    driver.execute_script("""
        const el = arguments[0];
        el.value = '';
        el.dispatchEvent(new Event('input', {bubbles:true}));
    """, element)

    for ch in text:
        element.send_keys(ch)

        # Angular reactivity trigger
        driver.execute_script("""
            arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
        """, element)

        time.sleep(random.uniform(min_delay, max_delay))

        # human hesitation (rare)
        if random.random() < pause_chance:
            time.sleep(random.uniform(0.25, 0.6))

    # finalize Angular model sync
    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
    """, element)

    element.send_keys(Keys.TAB)
    time.sleep(random.uniform(0.06, 0.16))

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


def dispatch_rich_input_events(driver, element, input_text=None):
    actions = ActionChains(driver)

    # Simulate focus on the element (click to focus)
    actions.move_to_element(element).click().perform()
    time.sleep(random.uniform(0.2, 0.5))  # Mimic human pause

    if input_text:
        # Simulate typing the text character by character
        for char in input_text:
            actions.send_keys(char)
            time.sleep(random.uniform(0.04, 0.12))  # Mimic typing speed

    # Simulate 'change' event (e.g., pressing Enter or tabbing away)
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(random.uniform(0.2, 0.5))  # Pause before blurring

    # Simulate blur by clicking somewhere else or moving focus
    actions.move_by_offset(0, 50).click().perform()
    time.sleep(random.uniform(0.3, 0.6))  # Pause to allow blur to be triggered


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


def simulate_human_mouse_path(driver, element, steps=7):
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
        element
    )
    random_pause(0.03, 0.07)

    start_x = random.randint(-280, 280)
    start_y = random.randint(-160, 160)
    if abs(start_x) < 30:
        start_x = 30 * (1 if start_x >= 0 else -1)

    cp_x = start_x * random.uniform(0.2, 0.8) + random.uniform(-25, 25)
    cp_y = start_y * random.uniform(0.2, 0.8) + random.uniform(-25, 25)

    actions = ActionChains(driver)
    actions.move_to_element_with_offset(element, start_x, start_y)

    prev_x, prev_y = float(start_x), float(start_y)

    for i in range(1, steps + 1):
        t = i / steps
        cur_x = (1 - t) ** 2 * start_x + 2 * (1 - t) * t * cp_x + t ** 2 * 0
        cur_y = (1 - t) ** 2 * start_y + 2 * (1 - t) * t * cp_y + t ** 2 * 0
        cur_x += random.uniform(-1.5, 1.5)
        cur_y += random.uniform(-1.5, 1.5)

        actions.move_by_offset(int(cur_x - prev_x), int(cur_y - prev_y))
        actions.pause(random.uniform(0.003, 0.008))

        prev_x, prev_y = cur_x, cur_y

    actions.perform()
    random_pause(0.02, 0.06)


def human_hover_and_click(driver, element, pre_click_pause=(0.08, 0.25)):
    simulate_human_mouse_path(driver, element)

    actions = ActionChains(driver)
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
    human_hover_and_click(driver, element)

def human_focus_element(driver, element, use_real_mouse=False):
    if use_real_mouse:
        real_mouse_click_element(driver, element)
    else:
        human_hover_and_click(driver, element)
    time.sleep(random.uniform(0.06, 0.18))


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
    time.sleep(random.uniform(0.10, 0.30))

    try:
        human_focus_element(driver, elem, use_real_mouse=use_real_mouse)
    except Exception:
        try:
            human_hover_and_click(driver, elem)
        except Exception:
            driver.execute_script("arguments[0].focus();", elem)
            time.sleep(random.uniform(0.04, 0.12))

    human_clear_element(elem, ctrl_a=True)
    time.sleep(random.uniform(0.04, 0.12))

    human_type_element(
        elem,
        target_text,
        min_delay=random.uniform(0.035, 0.055),
        max_delay=random.uniform(0.085, 0.145),
        typo_chance=random.uniform(0.0, 0.02),
    )

    time.sleep(random.uniform(0.04, 0.12))

    # Dispatch rich input events after typing
    try:
        dispatch_rich_input_events(driver, elem, input_text=target_text)
    except Exception:
        pass

    actual = (elem.get_attribute("value") or "").strip()
    if actual != target_text:
        js_set_input_value(driver, elem, target_text)
        wait_for_value(driver, elem, target_text, timeout=5)

    finish_field_with_tab(elem)
    time.sleep(random.uniform(0.08, 0.20))


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
        try:
            btn = d.find_element(By.ID, "consumerNextSuccessButton")

            # Must be displayed in real DOM
            if not btn.is_displayed():
                return False

            # Must NOT be disabled attribute
            if btn.get_attribute("disabled"):
                return False

            # Angular class-based hiding
            classes = btn.get_attribute("class") or ""
            if "ng-hide" in classes:
                return False

            # Check actual clickability
            return EC.element_to_be_clickable((By.ID, "consumerNextSuccessButton"))(d)

        except:
            return False

    WebDriverWait(driver, timeout).until(_ready)


def click_next(driver, timeout=10):

    wait = WebDriverWait(driver, timeout)

    # wait for ANY visible button
    btn = wait.until(lambda d: next(
        (b for b in d.find_elements(By.TAG_NAME, "button")
         if b.is_displayed() and b.text.strip() == "Next"),
        None
    ))

    driver.execute_script("arguments[0].click();", btn)


def click_checkbox_like_human(driver, css, timeout=20):
    checkbox = wait_visible_css(driver, css, timeout=timeout)

    try:
        # Hover and click to simulate real user behavior
        human_hover_and_click(driver, checkbox, pre_click_pause=(0.12, 0.22))
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
            random_pause(0.18, 2.2)
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

def wait_for_captcha_solved(driver, timeout=120):
    print("Waiting for captcha...")

    # capture existing token (important)
    initial_token = driver.execute_script("""
        const el = document.getElementById('g-recaptcha-response');
        return el ? el.value : '';
    """)

    def new_token(d):
        try:
            return d.execute_script("""
                const el = document.getElementById('g-recaptcha-response');
                if (!el) return false;

                const val = el.value.trim();
                return val.length > 0 && val !== arguments[0];
            """, initial_token)
        except:
            return False

    WebDriverWait(driver, timeout).until(new_token)

    time.sleep(random.uniform(2, 4))  # let site process captcha

    print("Captcha completed and accepted.")

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

    # Pre-click random pause to simulate thinking time
    random_pause(0.18, 0.40)

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
def check_duplicate_account(driver):
    """
    Returns:
        (skip_lead: bool, reason: str)
    """

    try:
        el = driver.find_element(
            By.CSS_SELECTOR,
            "div[ng-if='pageForm.$error.duplicatePii']"
        )

        if el.is_displayed():
            text = el.text.lower()

            if "already exist in our system" in text:
                return True, "duplicate_account"

    except:
        pass

    # fallback text check (Angular sometimes renders differently)
    if "already exist in our system" in driver.page_source.lower():
        return True, "duplicate_account"

    return False, None  

def click_terms_checkbox(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)

    checkbox = driver.find_element(By.ID, "applicantTermsConditionsCheckbox")

    if checkbox.is_selected():
        return True

    label = driver.find_element(
        By.CSS_SELECTOR,
        'label[for="applicantTermsConditionsCheckbox"]'
    )

    driver.execute_script("arguments[0].click();", label)

    # verify
    wait.until(lambda d: d.find_element(By.ID, "applicantTermsConditionsCheckbox").is_selected())

    return True


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


def click_consumer_submit(driver, timeout=20):

    wait = WebDriverWait(driver, timeout)

    def get_button(d):
        try:
            btn = d.find_element(By.ID, "consumerNextButton")

            if not btn.is_displayed():
                return False
            if btn.get_attribute("disabled"):
                return False
            if "ng-hide" in (btn.get_attribute("class") or ""):
                return False

            # extra Angular safety check: must be clickable size-wise
            if btn.size["width"] < 1 or btn.size["height"] < 1:
                return False

            return btn

        except:
            return False

    # wait until Angular exposes real button
    btn = wait.until(get_button)

    # IMPORTANT: allow Angular digest to settle
    time.sleep(0.5)

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        driver.execute_script("arguments[0].click();", btn)

    except (StaleElementReferenceException, ElementClickInterceptedException):
        # re-fetch and retry once (VERY important in Angular apps)
        btn = driver.find_element(By.ID, "consumerNextButton")
        driver.execute_script("arguments[0].click();", btn)

    return True


def detect_username_taken(driver):
    target_text = "The username you have selected is already in use"

    # 1. DOM-based detection (best signal)
    try:
        el = driver.find_element(
            "css selector",
            "div.indi-form__input-notification--has-error.ng-binding"
        )

        if el.is_displayed():
            if target_text in el.text:
                return "USERNAME_TAKEN"

    except:
        pass

    # 2. Angular fallback (ng-bind-html sometimes delays rendering)
    try:
        body_text = driver.find_element("tag name", "body").text.lower()

        if target_text.lower() in body_text:
            return "USERNAME_TAKEN"

    except:
        pass

    # 3. page_source fallback (last resort)
    if target_text.lower() in driver.page_source.lower():
        return "USERNAME_TAKEN"

    return None


# ----------------------------
# Navigation helpers
# ----------------------------

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

    print("Recovery: waiting before next lead...")
    random_pause(1, 3)


# ----------------------------
# Page progress / recovery helpers
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
            wait_for_dom_ready(driver, timeout=20)
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


# ----------------------------
# Result classification / sheet helpers
# ----------------------------

def classify_submission_result(driver):
    try:
        body_text = driver.execute_script(
            "return document.body.innerText.toLowerCase();"
        )

        if "approved" in body_text:
            return "APPROVED", "Application approved"

        elif "pending" in body_text or "under review" in body_text:
            return "PENDING", "Under review"

        elif "rejected" in body_text or "not eligible" in body_text:
            return "REJECTED", "Not eligible"

        elif "application submitted" in body_text:
            return "SUBMITTED", "Submission confirmed"

        return "UNKNOWN", "No clear result detected"

    except Exception as e:
        return "ERROR", f"Classifier failed: {str(e)[:100]}"


def should_skip_row(row):
    status = (row.get("Status") or "").strip().lower()
    return status in {"done", "submitted", "skip"}


# ----------------------------
# Error / condition detectors
# ----------------------------

def detect_duplicate_account(driver):

    target_text = "already exist in our system"

    # 1. DOM-based detection (best signal)
    try:
        el = driver.find_element(
            "css selector",
            "div[ng-if='pageForm.$error.duplicatePii']"
        )

        if el.is_displayed():
            if target_text in el.text.lower():
                return "DUPLICATE_ACCOUNT"

    except:
        pass

    # 2. Angular fallback (ng-bind-html sometimes delays rendering)
    try:
        body_text = driver.find_element("tag name", "body").text.lower()

        if target_text in body_text:
            return "DUPLICATE_ACCOUNT"

    except:
        pass

    # 3. page_source fallback (last resort)
    if target_text in driver.page_source.lower():
        return "DUPLICATE_ACCOUNT"

    return None

def detect_needs_more_info(driver):
    target = "we need more information to see if you qualify"

    reasons = []

    # 1. PRIMARY: full rendered text
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text.lower()

        if target in body_text:
            # Extract reasons from the fieldset
            try:
                fieldset = driver.find_element(By.CSS_SELECTOR, "fieldset.indi-form__fieldset")
                items = fieldset.find_elements(By.CSS_SELECTOR, "li")

                for item in items:
                    text = item.text.strip()
                    if text:
                        reasons.append(text)

            except Exception:
                pass

            return {
                "status": "WE_NEED_MORE_INFO_TO_QUALIFY",
                "reasons": reasons
            }

    except Exception:
        pass

    # 2. SECONDARY: page source fallback
    try:
        if target in driver.page_source.lower():
            return {
                "status": "WE_NEED_MORE_INFO_TO_QUALIFY",
                "reasons": []
            }
    except Exception:
        pass

    return None


def write_column_m(sheet, row_number, value):
    if isinstance(value, dict):
        value = str(value)

    if value is None:
        value = ""

    sheet.update_cell(row_number, 13, str(value))


def detect_almost_qualified(driver):
    target = "you are almost done qualifying"

    # 1. PRIMARY: visible DOM text (best for Angular)
    try:
        elements = driver.find_elements(
            "css selector",
            "p.indi-long-form-text__p--intro"
        )

        for el in elements:
            if el.is_displayed():
                text = (el.text or "").lower()
                if target in text:
                    return "GOOD_LEAD"

    except:
        pass

    # 2. FALLBACK: full page text
    try:
        body_text = driver.find_element("tag name", "body").text.lower()
        if target in body_text:
            return "GOOD_LEAD"
    except:
        pass

    # 3. LAST RESORT: raw HTML
    try:
        if target in driver.page_source.lower():
            return "GOOD_LEAD"
    except:
        pass

    return None


# ----------------------------
# Data normalization helpers
# ----------------------------

def clean(v):
    if isinstance(v, dict):
        return str(v)
    if v is None:
        return ""
    return str(v).strip()    

def row_from_values(values):
    safe = []

    for v in values:
        if isinstance(v, dict):
            safe.append(str(v))
        elif v is None:
            safe.append("")
        else:
            safe.append(str(v))

    safe = safe + [""] * (12 - len(safe))

    return {
        "First Name": safe[0],
        "Last Name": safe[1],
        "DOB": safe[2],
        "SSN4": safe[3],
        "address1": safe[4],
        "address2": safe[5],
        "City": safe[6],
        "State": safe[7],
        "Zip": safe[8],
        "Email": safe[9],
        "Status": safe[10],
        "Note": safe[11],
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
# Application page actions
# ----------------------------

def click_start_lifeline_application(driver, timeout=30):
    print("Clicking Start Lifeline Application...")

    btn = wait_clickable_xpath(
        driver,
        "//button[contains(normalize-space(), 'Start Lifeline Application')]",
        timeout=timeout
    )

    random_pause(0.25, 0.75)  # human hesitation before click

    try:
        click_element_hybrid(driver, element=btn, use_real_mouse=True)
    except Exception:
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)

    random_pause(0.2, 1)


def extract_username_from_email(email):
    email = (email or "").strip()
    if "@" in email:
        return email.split("@")[0]
    return email


def fill_create_account_page(driver, row):
    wait_for_dom_ready(driver, timeout=30)

    email_full = (row.get("Email") or "").strip()
    username = extract_username_from_email(email_full)
    password = "Water123!!"

    print("Typing username...")
    strong_type_css_human_first(
        driver,
        "#username",
        username,
        timeout=30,
        use_real_mouse=True
    )

    print("Typing password...")
    strong_type_css_human_first(
        driver,
        "#password",
        password,
        timeout=30,
        use_real_mouse=False
    )

    print("Typing confirm password...")
    strong_type_css_human_first(
        driver,
        "#password2",
        password,
        timeout=30,
        use_real_mouse=False
    )

    print("Typing email...")
    strong_type_css_human_first(
        driver,
        "#email",
        email_full,
        timeout=30,
        use_real_mouse=False
    )

    # Checkbox click
    print("Clicking agreement checkbox...")
    click_terms_checkbox(driver)

    random_pause(0.6, 1)
    if not click_consumer_submit(driver, timeout=20):
        return False, "consumer next failed"

    skip, reason = check_duplicate_account(driver)

    if skip:
        print(f"Skipping lead due to: {reason}")

        return False, "NEXT_LEAD"
    
    return True, "Account created"

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

    return False


def sign_out_account(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)

    print("Opening My account menu...")

    try:
        # ✅ Click the REAL button (not the span)
        my_account_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "MyAccount_btn"))
        )

        click_element_hybrid(driver, element=my_account_btn, use_real_mouse=True)

        # wait for dropdown to OPEN (critical)
        wait.until(lambda d: "in" in d.find_element(By.ID, "CollapseMy_Acc").get_attribute("class"))

        random_pause(0.4, 1.0)

    except TimeoutException:
        print("Could not open My account menu")
        return False

    print("Clicking Sign out...")

    try:
        # ✅ target the ng-click directly (most stable)
        sign_out = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[@ng-click='signOut();']"
            ))
        )

        click_element_hybrid(driver, element=sign_out, use_real_mouse=True)

        random_pause(1.0, 2.0)

        print("Signed out successfully")
        return True

    except TimeoutException:
        print("Could not find 'Sign out' option")
        return False


def start_browser():
    sb = SB(
        uc=True,
        test=True,
        agent=USER_AGENT,
        maximize=True
    )

    driver = sb.driver
    driver.set_page_load_timeout(60)

    return sb, driver


def click_consumer_next(driver):

    # optional human delay
    if random.random() < 0.4:
        time.sleep(random.uniform(1.5, 3.5))

    print("Checking for invalid fields...")
    print_invalid_fields(driver)

    before_url = driver.current_url

    print("Waiting for consumerNextSuccessButton to be clickable...")

    def is_ready(d):
        try:
            btn = d.find_element(By.ID, "consumerNextSuccessButton")

            return (
                btn.is_displayed()
                and not btn.get_attribute("disabled")
                and "ng-hide" not in (btn.get_attribute("class") or "")
                and btn.size["height"] > 0
                and btn.size["width"] > 0
            )
        except:
            return False

    # wait until Angular actually makes it active
    WebDriverWait(driver, 15).until(is_ready)

    btn = driver.find_element(By.ID, "consumerNextSuccessButton")

    print("Clicking consumerNextSuccessButton...")

    # safe click (bypasses overlays / Angular quirks)
    driver.execute_script("arguments[0].click();", btn)

    try:
        wait_for_page_advance(driver, before_url, timeout=8)
        return True, "Next clicked successfully"
    except TimeoutException:
        debug_post_click_state(driver, before_url)
        return False, "Next click did not advance"


def angular_safe_click(driver, locator, timeout=15, text=None):
    """
    Clicks an element safely in AngularJS/React-like dynamic UIs.

    Args:
        driver: Selenium driver
        locator: tuple (By.ID / By.CSS_SELECTOR / etc, value)
        timeout: max wait time
        text: optional filter if multiple buttons match

    Returns:
        WebElement that was clicked
    """

    wait = WebDriverWait(driver, timeout)

    def find_clickable(d):
        try:
            elements = d.find_elements(*locator)

            for el in elements:

                # 1. Must be visible
                if not el.is_displayed():
                    continue

                # 2. Must not be disabled (HTML)
                if el.get_attribute("disabled"):
                    continue

                # 3. Must not be Angular-hidden
                classes = el.get_attribute("class") or ""
                if "ng-hide" in classes:
                    continue

                # 4. Optional text filter
                if text and text not in el.text:
                    continue

                return el

        except StaleElementReferenceException:
            return False

        return False

    element = wait.until(find_clickable)

    # Retry-safe click
    try:
        element.click()
    except (ElementClickInterceptedException, StaleElementReferenceException):
        driver.execute_script("arguments[0].click();", element)

    return element

# ----------------------------
# Main row processor
# ----------------------------

def fill_form_from_row(driver, row, sheet, row_number):
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

    

    # Block all WebSocket connections on the page at the JS level.
    # The fake immediately fires onerror + onclose so ServiceNow stops
    # waiting and falls back to HTTP polling instead of hanging.
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            (function() {
                const _WS = window.WebSocket;
                function FakeWebSocket(url, protocols) {
                    this.url = url;
                    this.readyState = 3;
                    this.onopen = null;
                    this.onclose = null;
                    this.onerror = null;
                    this.onmessage = null;
                    const self = this;
                    setTimeout(function() {
                        const err = new Event('error');
                        if (self.onerror) self.onerror(err);
                        const cls = new CloseEvent('close', { code: 1006, reason: 'blocked', wasClean: false });
                        if (self.onclose) self.onclose(cls);
                    }, 0);
                }
                FakeWebSocket.prototype.send = function() {};
                FakeWebSocket.prototype.close = function() {};
                FakeWebSocket.prototype.addEventListener = function() {};
                FakeWebSocket.prototype.removeEventListener = function() {};
                FakeWebSocket.CONNECTING = 0;
                FakeWebSocket.OPEN      = 1;
                FakeWebSocket.CLOSING   = 2;
                FakeWebSocket.CLOSED    = 3;
                window.WebSocket = FakeWebSocket;
                console.log('[WS blocked] WebSocket replaced with FakeWebSocket');
            })();
        """
    })

    # Also block at the network level so the TCP handshake never starts.
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setBlockedURLs", {
            "urls": ["wss://*", "ws://*"]
        })
    except Exception:
        pass

    #navigate to website
    driver.get("https://www.getinternet.gov/apply?id=nv_flow&ln=RW5nbGlzaA%3D%3D")
    wait_for_dom_ready(driver, timeout=30)
    random_pause(1.0, 2.0)
    driver.get("https://www.getinternet.gov/apply?id=nv_flow&ln=RW5nbGlzaA%3D%3D")
    wait_for_dom_ready(driver, timeout=30)

    # Wait for ServiceNow session to establish before touching the form.
    # The 428 on /api/now/sp/rectangle means the session token isn't ready yet.
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("""
            return typeof window.NOW !== 'undefined' &&
                   document.querySelector('#firstName') !== null;
        """)
    )
    random_pause(5,10)

    

    # ----------------------------
    # Prepare data for fill in information page
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
    # Name (ActionChains version)
    # ----------------------------
    name_fields = [
        ("#firstName", first_name),
        ("#lastName", last_name),
    ]

    random.shuffle(name_fields)

    for css, value in name_fields:
        el = wait_clickable_css(driver, css)
        angular_human_type(driver, el, value)
        human_behavior()


    # ----------------------------
    # DOB (ActionChains version)
    # ----------------------------
    angular_human_type(driver, wait_clickable_css(driver, "#dobMonth"), dob_month)

    if random.random() < 0.3:
        time.sleep(random.uniform(0.5, 1.5))

    angular_human_type(driver, wait_clickable_css(driver, "#dobDay"), dob_day)
    angular_human_type(driver, wait_clickable_css(driver, "#dobYear"), dob_year)

    human_behavior()


    # ----------------------------
    # SSN
    # ----------------------------
    angular_human_type(driver, wait_clickable_css(driver, "#ssnText"), ssn4)


    # ----------------------------
    # Address block
    # ----------------------------
    angular_human_type(driver, wait_clickable_css(driver, "#streetNumberName"), address1)

    if address2:
        if random.random() < 0.6:
            angular_human_type(driver, wait_clickable_css(driver, "#apt"), address2)

    human_behavior()

    angular_human_type(driver, wait_clickable_css(driver, "#city"), city)


    # ----------------------------
    # State selection (unchanged - good as-is)
    # ----------------------------
    if random.random() < 0.5:
        time.sleep(random.uniform(0.6, 2.0))

    strong_select_state(
        driver,
        state,
        timeout=30,
        use_real_mouse=(random.random() < 0.4)
    )


    # ----------------------------
    # ZIP code
    # ----------------------------
    angular_human_type(driver, wait_clickable_css(driver, "#zipcode"), zip_code)

    if random.random() < 0.15:
        angular_human_type(driver, wait_clickable_css(driver, "#zipcode"), zip_code)

    human_behavior()
    # ----------------------------
    # Pre-submit hesitation
    # ----------------------------
    if random.random() < 0.4:
        time.sleep(random.uniform(1.5, 3.5))

    print("Checking for invalid fields...")
    print_invalid_fields(driver)

    # Capture URL before clicking so we can detect navigation afterward
    before_url = driver.current_url

    angular_safe_click(
        driver,
        (By.ID, "consumerNextSuccessButton"),
        timeout=15,
        text="Next"
    )

    print("Clicking first Next...")

    try:
        wait_for_page_advance(driver, before_url, timeout=8)
    except TimeoutException:
        debug_post_click_state(driver, before_url)
        return False, "Next click did not advance"

    
    solver = RecaptchaSolver(driver)

    try:
        solver.solveCaptcha()
        print("Captcha handling finished")
    except Exception as e:
        print("Captcha failed:", e)

    #create an account page
    ok, msg = fill_create_account_page(driver, row)
    if not ok:
        return False, msg
    time.sleep(random.uniform(1.5, 3.5))

    # Check if username already taken
    username_status = detect_username_taken(driver)
    
    if username_status == "USERNAME_TAKEN":
        print("Username already taken, skipping lead.")
        write_column_m(sheet, row_number, username_status)
        return False, "NEXT_LEAD"


    status = detect_duplicate_account(driver)

    if status == "DUPLICATE_ACCOUNT":
        sheet.update_cell(row_number, 13, status)
        return False, "NEXT_LEAD"
    

    time.sleep(random.uniform(1.5, 3.5))
    #start application page
    click_start_lifeline_application(driver)

    ## ----------------------------
    # SNAP checkbox (with reading delay)
    # ----------------------------
    human_behavior()
    click_checkbox_like_human(driver, "#eligSnapSpan", timeout=20)

    # simulate reading eligibility text
    time.sleep(random.uniform(1, 2))

    # ----------------------------
    # Gov program Next
    # ----------------------------
    if not click_gov_program_next(driver, timeout=20):
        return False, "Gov program next failed"
    
    # ----------------------------
    # Agreement checkbox
    # ----------------------------
    time.sleep(random.uniform(2,4))
    solver = RecaptchaSolver(driver)

    try:
        time.sleep(random.uniform(2, 3))
        solver.solveCaptcha()
        print("Captcha handling finished")
    except Exception as e:
        print("Captcha failed:", e)

    print("Captcha completed, continuing...")

    human_behavior()
    time.sleep(random.uniform(.5,1.5))
    driver.execute_script("window.scrollBy(0, 200);")
    click_checkbox_like_human(driver, "span.indi-form__checkbox-icon", timeout=20)

    # simulate reading terms
    time.sleep(random.uniform(1,2))

    # ----------------------------
    # Submit
    # ----------------------------
    human_click("#nextSuccessButton9")

    # loading watchdog
    ok = wait_for_loader_or_timeout(driver, timeout=20)
    if not ok:
        return False, "Loading stuck → recovered"
    
    if status:
        print("Good lead detected")

        write_column_m(sheet, row_number, status)

        sign_out_account(driver)

        return False, "GODD_NEXT_LEAD"
    


    # post-submit idle (user waiting / reading)

    status = detect_needs_more_info(driver)

    if status:
        write_column_m(sheet, row_number, status)
        sign_out_account(driver)
        return False, "NEXT_LEAD"
    # ----------------------------
    # sign out back to homepage
    # ----------------------------
    sign_out_account(driver)

    time.sleep(random.uniform(1,2))

    return True

# ----------------------------
# Orchestrator
# ----------------------------
def process_single_lead(row, sheet, row_number):

    with SB(
        uc=True,
        test=True,
        agent=USER_AGENT,
        maximize=True
    ) as sb:

        driver = sb.driver

        try:
            print(f"Processing row {row_number}")

            ok, result = fill_form_from_row(driver, row, sheet, row_number)

            if ok:
                status, note = result
                update_status(sheet, row_number, status, note)
                print(f"Row {row_number} completed → {status}")
            else:
                update_status(sheet, row_number, "FAILED", result)

            return ok

        except Exception as e:
            print(f"Error row {row_number}: {e}")
            update_status(sheet, row_number, "ERROR", str(e)[:200])
            return False

def process_rows():
    sheet = connect_sheet()
    all_rows = sheet.get_all_values()
    data_rows = all_rows[1:]

    print("Starting per-lead browser system...")

    for row_number, values in enumerate(data_rows, start=2):
        row = row_from_values(values)

        if should_skip_row(row):
            print(f"Skipping row {row_number}")
            continue

        print(f"\n=== NEW LEAD {row_number} ===")

        try:
            success = process_single_lead(row, sheet, row_number)

            if not success:
                print(f"Lead {row_number} did not complete successfully")

        except Exception as e:
            print(f"Fatal error on row {row_number}: {e}")

        wait_before_next_lead(7, 15)

if __name__ == "__main__":
    process_rows()
