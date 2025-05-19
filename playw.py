import os
import random
from faker import Faker
from loguru import logger
from fake_useragent import FakeUserAgent
from patchright.async_api import async_playwright
# from playwright_stealth import stealth_async # Removed
import string
import secrets
import asyncio # Ensure asyncio is imported at the top
from tenacity import retry as tenacity_retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from playwright.async_api import TimeoutError as PlaywrightTimeoutError
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException # For checking if file is valid excel

# the logger configs
info_log = logger.info
success_log = logger.success
error_log = logger.error
debug_log = logger.debug
warning_log = logger.warning

# the root name and family txt files
dir_path = os.path.dirname(os.path.abspath(__file__))
name_file = os.path.join(dir_path, "usersdata/names.txt")
family_file = os.path.join(dir_path, "usersdata/family.txt")

# Excel configuration
EXCEL_FILE_NAME = "created_outlook_accounts.xlsx"
EXCEL_HEADERS = ["Email", "Password", "First Name", "Last Name", "Birth Date"]

def save_account_to_excel(email, password, first_name, last_name, birth_date_str):
    workbook = None
    sheet = None
    file_exists_and_is_accessible = os.path.exists(EXCEL_FILE_NAME) and os.access(EXCEL_FILE_NAME, os.W_OK)

    try:
        if file_exists_and_is_accessible:
            try:
                workbook = openpyxl.load_workbook(EXCEL_FILE_NAME)
                sheet = workbook.active
            except InvalidFileException:
                info_log(f"{EXCEL_FILE_NAME} is not a valid Excel file or is corrupted. Creating a new one.")
                os.remove(EXCEL_FILE_NAME) # Attempt to remove corrupted file
                workbook = Workbook()
                sheet = workbook.active
                sheet.append(EXCEL_HEADERS)
            except Exception as e: # Other errors during load (e.g. file locked)
                error_log(f"Could not load {EXCEL_FILE_NAME}: {e}. Will try to create new if it was removed or write to a new one.")
                # If removal failed or wasn't attempted, try creating a new one.
                if not os.path.exists(EXCEL_FILE_NAME): # If it was removed or never existed
                    workbook = Workbook()
                    sheet = workbook.active
                    sheet.append(EXCEL_HEADERS)
                else: # File still exists but couldn't be loaded, avoid overwriting potentially good data with just headers
                    error_log(f"Cannot proceed with saving to {EXCEL_FILE_NAME} due to load error and file still existing.")
                    return # Exit if we can't be sure about the sheet's state
        else: # File does not exist or is not writable
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(EXCEL_HEADERS)

        # Ensure sheet is valid and headers are present if it's an existing sheet
        if sheet is not None and sheet.max_row > 0:
            current_headers = [sheet.cell(row=1, column=i+1).value for i in range(len(EXCEL_HEADERS))]
            if current_headers != EXCEL_HEADERS:
                # This case is tricky: headers mismatch. Could be a different file.
                # For simplicity, we'll append, but ideally, this might need more sophisticated handling or a new sheet/file.
                warning_log(f"Header mismatch in {EXCEL_FILE_NAME}. Appending data anyway.")
                if sheet.max_row == 1 and all(c is None for c in current_headers): # Empty first row
                    for i, header_text in enumerate(EXCEL_HEADERS):
                        sheet.cell(row=1, column=i + 1, value=header_text)
        elif sheet is not None and (sheet.max_row == 0 or sheet.cell(row=1, column=1).value is None): # Sheet is empty or first cell is empty
             sheet.append(EXCEL_HEADERS)


        if sheet is None: # If sheet could not be initialized
            error_log(f"Failed to obtain a valid sheet in {EXCEL_FILE_NAME}. Cannot save account.")
            return

        sheet.append([email, password, first_name, last_name, birth_date_str])
        workbook.save(EXCEL_FILE_NAME)
        success_log(f"Successfully saved account {email} to {EXCEL_FILE_NAME}")

    except PermissionError:
        error_log(f"Permission denied when trying to save {EXCEL_FILE_NAME}. Make sure the file is not open in Excel or another program and you have write permissions.")
    except Exception as e:
        error_log(f"An unexpected error occurred while saving to Excel: {e}")


# function to get the random name and family
def get_random_data():
    with open(name_file, "r") as f:
        names = f.readlines()
    first_name = random.choice(names).strip()
    with open(family_file, "r") as f:
        families = f.readlines()
    last_name = random.choice(families).strip()
    birth_date = Faker().date_of_birth(minimum_age=18, maximum_age=80)
    email_username = f"{first_name}{last_name}{random.randint(1, 999999)}"
    return {"first_name": first_name, "last_name": last_name, "email_username": email_username, "birth_date": birth_date}

def generate_strong_password(length=12):
    if length < 4: length = 12 # Ensure min length for variety, default to 12
    char_sets = [string.ascii_lowercase, string.ascii_uppercase, string.digits, string.punctuation]
    # Start password with one character from each set
    password_list = [secrets.choice(s) for s in char_sets]
    # Characters for the rest of the password
    all_chars = "".join(char_sets)
    # Fill remaining length if length > 4
    if length > len(password_list):
        password_list.extend(secrets.choice(all_chars) for _ in range(length - len(password_list)))
    elif length < len(password_list): # Should not happen if length >=4
        password_list = password_list[:length]
    random.shuffle(password_list)
    return "".join(password_list)

# Default retry settings for tenacity
RETRY_WAIT = wait_exponential(multiplier=1, min=1, max=5) # Wait 1s, then 2s, then 4s (max 5s for any single wait)
RETRY_STOP = stop_after_attempt(3) # Try 3 times in total
RETRY_ON_EXCEPTION = retry_if_exception_type(PlaywrightTimeoutError)

# Helper functions with tenacity
@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_goto(page, url, **kwargs):
    info_log(f"Attempting to navigate to {url} with retries...")
    await page.goto(url, **kwargs)
    success_log(f"Successfully navigated to {url}")

@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_wait_for_selector(page, selector, **kwargs):
    info_log(f"Attempting to wait for selector '{selector}' with retries...")
    element = await page.wait_for_selector(selector, **kwargs)
    success_log(f"Successfully found selector '{selector}'")
    return element

@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_hover(page, locator, **kwargs):
    info_log(f"Attempting to move mouse over locator '{locator}' with retries...")
    try:
        await locator.scroll_into_view_if_needed(timeout=5000) # Scroll into view first
    except Exception as e:
        warning_log(f"Failed to scroll locator '{locator}' into view for hover: {e}. Proceeding with hover attempt.")

    # Get the bounding box of the element
    bounding_box = await locator.bounding_box()
    if bounding_box:
        # Calculate the center of the element
        x = bounding_box['x'] + bounding_box['width'] / 2
        y = bounding_box['y'] + bounding_box['height'] / 2
        # Move the mouse to the center of the element
        await page.mouse.move(x, y)
        success_log(f"Successfully moved mouse over locator '{locator}'")
    else:
        # Fallback or error if bounding_box is None (element might not be visible/attached)
        warning_log(f"Could not get bounding_box for locator '{locator}'. Falling back to hover().")
        try:
            await locator.scroll_into_view_if_needed(timeout=5000) # Ensure in view for fallback too
        except Exception as e:
            warning_log(f"Failed to scroll locator '{locator}' into view for fallback hover: {e}")
        await locator.hover(**kwargs) # fallback to original hover
        success_log(f"Successfully hovered over locator '{locator}' (fallback)")

@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_click(locator, **kwargs):
    info_log(f"Attempting to click locator '{locator}' with retries...")
    try:
        await locator.scroll_into_view_if_needed(timeout=5000)
    except Exception as e:
        warning_log(f"Failed to scroll locator '{locator}' into view for click: {e}. Proceeding with click attempt.")
    await locator.click(**kwargs)
    success_log(f"Successfully clicked locator '{locator}'")

@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_type(locator, text, **kwargs):
    original_timeout = kwargs.get('timeout')
    original_delay = kwargs.get('delay')

    log_message = f"Attempting to type text '{text[:20]}...' into locator '{locator}' char-by-char with random delays."
    if original_delay is not None:
        # Inform that the 'delay' kwarg is now handled differently
        log_message += f" (Note: 'delay={original_delay}' for locator.type is superseded by internal per-char random delays)."
    if original_timeout is not None:
        # Inform that 'timeout' for the whole operation is now mainly governed by tenacity
        log_message += f" (Note: Overall operation timeout is primarily handled by tenacity retries, not 'timeout={original_timeout}' for a single type op)."
    info_log(log_message)

    try:
        await locator.scroll_into_view_if_needed(timeout=5000) # Standard timeout for scrolling
    except Exception as e:
        warning_log(f"Failed to scroll locator '{locator}' into view for type: {e}. Proceeding with type attempt.")

    try:
        # Click to focus the element before typing. Use a reasonable timeout for the click itself.
        await locator.click(timeout=kwargs.get('click_timeout', 5000)) # Allow override or use default
    except Exception as e:
        warning_log(f"Failed to click/focus locator '{locator}' before typing: {e}. Attempting to type anyway.")

    for char_index, char_to_press in enumerate(text):
        try:
            await locator.press(char_to_press)
            # Randomized delay after each character press (50ms to 250ms)
            char_delay = random.uniform(0.05, 0.25)
            await asyncio.sleep(char_delay)
        except PlaywrightTimeoutError as e:
            error_log(f"Playwright TimeoutError while pressing character '{char_to_press}' (index {char_index}) for locator '{locator}': {e}")
            raise # Re-raise to trigger tenacity retry for the whole robust_type
        except Exception as e:
            error_log(f"Error pressing character '{char_to_press}' (index {char_index}) for locator '{locator}': {e}")
            raise # Re-raise to trigger tenacity

    success_log(f"Successfully typed text into locator '{locator}' using char-by-char method.")

@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_select_option(locator, value, **kwargs):
    info_log(f"Attempting to select option '{value}' for locator '{locator}' with retries...")
    try:
        await locator.scroll_into_view_if_needed(timeout=5000)
    except Exception as e:
        warning_log(f"Failed to scroll locator '{locator}' into view for select_option: {e}. Proceeding with select_option attempt.")
    await locator.select_option(value=value, **kwargs)
    success_log(f"Successfully selected option '{value}' for locator '{locator}'")

@tenacity_retry(wait=RETRY_WAIT, stop=RETRY_STOP, retry=RETRY_ON_EXCEPTION)
async def robust_wait_for_load_state(page, state="domcontentloaded", **kwargs):
    info_log(f"Attempting to wait for load state '{state}' with retries...")
    await page.wait_for_load_state(state, **kwargs)
    success_log(f"Successfully reached load state '{state}'")

# function to get the random date of birt and region

# this code will open the browser and go to the signup page
async def main():
    async with async_playwright() as p:
        info_log("Starting the browser")
        # this code will open the browser and go to the signup page
        browser = await p.chromium.launch(headless=False) # Removed args=launch_args to use patchright defaults

        # Define common viewport sizes
        common_viewports = [{'width': 1920, 'height': 1080}, {'width': 1366, 'height': 768}, {'width': 1280, 'height': 720}, {'width': 1600, 'height': 900}, {'width': 1440, 'height': 900}, {'width': 1280, 'height': 800}, {'width': 1280, 'height': 1024},{'width': 1024, 'height': 768}]
        # Select a random viewport
        random_viewport = random.choice(common_viewports)
        info_log(f"Setting random viewport: {random_viewport['width']}x{random_viewport['height']}")

        context = await browser.new_context(user_agent=FakeUserAgent().random, viewport=random_viewport)
        # await context.add_init_script(init_script) # Ensure it runs on every new document

        page = await context.new_page()
        # await stealth_async(page) # Removed

        # Intercept network requests to block images, media, fonts, and stylesheets
        async def handle_route(route):
            if route.request.resource_type in ["image", "media"]: # Allow fonts and stylesheets
                await route.abort()
            else:
                await route.continue_()
        
        await page.route("**/*", handle_route)

        try: # this code will try to go to the signup page
            await robust_goto(page, "https://signup.live.com/signup")
        except Exception as e:
            error_log(f"Error navigating to signup page: {e}")
            await browser.close()
            return
        # input the email to the first page
        random_data = get_random_data()
        email_input_selector = 'input[aria-label="Email"]'
        email_input = page.locator(email_input_selector)
        info_log(f"Attempting to find email input with selector: {email_input_selector}")
        # this code will try to input the email to the first page
        try:
            # Best practice: Hover to simulate mouse movement
            info_log("Hovering over the email input.")
            await robust_hover(page, email_input, timeout=5000) # Wait up to 5s for hover
            # Best practice: Click to focus (fill also does this, but explicit click is fine)
            info_log("Clicking the email input.")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(email_input, timeout=5000) # Wait up to 5s for click
            # Type the email into the input field, simulating human behavior
            info_log(f"Attempting to type email: {random_data['email_username']}") # Updated log message
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_type(email_input, f"{random_data['email_username']}@outlook.com", delay=random.uniform(90, 300), timeout=15000)
            success_log("Successfully typed the email input.") # Updated log message
            # Click the next button
            next_button_selector = "[data-testid='primaryButton']"
            next_button = page.locator(next_button_selector)
            info_log(f"Attempting to find 'Next' button with selector: {next_button_selector}")
            # Hover to simulate human-like mouse movement
            info_log("Hovering over the 'Next' button.")
            await robust_hover(page, next_button, timeout=5000)  # Wait up to 5 seconds for hover
            # Click the 'Next' button
            info_log("Clicking the 'Next' button.")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(next_button, timeout=5000)  # Wait up to 5 seconds for click
            success_log("Successfully clicked the 'Next' button.")
            
            info_log("Waiting for the next page to load (password page expected).")
            try:
                await robust_wait_for_load_state(page, 'domcontentloaded', timeout=15000) # Wait for DOM content
                success_log("Next page DOM content loaded.")
            except Exception as e:
                error_log(f"Error waiting for page load state after clicking next: {e}")
                # Potentially add a screenshot here for debugging if it fails often
                # await page.screenshot(path="error_screenshot_pageload.png")
                await browser.close()
                return

        except Exception as e:
            error_log(f"Error interacting with email input: {e}")
            await browser.close()
            return
        # this code will try to insert the password to the second page
        try:
            info_log("Waiting for the password input to appear")
            # Wait for the password input field to be present and visible
            password_selector = 'input[type="password"][autocomplete="new-password"]'
            await robust_wait_for_selector(page, password_selector, timeout=20000) # Wait up to 20s
            password_input_locator = page.locator(password_selector)
            password_to_save = generate_strong_password() # Assign the generated password
            info_log("Hovering over the password input.")
            await robust_hover(page, password_input_locator, timeout=5000)
            info_log("Clicking the password input.")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(password_input_locator, timeout=5000)
            info_log(f"Attempting to type password...") # Removed password from log for security
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_type(password_input_locator, text=password_to_save, delay=random.uniform(90, 300), timeout=15000) # Use password_to_save
            next_button_password_selector = "[data-testid='primaryButton']"
            next_button_password = page.locator(next_button_password_selector)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_hover(page, next_button_password, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(next_button_password, timeout=5000)
            success_log("Successfully clicked the 'Next' button after password input.")
            info_log("Waiting for the birth date page to load.")
            try:
                await robust_wait_for_load_state(page, 'domcontentloaded', timeout=15000) # Wait for DOM content
                success_log("Birth date page DOM content loaded.")
            except Exception as e:
                error_log(f"Error waiting for birth date page load state: {e}")
                await browser.close()
                return

        except Exception as e:
            error_log(f"Error interacting with password input: {e}")
            await browser.close()
            return
        # give the mont, day, year and the region to the third page
        try:
            info_log("Waiting for the birth month dropdown to appear.")
            month_selector = 'select[name="BirthMonth"]'
            await robust_wait_for_selector(page, month_selector, timeout=10000) # Wait up to 10s
            
            birth_date_obj = random_data['birth_date']
            birth_month_value = str(birth_date_obj.month)
            info_log(f"Attempting to select birth month: {birth_month_value}")
            month_dropdown_locator = page.locator(month_selector)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_hover(page, month_dropdown_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_select_option(month_dropdown_locator, value=birth_month_value)
            success_log(f"Successfully selected birth month: {birth_month_value}")

            # Select Day
            info_log("Waiting for the birth day dropdown to appear.")
            day_selector = 'select[name="BirthDay"]'
            await robust_wait_for_selector(page, day_selector, timeout=10000)
            day_dropdown_locator = page.locator(day_selector)
            birth_day_value = str(birth_date_obj.day) # birth_date_obj is already defined
            info_log(f"Attempting to select birth day: {birth_day_value}")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_hover(page, day_dropdown_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_select_option(day_dropdown_locator, value=birth_day_value)
            success_log(f"Successfully selected birth day: {birth_day_value}")
            
            # Input Year
            info_log("Waiting for the birth year input to appear.")
            year_selector = 'input[name="BirthYear"]'
            await robust_wait_for_selector(page, year_selector, timeout=10000)
            year_input_locator = page.locator(year_selector)
            birth_year_value = str(birth_date_obj.year) # birth_date_obj is already defined
            info_log(f"Attempting to input birth year: {birth_year_value}")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_hover(page, year_input_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(year_input_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_type(year_input_locator, text=birth_year_value, delay=random.uniform(90, 250))
            success_log(f"Successfully inputted birth year: {birth_year_value}")
            # Click the Next button after birth date input
            info_log("Attempting to find the 'Next' button after birth date input.")
            next_button_birthdate_selector = "[data-testid='primaryButton']"
            next_button_birthdate = page.locator(next_button_birthdate_selector)
            info_log(f"Found 'Next' button with selector: {next_button_birthdate_selector}")
            info_log("Hovering over the 'Next' button (after birth date).")
            await robust_hover(page, next_button_birthdate, timeout=5000)
            info_log("Clicking the 'Next' button (after birth date).")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(next_button_birthdate, timeout=5000)
            success_log("Successfully clicked the 'Next' button after birth date input.")

            info_log("Waiting for the name input page to load.")
            try:
                await robust_wait_for_load_state(page, 'domcontentloaded', timeout=15000)
                success_log("Name input page DOM content loaded.")
            except Exception as e:
                error_log(f"Error waiting for name input page load state: {e}")
                await browser.close()
                return

        except Exception as e:
            error_log(f"Error interacting with birth date input: {e}")
            await browser.close()
            return
        # Input first name and last name for the next page
        try:
            info_log("Waiting for first name input to appear.")
            first_name_selector = "input[id='firstNameInput']"
            await robust_wait_for_selector(page, first_name_selector, timeout=10000)
            first_name_input_locator = page.locator(first_name_selector)
            user_first_name = random_data['first_name']
            info_log(f"Attempting to input first name: {user_first_name}")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_hover(page, first_name_input_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(first_name_input_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_type(first_name_input_locator, text=user_first_name, delay=random.uniform(90, 250))
            success_log(f"Successfully inputted first name: {user_first_name}")

            info_log("Waiting for last name input to appear.")
            last_name_selector = "input[id='lastNameInput']"
            await robust_wait_for_selector(page, last_name_selector, timeout=10000)
            last_name_input_locator = page.locator(last_name_selector)
            user_last_name = random_data['last_name']
            info_log(f"Attempting to input last name: {user_last_name}")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_hover(page, last_name_input_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(last_name_input_locator, timeout=5000)
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_type(last_name_input_locator, text=user_last_name, delay=random.uniform(90, 250))
            success_log(f"Successfully inputted last name: {user_last_name}")
            # Click the Next button after name input
            info_log("Attempting to find the 'Next' button after name input.")
            next_button_name_selector = "[data-testid='primaryButton']"
            next_button_name = page.locator(next_button_name_selector)
            info_log(f"Found 'Next' button with selector: {next_button_name_selector}")
            info_log("Hovering over the 'Next' button (after name input).")
            await robust_hover(page, next_button_name, timeout=5000)
            info_log("Clicking the 'Next' button (after name input).")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(next_button_name, timeout=5000)
            success_log("Successfully clicked the 'Next' button after name input.")
            # the section for the captcha
            info_log("SCRIPT PAUSED: Please solve the CAPTCHA now in the browser window.")
            info_log("The script will wait indefinitely for the page to navigate after you solve the CAPTCHA...")
            # check if the page is loaded
            try:
                # Wait indefinitely for navigation to complete after manual CAPTCHA solving.
                # 'load' ensures the next page is reasonably loaded before proceeding.
                await page.wait_for_load_state('load', timeout=0) # Changed from page.wait_for_navigation
                success_log("CAPTCHA appears to be solved (navigation detected). Checking current page...")
            except PlaywrightTimeoutError: 
                # This should ideally not happen with timeout=0, but included for robustness.
                error_log("Timeout unexpectedly occurred while waiting for navigation after manual CAPTCHA.")
                return # Exit this attempt
            except Exception as e_nav:
                error_log(f"An error occurred while waiting for navigation after manual CAPTCHA: {e_nav}")
                return # Exit this attempt

            # --- Check for privacy notice page (or other outcomes) AFTER manual CAPTCHA ---
            current_url = page.url
            page_title = await page.title()
            info_log(f"Landed on URL: {current_url} - Title: {page_title} (after manual CAPTCHA)")

            if "privacynotice.account.microsoft.com" in current_url:
                info_log("Microsoft account notice page detected. Preparing to save details and click OK.")
                if password_to_save and random_data:
                    email_full = f"{random_data['email_username']}@outlook.com"
                    birth_date_str = random_data['birth_date'].strftime('%Y-%m-%d')
                    save_account_to_excel(email_full, password_to_save, random_data['first_name'], random_data['last_name'], birth_date_str)
                else:
                    error_log("Missing password or random_data, cannot save account details to Excel.")

                ok_button_locator = page.get_by_role("button", name="OK")
                # Fallback selector if get_by_role doesn't work: page.locator("button:has-text('OK')")
                if await ok_button_locator.is_visible(timeout=10000):
                    info_log("Clicking 'OK' on the Microsoft account notice page.")
                    await robust_hover(page, ok_button_locator, timeout=5000) # Hover before click
                    await robust_click(ok_button_locator)
                    # Wait for page to potentially change/settle after clicking OK
                    try:
                        await robust_wait_for_load_state(page, 'domcontentloaded', timeout=10000)
                        success_log("Successfully clicked 'OK' and notice page processed.")
                    except PlaywrightTimeoutError:
                        info_log("Page did not fully reload or change after clicking OK on notice, but proceeding.")
                    except Exception as e_load:
                        warning_log(f"Error waiting for load state after clicking OK on notice: {e_load}. Proceeding.")
                else:
                    warning_log("'OK' button not found on the notice page. Trying to proceed.")
                    await page.screenshot(path=f"notice_ok_button_not_found_{random_data.get('email_username', 'unknown')}.png")
            
            elif "account.microsoft.com" in current_url and "verify" in current_url.lower():
                 warning_log(f"Landed on a verification page unexpectedly after names: {current_url}. This might indicate CAPTCHA or other challenge.")

            else:
                warning_log(f"Did not land on the expected Microsoft account notice page. Current URL: {current_url}, Title: {page_title}")

        except Exception as e:
            error_log(f"Error interacting with name and family input OR on the subsequent notice/final page: {e}")
        finally:
            info_log("Account creation attempt finished for this iteration. Closing browser.")
            await page.wait_for_timeout(3000) # Brief pause before closing browser.
            await browser.close()
            # No explicit return needed here, end of function implies return

# this code will run the main function
if __name__ == "__main__":
    # import asyncio # asyncio already imported at the top
    asyncio.run(main())