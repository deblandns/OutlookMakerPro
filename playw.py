import os
import random
from faker import Faker
from loguru import logger
from fake_useragent import FakeUserAgent
from patchright.async_api import async_playwright
from playwright_stealth import stealth_async
import string
import secrets
import asyncio # Ensure asyncio is imported at the top
from tenacity import retry as tenacity_retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

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
        browser = await p.chromium.launch(headless=False)
        # Create a new browser context with the random user agent
        context = await browser.new_context(user_agent=FakeUserAgent().random)
        page = await context.new_page()
        await stealth_async(page)

        # Intercept network requests to block images, media, fonts, and stylesheets
        async def handle_route(route):
            if route.request.resource_type in ["image", "media", "font", "stylesheet"]:
                await route.abort()
            else:
                await route.continue_()
        
        await page.route("**/*", handle_route)

        try: # this code will try to go to the signup page
            await robust_goto(page, "https://signup.live.com/signup")
        except Exception as e:
            error_log(f"Error: {e}")
            await browser.close()
            return
        # input the email to the first page
        random_data = get_random_data()
        email_input_selector = "#floatingLabelInput5"
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
            password = generate_strong_password()
            info_log("Hovering over the password input.")
            await robust_hover(page, password_input_locator, timeout=5000)
            info_log("Clicking the password input.")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_click(password_input_locator, timeout=5000)
            info_log(f"Attempting to type password: {password}")
            await asyncio.sleep(random.uniform(1.0, 3.0))
            await robust_type(password_input_locator, text=password, delay=random.uniform(90, 300), timeout=15000)
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
        except Exception as e:
            error_log(f"Error interacting with name and family input: {e}")
            await browser.close()
            return
        await page.wait_for_timeout(500000)
        # finally it will close when everything is done
        await browser.close()

# this code will run the main function
if __name__ == "__main__":
    # import asyncio # asyncio already imported at the top
    asyncio.run(main())

