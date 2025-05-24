import questionary
import asyncio
from loguru import logger
from playw import main

# the logger configs
info_log = logger.info
error_log = logger.error
debug_log = logger.debug
warning_log = logger.warning
success_log = logger.success

# the code will ask the user for the starting and ending the bot
question_result = questionary.select("what do you want to do?", choices=["start", "end"]).ask()

# the code will start the bot
if question_result == "start":
    account_creation_number = questionary.text("how many accounts do you want to create?", default="1").ask()
    info_log(f"Creating {account_creation_number} accounts")
    # this code will create the accounts by the number of accounts the user wants to create
    async def run_outlook_creation(num_accounts_to_create):
        successful_creations = 0
        attempts = 0
        max_attempts_per_account = 3 # Max attempts for a single account before moving on or erroring out

        while successful_creations < int(num_accounts_to_create):
            info_log(f"Attempting to create account {successful_creations + 1} of {num_accounts_to_create}")
            current_account_attempts = 0
            created_successfully = False
            while not created_successfully and current_account_attempts < max_attempts_per_account:
                try:
                    await main() # Run the account creation logic
                    success_log(f"Successfully created account {successful_creations + 1}")
                    created_successfully = True
                    successful_creations += 1
                except Exception as e:
                    current_account_attempts += 1
                    error_log(f"Attempt {current_account_attempts}/{max_attempts_per_account} for account {successful_creations + 1} failed: {e}")
                    if current_account_attempts >= max_attempts_per_account:
                        error_log(f"Failed to create account {successful_creations + 1} after {max_attempts_per_account} attempts. Moving to next account or stopping if critical.")
                        break # Break from inner while loop for this specific account
                    else:
                        info_log(f"Retrying account {successful_creations + 1}...")
                        await asyncio.sleep(3) # Wait for 3 seconds before retrying

            if not created_successfully and current_account_attempts >= max_attempts_per_account:
                warning_log(f"Could not create account {successful_creations + 1}. Continuing with the next one if any.")

        if successful_creations == int(num_accounts_to_create):
            success_log(f"Successfully created all {num_accounts_to_create} accounts.")
        else:
            warning_log(f"Finished. Successfully created {successful_creations} out of {num_accounts_to_create} desired accounts after exhausting retries for some.")

    asyncio.run(run_outlook_creation(account_creation_number))

# the code will end the bot
elif question_result == "end":
    info_log("Ending the bot")
    exit()
