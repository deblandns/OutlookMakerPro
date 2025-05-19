import questionary
import asyncio
from loguru import logger
from playw import main

# the logger configs
info_log = logger.info
error_log = logger.error
debug_log = logger.debug
warning_log = logger.warning

# the code will ask the user for the starting and ending the bot
question_result = questionary.select("what do you want to do?", choices=["start", "end"]).ask()

# the code will start the bot
if question_result == "start":
    account_creation_number = questionary.text("how many accounts do you want to create?", default="1").ask()
    info_log(f"Creating {account_creation_number} accounts")
    # this code will create the accounts by the number of accounts the user wants to create
    async def run_outlook_creation(num_accounts_to_create):
        for i in range(int(num_accounts_to_create)):
            info_log(f"Creating account {i+1}")
            await main() # Run sequentially
    asyncio.run(run_outlook_creation(account_creation_number))

# the code will end the bot
elif question_result == "end":
    info_log("Ending the bot")
    exit()
