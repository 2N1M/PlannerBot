import os
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from interactions import Client, Intents, listen, slash_command, SlashContext, OptionType, slash_option
from interactions.api.events import MessageCreate, Startup
import datetime
import json
import pytz
import schedule
import asyncio

load_dotenv()
TOKEN = os.getenv("DISCORD_TOKEN")
loop = None

bot = Client(
    intents=Intents.DEFAULT, # intents are what events we want to receive from discord, `DEFAULT` is usually fine
    sync_interactions=True
)

# Authenticate using the service account's credentials
json_path = os.path.join(os.path.dirname(__file__), 'plannerbot-396822-76b07e4b2928.json')
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_path, scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])
gspread_client = gspread.authorize(credentials) # Authorize the credentials

def run_send_reminders():
    if datetime.datetime.today().weekday() not in [5, 6]:
        for guild_id in server_configs.keys():
            if is_reminder_channel_set(guild_id): # TODO: What if no reminder channel set?
                # Schedule the send_reminders function to run in the event loop
                loop.create_task(send_reminders_not_added_hours(guild_id))

@listen(Startup)
async def on_startup():
    global loop
    print(f'Logged in as {bot.user.display_name}')
    print(f"The bot is in {len(server_configs)} servers.")

    # Create an asyncio event loop
    loop = asyncio.get_event_loop()

    # Schedule the send_reminders function to run every day at 7:00 PM
    schedule.every().day.at("17:00").do(run_send_reminders)

    # Start the scheduler to send reminders
    while True:
        schedule.run_pending()
        await asyncio.sleep(1)  # Sleep to avoid high CPU usage

# Read configuration
def read_configs():
    try:
        with open("server_configs.json", "r") as f:
            configs = json.load(f)
            # Convert stored guild IDs to integers
            configs = {int(guild_id): data for guild_id, data in configs.items()}
            return configs
    except FileNotFoundError:
        return {}

# Initialize server_configs with the loaded data 
server_configs = read_configs()

# Update sheet URL in configuration
def update_sheet_url(guild_id, sheet_url):
    global server_configs
    server_configs = read_configs()
    server_configs[guild_id]["sheet_url"] = sheet_url
    with open("server_configs.json", "w") as f:
        json.dump(server_configs, f, indent=2)

# Update reminder channel ID in configuration
def update_reminder_channel(guild_id, reminder_channel_id):
    global server_configs
    server_configs = read_configs()
    server_configs[guild_id]["reminder_channel_id"] = reminder_channel_id
    with open("server_configs.json", "w") as f:
        json.dump(server_configs, f, indent=2)

def is_sheet_url_set(guild_id):
    configs = read_configs()
    return guild_id in configs and "sheet_url" in configs[int(guild_id)]

# Check if a reminder channel ID is set for the given guild
def is_reminder_channel_set(guild_id):
    configs = read_configs()
    return guild_id in configs and "reminder_channel_id" in configs[guild_id]

def clean_sheet_url(sheet_url):
    # Check if the link contains "?usp=sharing"
    if "?usp=sharing" in sheet_url:
        # Remove the query parameter
        sheet_url = sheet_url.split("?usp=sharing")[0]

    return sheet_url

@slash_command(name="set_sheet", description="Set the Google Sheet for the server.")
@slash_option(
    name="sheet_url",
    description="URL of the Google Sheet",
    required=True,
    opt_type=OptionType.STRING
)
async def set_sheet(ctx: SlashContext, sheet_url: str):
    # Store the sheet URL in the server_configs dictionary
    update_sheet_url(ctx.guild.id, clean_sheet_url(sheet_url))

    await ctx.send(f"The Google Sheet for this server has been set to {sheet_url}.")

@slash_command(name="hours", description="Assign hours, worked on product and task to sheet for today.")
@slash_option(
    name="hours",
    description="Hours spent working on task",
    required=True,
    opt_type=OptionType.STRING
)
@slash_option(
    name="product",
    description="Name of the product",
    required=True,
    opt_type=OptionType.STRING
)
@slash_option(
    name="task",
    description="Description of finished task",
    required=True,
    opt_type=OptionType.STRING
)
async def hours(ctx: SlashContext, hours: str, product: str, task: str):
    await ctx.defer()

    # Replace commas with dots and convert to float
    try:
        hours = float(hours.replace(',', '.'))
    except ValueError:
        await ctx.send("Invalid input for hours. Please provide a valid number.")
        return

    # Get today's date
    now = datetime.datetime.now(pytz.timezone('Europe/Amsterdam'))
    date_str = f'{now.day}/{now.month}/{now.year}'

    if not is_sheet_url_set(ctx.guild.id):
        await ctx.send("No sheet yet set for this server, set one using /set_sheet")
        return

    sheet_url = server_configs[ctx.guild.id]["sheet_url"]

    # Open the 'Hours Registration' Google Sheets document
    hours_registration = gspread_client.open_by_url(sheet_url)

    try:
        # Try to get the existing "Hours" worksheet
        user_sheet = hours_registration.worksheet("Hours")
    except gspread.exceptions.WorksheetNotFound:
        # If the worksheet doesn't exist, create a new one
        user_sheet = hours_registration.add_worksheet(title="Hours", rows="1000", cols="6")

        # Set headers in the first row
        headers = ["Date", "Discord ID", "Nickname", "Product", "Task", "Hours"]
        user_sheet.update("A1:F1", [headers])

        # Add another worksheet called "Overview"
        overview_sheet = hours_registration.add_worksheet(title="Overview", rows="8", cols="4")

        # Set headers in the first row for "Overview" worksheet
        headers_overview = ["Discord ID", "Total hours", "Name", "Added hours today"]
        overview_sheet.update("A1:D1", [headers_overview])

        # Set the formula in A2 for "Overview" worksheet
        overview_sheet.update("A2", '="=UNIQUE(Hours!B2:B)"')

        # Set the formulas in B2:D8 for "Overview" worksheet
        formulas_overview = [
            '="=IFERROR(SUMIFS(Hours!F:F, Hours!B:B, A2), "")"',
            '="=IFERROR(INDEX(Hours!C:C, MATCH(A2, Hours!B:B, 0)), "")"',
            '="=IF(ISBLANK(A2), "", IF(MAX(IF(Hours!B:B=A2, Hours!A:A))=TODAY(), \\"Yes\\", \\"No\\"))"'
        ]
        overview_sheet.update("B2:D8", [formulas_overview])

        await ctx.send(f"No \"Hours\" tab found in sheet, created a new \"Hours\" and \"Overview\" tab.")
        # Now user_sheet refers to the newly created worksheet
    else:
        # Find the first empty row in the worksheet
        row = 2
        while any(user_sheet.row_values(row)):
            row += 1

    # Update the row with the provided data
    user_sheet.update_cell(row, 1, date_str)
    user_sheet.update_cell(row, 2, str(ctx.author.id))
    user_sheet.update_cell(row, 3, ctx.author.display_name)
    user_sheet.update_cell(row, 4, product)
    user_sheet.update_cell(row, 5, task)
    user_sheet.update_cell(row, 6, hours)

    await ctx.send(f'{ctx.author.display_name} added task {task}, for product {product} to the hours registration sheet for {date_str}.')

@slash_command(name="set_reminder_channel", description="Set the reminder channel for the server.")
@slash_option(
    name="reminder_channel_id",
    description="ID of the reminder channel",
    required=True,
    opt_type=OptionType.STRING
)
async def set_reminder_channel(ctx: SlashContext, reminder_channel_id: str):
    # Store the reminder channel ID in the server_configs dictionary
    update_reminder_channel(ctx.guild.id, reminder_channel_id)

    await ctx.send(f"The reminder channel for this server has been set.")

@slash_command(name="set_reminder_channel_current", description="Set the reminder channel for the server to the current channel.")
async def set_reminder_channel_current(ctx: SlashContext):
    # Store the reminder channel ID in the server_configs dictionary
    update_reminder_channel(ctx.guild.id, ctx.channel_id)

    await ctx.send(f"The reminder channel for this server has been set.")

@slash_command(name="test_reminder", description="Test reminder messages.")
async def test_reminder(ctx: SlashContext):
    await send_reminders_not_added_hours(ctx.guild.id)

async def send_reminders_not_added_hours(guild_id):
    sheet_url = server_configs[guild_id]["sheet_url"]
    hours_registration = gspread_client.open_by_url(sheet_url)

    # Get the "Overview" worksheet
    overview_sheet = hours_registration.worksheet('Overview')

    # Get the data range
    data_range = overview_sheet.get_all_values()

    # Find the column index for the "Added hours today" column
    added_hours_col_idx = data_range[0].index("Added hours today") + 1

    # Iterate through the rows starting from the second row
    for row_idx in range(2, len(data_range) + 1):
        discord_id = data_range[row_idx - 1][0]  # Discord ID is in the first column
        added_hours_status = data_range[row_idx - 1][added_hours_col_idx - 1]

        # Check if the user has not added hours today
        if added_hours_status.lower() == "no":
            # Retrieve the Discord member
            guild = bot.get_guild(guild_id)
            member = await guild.fetch_member(int(discord_id)) # TODO: Find a better way to do this so I don't get rate limited.

            # Check if the member is found
            if member:                
                channel = bot.get_channel(server_configs[guild_id]["reminder_channel_id"])
                await channel.send(f"{member.mention} Don't forget to log your hours for today!")
            else:
                print(f'Member not found.')

bot.start(TOKEN)