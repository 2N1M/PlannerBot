import os
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from interactions import Client, Intents, listen, slash_command, SlashContext, OptionType, slash_option, User, to_snowflake
import datetime
import json
import pytz
import schedule
import asyncio

load_dotenv()
TOKEN = os.getenv("DISCORD_TOKEN")
loop = None

# intents are what events we want to receive from discord, `DEFAULT` is usually fine
bot = Client(
    #debug_scope= to_snowflake(1143957256145223740), 
    intents=Intents.DEFAULT,
    sync_interactions=True
)

# Authenticate using the service account's credentials
json_path = os.path.join(os.path.dirname(__file__), 'plannerbot-396822-76b07e4b2928.json')
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_path, scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])
# Authorize the credentials
gspread_client = gspread.authorize(credentials)

async def send_reminders():
    # Get the current date in the format "day/month/year"
    now = datetime.datetime.now(pytz.timezone('Europe/Amsterdam'))
    date_str = f'{now.day}/{now.month}/{now.year}'
    
    # Open the Google Sheets document
    hours_registration_doc = gspread_client.open('HoursRegistration')\
    
    # Iterate through registered users
    for user_sheet in hours_registration_doc.worksheets():
        # Check if the user's sheet exists
        if user_sheet.title != "TemplateSheet" and user_sheet.title != "Overview":

            print(user_sheet.cell(3, 1).value)

            # Find the correct column for the date
            date_range = user_sheet.row_values(2)

            try:
                col_idx = date_range.index(date_str) + 1
            except ValueError:
                await print("No matching column found for today's date.")
                return
                
            # Check if the cell is empty
            if user_sheet.cell(5, col_idx).value is None:
                user_name = user_sheet.cell(3, 1).value
                channel = bot.get_channel(1143957904215523389)

                guild = bot.get_guild(1143957256145223740)
                members = await guild.search_members(user_name)
                if members:
                    member = members[0]
                    await channel.send(f"{member.mention} Don't forget to log your hours and tasks for today!")

def run_send_reminders():
    # Schedule the send_reminders function to run in the event loop
    loop.create_task(send_reminders())

@listen()
async def on_ready():
    global loop
    print(f'Logged in as {bot.user.display_name}')

    # Create an asyncio event loop
    loop = asyncio.get_event_loop()

    # Schedule the send_reminders function to run every day at 7:13 PM
    schedule.every().day.at("17:00").do(run_send_reminders)
    
    # Start the scheduler to send reminders
    while True:
        schedule.run_pending()
        await asyncio.sleep(1)  # Sleep to avoid high CPU usage

@slash_command(name="test", description="Test command")
async def test(ctx: SlashContext):
    await ctx.send('This is a test command!')

#Self registration
@slash_command(name="register", description="Register yourself on the spreadsheet using your Discord nickname.")
async def register(ctx: SlashContext):
    await ctx.defer()
    
    # Open the 'Hours Registration' Google Sheets document
    hours_registration_doc = gspread_client.open('HoursRegistration')
    
    # Check if a sheet with the user's name already exists
    sheet_title = ctx.author.display_name
    existing_sheets = hours_registration_doc.worksheets()
    sheet_exists = any(sheet.title == sheet_title for sheet in existing_sheets)
    
    if sheet_exists:
        await ctx.send(f'Welcome back, {ctx.author.display_name}! Your sheet already exists.')
    else:
        # Duplicate the 'TemplateSheet'
        template_sheet = hours_registration_doc.worksheet('TemplateSheet')
        new_sheet = template_sheet.duplicate(new_sheet_name=sheet_title)
        
        # Add user's name to the team member cell
        new_sheet.update('A3', ctx.author.display_name)

        overview_sheet = hours_registration_doc.worksheet('Overview')
        
        # Check if the cell is empty
        if overview_sheet.cell(4, 1).value is None:
            # Set team member name
            overview_sheet.update_cell(4, 1, ctx.author.display_name)

            await ctx.send(f'Welcome, {ctx.author.display_name}! Your sheet has been created.')
        else:
            # Check the next row if the cell is occupied
            row = 5 
            while overview_sheet.cell(row, 1).value is not None:
                # The cell is not empty, move to the next row and check again
                row += 1                    
            
            # The cell is truly empty, set the team member name
            overview_sheet.update_cell(row, 1, ctx.author.display_name)

            await ctx.send(f'Welcome, {ctx.author.display_name}! Your sheet has been created.')

#Discord server member registration
@slash_command(name="registeruser", description="Register member on the spreadsheet.")
@slash_option(
    name="user_option",
    description="User Option",
    required=True,
    opt_type=OptionType.USER
)
async def registeruser(ctx: SlashContext, user_option: User):
    await ctx.defer()
    
    # Open the 'Hours Registration' Google Sheets document
    hours_registration_doc = gspread_client.open('HoursRegistration')
    
    # Check if a sheet with the user's name already exists
    sheet_title = user_option.display_name
    existing_sheets = hours_registration_doc.worksheets()
    sheet_exists = any(sheet.title == sheet_title for sheet in existing_sheets)

    if sheet_exists:
        await ctx.send(f'{user_option.display_name}\'s sheet already exists.')
    else:
        # Duplicate the 'TemplateSheet'
        template_sheet = hours_registration_doc.worksheet('TemplateSheet')
        new_sheet = template_sheet.duplicate(new_sheet_name=sheet_title)
        
        # Add user's name to the team member cell
        new_sheet.update('A3', user_option.display_name)
        
        overview_sheet = hours_registration_doc.worksheet('Overview')
        
        # Check if the cell is empty
        if overview_sheet.cell(4, 1).value is None:
            # Set team member name
            overview_sheet.update_cell(4, 1, user_option.display_name)

            await ctx.send(f'{user_option.display_name}\'s sheet has been created.')
        else:
            # Check the next row if the cell is occupied
            row = 5 
            while overview_sheet.cell(row, 1).value is not None:
                # The cell is not empty, move to the next row and check again
                row += 1        
            # The cell is truly empty, set the team member name
            overview_sheet.update_cell(row, 1, user_option.display_name)

            await ctx.send(f'{user_option.display_name}\'s sheet has been created.')

@slash_command(name="hours", description="Assign hours and finished task to sheet for today.")
@slash_option(
    name="hours",
    description="Hours spent working on task",
    required=True,
    opt_type=OptionType.INTEGER
)
@slash_option(
    name="task",
    description="Description of finished task",
    required=True,
    opt_type=OptionType.STRING
)
async def hours(ctx: SlashContext, hours: int, task: str):
    await ctx.defer()

    # Get today's date
    now = datetime.datetime.now(pytz.timezone('Europe/Amsterdam'))
    date_str = f'{now.day}/{now.month}/{now.year}'
    
    # Open the 'Hours Registration' Google Sheets document
    hours_registration_doc = gspread_client.open('HoursRegistration')

    # Check if the user's sheet exists
    try:
        user_sheet = hours_registration_doc.worksheet(ctx.author.display_name)
    except gspread.exceptions.WorksheetNotFound:
        await ctx.send(f"You are not registered or your sheet has not been created. Use '/register' to register.")
        return

    # Find the correct column for the date
    date_range = user_sheet.row_values(2)
    try:
        col_idx = date_range.index(date_str) + 1
    except ValueError:
        await ctx.send("No matching column found for today's date.")
        return
    
    # Check if the cell is empty
    if user_sheet.cell(5, col_idx).value is None:
        # Update hours and task
        user_sheet.update_cell(5, col_idx, hours)
        user_sheet.update_cell(5, col_idx+1, task)
        await ctx.send(f'{hours} hour(s) spent on {task} added to hours registration for {date_str} by {ctx.author.display_name}.')
    else:
        # Check the next row if the cell is occupied
        row = 6 
        while user_sheet.cell(row, col_idx).value is not None:
            # The cell is not empty, move to the next row and check again
            #print(f"Value at row {row}: '{user_sheet.cell(row, col_idx).value}' (Type: {type(user_sheet.cell(row, col_idx).value)})")
            row += 1        
        # The cell is truly empty, update the hours and task
        user_sheet.update_cell(row, col_idx, hours)
        user_sheet.update_cell(row, col_idx+1, task)
        await ctx.send(f'{hours} hour(s) spent on {task} added to hours registration for {date_str} by {ctx.author.display_name}.')

@slash_command(name="minutes", description="Assign minutes and finished task to sheet for today.")
@slash_option(
    name="minutes",
    description="Minutes spent working on task",
    required=True,
    opt_type=OptionType.INTEGER
)
@slash_option(
    name="task",
    description="Description of finished task",
    required=True,
    opt_type=OptionType.STRING
)
async def minutes(ctx: SlashContext, minutes: int, task: str):
    await ctx.defer()
    
    # Get today's date
    now = datetime.datetime.now(pytz.timezone('Europe/Amsterdam'))
    date_str = f'{now.day}/{now.month}/{now.year}'
    
    # Open the 'Hours Registration' Google Sheets document
    hours_registration_doc = gspread_client.open('HoursRegistration')

    # Check if the user_option's sheet exists
    try:
        user_sheet = hours_registration_doc.worksheet(ctx.author.display_name)
    except gspread.exceptions.WorksheetNotFound:
        await ctx.send(f"You are not registered or your sheet has not been created. Use '/register' to register.")
        return

    # Find the correct column for the date
    date_range = user_sheet.row_values(2)
    try:
        col_idx = date_range.index(date_str) + 1
    except ValueError:
        await ctx.send("No matching column found for today's date.")
        return
    
    # Check if the cell is empty
    if user_sheet.cell(5, col_idx).value is None:
        # Update minutes and task
        user_sheet.update_cell(5, col_idx, minutes/60)
        user_sheet.update_cell(5, col_idx+1, task)
        await ctx.send(f'{minutes} minute(s) spent on {task} added to hours registration for {date_str} by {ctx.author.display_name}.')
    else:
        # Check the next row if the cell is occupied
        row = 6 
        while user_sheet.cell(row, col_idx).value is not None:
            # The cell is not empty, move to the next row and check again
            #print(f"Value at row {row}: '{user_sheet.cell(row, col_idx).value}' (Type: {type(user_sheet.cell(row, col_idx).value)})")
            row += 1
        
        # The cell is truly empty, update the hours and task
        user_sheet.update_cell(row, col_idx, minutes/60)
        user_sheet.update_cell(row, col_idx+1, task)
        await ctx.send(f'{minutes} minute(s) spent on {task} added to hours registration for {date_str} by {ctx.author.display_name}.')

@slash_command(name="pbhelp", description="Shows available commands")
async def pbhelp(ctx: SlashContext):
    help_message = (
        "Welcome to the Hours Registration Bot!\n\n"
        "Available commands:\n"
        "/pbhelp - Show this help message\n"
        "/register - Register yourself on the spreadsheet using your Discord nickname\n"
        "/registeruser `<@user_option>` - Register a Discord server member for hours tracking\n"
        "/hours `<hours>`, `<task>` - Log hours and task for today\n"
        "/minutes `<minutes>`, `<task>` - Log minutes and task for today\n"
    )
    await ctx.send(help_message)

bot.start(TOKEN)
