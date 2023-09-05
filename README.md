# PlannerBot WIP
## Hours Registration Bot

The Hours Registration Bot is a Discord bot designed to help users log their work hours and tasks for tracking and management. It integrates with Google Sheets to keep a record of users' daily activities.
[Hours Registration Template](https://docs.google.com/spreadsheets/d/1UK_veVeC3QWL8_2fPv0StziR0YQZhGvGAtT7ThK9Y-0/edit?usp=sharing)

## Table of Contents

- [Features](#features)
- [Usage](#usage)
- [Commands](#commands)
- [Setup](#setup)

## Features

- Self-registration of users using Discord nicknames.
- Ability to register other Discord server members for hours tracking.
- Logging of hours and finished tasks for the current date.
- Automated daily reminders for users who haven't logged their hours and tasks.

## Usage

To use the Hours Registration Bot, invite it to your Discord server and use the provided commands to register, log hours, and more. Here's how to get started:

1. Invite the bot to your server.

2. Register yourself or other members using the `/register` or `/registeruser` command.

3. Log your hours using the `/hours` or `/minutes` command.

4. Receive daily reminders if you forget to log your hours and tasks.

## Commands

- `/register`: Register yourself on the spreadsheet using your Discord nickname.
- `/registeruser <@user_option>`: Register a Discord server member for hours tracking.
- `/hours <hours> <task>`: Log hours and a description of the finished task for today.
- `/minutes <minutes> <task>`: Log minutes and a description of the finished task for today.
- `/pbhelp`: Show available commands and usage instructions.

## Setup

To set up the Hours Registration Bot for your Discord server, follow these steps:

1. Clone this repository to your local machine.

2. Create a Google service account and obtain the JSON key file. Save it in the same directory as the bot script and link to it in the script.

3. Create a `.env` file and add your Discord bot token as `DISCORD_TOKEN`.

4. Install the required Python packages using `pip`:

`pip install -r requirements.txt`

5. Run the bot script:

`python bot_script.py`

6. Invite the bot to your Discord server and grant it necessary permissions.

7. Start using the bot with the available commands.
