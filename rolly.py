#!/bin/python3
import os
import re
import math
import discord
import httplib2
import pickle
from argparse import ArgumentParser
from configparser import ConfigParser
from threading import Thread, Event, RLock
from datetime import datetime, timedelta
from discord import Client
from googleapiclient.discovery import build
from oauth2client.client import OAuth2WebServerFlow, OAuth2Credentials

# Define constants
DISCORD_PREFIX = "[Discord] "
GOOGLE_PREFIX = "[Google] "

# Define globals
config: ConfigParser = None
config_path: str = "rolly.conf"
google_id: str = None
google_secret: str = None
google_redirect: str = None
google_sheet_id: str = None
google_sheet_ranges: str = None
google_credentials: OAuth2Credentials = None
credentials_file: str = "credentials.pkl"
discord_id: str = None
discord_bot_token: str = None
discord_bot_server: str = None
discord_bot_owners: list[str] = []
discord_roll_call_channel: str = None
sheets = None
rolly_discord: Client = discord.Client()
reaction_colours: dict[str, str] = {}

# Protect concurrency of sheets_queued_changes, sheets_next_request_after, and sheets_retries
sheets_queue_lock: RLock = RLock()
sheets_queued_changes: list[any] = []
sheets_next_request_after: datetime = datetime.now()
sheets_retries: int = 0


# Define some classes
class RepeatingTimer(Thread):
    def __init__(self, interval_seconds, callback):
        super().__init__()
        self.stop_event = Event()
        self.interval_seconds = interval_seconds
        self.callback = callback

    def run(self):
        while not self.stop_event.wait(self.interval_seconds):
            self.callback()

    def stop(self):
        self.stop_event.set()


# Define some functions
def init_from_args():
    """
    Parses command-line arguments and applies them to the corresponding globals.
    """
    global config_path, credentials_file

    # Configure parser
    parser = ArgumentParser()
    parser.add_argument(
        "-c",
        "--config",
        metavar="PATH",
        help="Path to configuration file (default: rolly.py)",
    )
    parser.add_argument(
        "-d",
        "--credentials",
        metavar="PATH",
        help="Path to credential cache (default: credentials.pkl)",
    )
    args = parser.parse_args()

    # Parse provided arguments
    if args.config:
        config_path = args.config
    if args.credentials:
        credentials_file = args.credentials


def init_from_config():
    """
    Reads in config and applies it to the corresponding globals.
    """
    global config, google_id, google_secret, google_redirect, google_sheet_id, google_sheet_ranges, discord_id, discord_bot_token, discord_bot_server, discord_bot_owners, discord_roll_call_channel, reaction_colours

    # Read in our ID and secret from config
    config = ConfigParser()
    config.read(config_path)

    if not config.sections():
        print("No existing config was found, creating default...")
        with open(config_path, "w") as file:
            file.write(
                "[Google]\n"
                + "client_id = \n"
                + "client_secret = \n"
                + "redirect_url = \n"
                + "sheet_id = \n"
                + "sheet_ranges = \n"
                + "\n"
                + "[Discord]\n"
                + "client_id = \n"
                + "client_secret = \n"
                + "bot_token = \n"
                + "bot_owners = \n"
                + "bot_server = \n"
                + "reaction_colours = \n"
                + "roll_call_channel = \n"
            )
        exit(1)

    # Check that all mandatory config options are present
    mandatory_sections = {
        "Google": [
            "client_id",
            "client_secret",
            "redirect_url",
            "sheet_id",
            "sheet_ranges",
        ],
        "Discord": ["client_id", "bot_token", "bot_server"],
    }
    for section, keys in mandatory_sections.items():
        if section not in config:
            print("Failed to read config: '%s' section missing" % section)
            exit(1)
        else:
            for key in keys:
                if key not in config[section]:
                    print(
                        "Failed to read config: '%s' missing from section '%s'"
                        % (key, section)
                    )
                    exit(1)

    # Apply config
    google_id = config["Google"]["client_id"]
    google_secret = config["Google"]["client_secret"]
    google_redirect = config["Google"]["redirect_url"]
    google_sheet_id = config["Google"]["sheet_id"]
    google_sheet_ranges = config["Google"]["sheet_ranges"]

    discord_id = config["Discord"]["client_id"]
    discord_bot_token = config["Discord"]["bot_token"]
    discord_bot_server = config["Discord"]["bot_server"]
    try:
        discord_bot_owners = config["Discord"]["bot_owners"].split()
    except KeyError:
        print("Couldn't read bot owners, defaulting to none")
        discord_bot_owners = []
    try:
        for pair in config["Discord"]["reaction_colours"].split():
            reaction, colour = pair.split(":")
            reaction_colours[reaction] = colour
    except KeyError:
        print(
            "Couldn't read reaction/colour mapping, defaulting to %s" % reaction_colours
        )
        reaction_colours = {"✅": "00ff00", "❔": "ffff00", "❌": "ff0000"}
    try:
        discord_roll_call_channel = config["Discord"]["roll_call_channel"]
    except KeyError:
        print("Couldn't read roll call channel, defaulting to none")


def google_refresh_tokens():
    """
    Refreshes Google API tokens and saves them to disk.
    """
    if google_credentials and not google_credentials.invalid:
        google_credentials.refresh(httplib2.Http())
        print(GOOGLE_PREFIX + "Refreshed tokens")
        print(
            GOOGLE_PREFIX
            + "New token expires in "
            + str(google_credentials.token_expiry - datetime.now())
        )

        # Pickle the credentials object
        with open(credentials_file, "wb") as file:
            pickle.dump(google_credentials, file)

        return True
    else:
        print(GOOGLE_PREFIX + "Failed to refresh tokens, credentials are invalid now")
        return False


def google_load_credentials():
    """
    Loads Google API credentials from disk.
    """
    global google_credentials, sheets

    if os.path.exists(credentials_file):
        with open(credentials_file, "rb") as file:
            google_credentials = pickle.load(file)

    if not google_refresh_tokens():
        flow = OAuth2WebServerFlow(
            client_id=google_id,
            client_secret=google_secret,
            scope="https://www.googleapis.com/auth/spreadsheets",
            redirect_uri=google_redirect,
            prompt="consent",
        )

        google_auth_uri = flow.step1_get_authorize_url()

        print("Please authorise yourself at the below URL and paste the code here")
        print(google_auth_uri)
        google_auth_code = input("Code: ")
        google_credentials = flow.step2_exchange(google_auth_code)

        # Pickle the credentials object
        with open(credentials_file, "wb") as file:
            pickle.dump(google_credentials, file)

    # Set up the Google Sheets service
    sheets_service = build("sheets", "v4", credentials=google_credentials)
    sheets = sheets_service.spreadsheets()


async def setup(channel, message_content=""):
    """
    Creates a roll call message in the given channel
    :param channel: Channel to send a message to
    :param message_content: Optional message content
    """
    message = await channel.send(message_content if message_content else "Roll call!")
    for key, _ in reaction_colours.items():
        await message.add_reaction(key)


def parse_a1_coords(a1):
    """
    Parses an A1 string and returns the coordinates as integers
    Does not accept ranges
    :param a1: A1 string to parse
    :return: x and y coordinates as a list
    """
    if ":" in a1:
        raise ValueError("Cannot process an A1 range.")

    match = re.match(r"(?:\w+!)?(([A-Z]+)([0-9]+))", a1, re.I)
    if match:
        items = match.groups()
        if len(items) == 3:
            # return items[2:3]
            def col_name_to_num(cn):
                return sum(
                    [((ord(cn[-1 - pos]) - 64) * 26**pos) for pos in range(len(cn))]
                )

            return col_name_to_num(items[1]) - 1, int(items[2]) - 1
        else:
            raise ValueError("Got more parts than expected.")
    else:
        raise ValueError(
            "Unable to process A1 string properly, are you sure it is valid?"
        )


def sheet_update_user(name, colour_hex):
    """
    Updates the background colour of a user in Google Sheets
    :param name: Name of the user
    :param colour_hex: Colour to set
    """
    locked = False
    try:
        # Gain exclusive access to the Sheets queue
        if sheets_queue_lock.acquire(blocking=True):
            locked = True
        else:
            raise Exception("Unable to acquire lock on Sheets queue")

        # Queue the change to be commited in the next batch
        sheets_queued_changes.append({"name": name, "colour": colour_hex})
    except Exception as e:
        print(e)
    finally:
        if locked:
            sheets_queue_lock.release()


def sheets_commit_changes():
    """
    Commits any outstanding changes in sheets_queue_lock to Google Sheets.
    """
    # Don't do anything if we're called too early or there's nothing queued
    if datetime.now() <= sheets_next_request_after or not sheets_queued_changes:
        return

    try:
        # Try match each of the requested changes to cells within the ranges we are permitted to operate upon
        sheet_response = (
            sheets.values()
            .batchGet(spreadsheetId=google_sheet_id, ranges=google_sheet_ranges.split())
            .execute()
        )
        value_ranges = sheet_response.get("valueRanges", [])
    except Exception as e:
        # Delay next execution exponentially up to 32s
        sheets_increment_retry_delay()
        print(
            GOOGLE_PREFIX
            + "Failed to get sheet data, will try again in {}. Original exception: {}".format(
                sheets_next_request_after - datetime.now(), str(e)
            )
        )
        return

    requests = []
    locked = False
    try:
        # Gain exclusive access to the Sheets queue
        if not sheets_queue_lock.acquire(blocking=True):
            raise Exception("Unable to acquire lock on Sheets queue")
        else:
            locked = True

            for value_range in value_ranges:
                # Make sure we got something
                if not value_range["values"]:
                    print(
                        GOOGLE_PREFIX
                        + "No data found for range {}".format(value_range["range"])
                    )
                    continue

                # Get the x and y origin offset for this range
                x_offset, y_offset = parse_a1_coords(value_range["range"].split(":")[0])

                # Iterate over each of the changes to be made
                for pair in sheets_queued_changes:
                    # Try find the target name within the range
                    for y, row in enumerate(value_range["values"]):
                        for x, value in enumerate(row):
                            if value and value.lower() in pair["name"].lower():
                                # Find the boundary of this cell
                                column_start = x_offset + x
                                column_end = x_offset + x + 1
                                row_start = y_offset + y
                                row_end = y_offset + y + 1

                                # Convert the hex string to RGB values
                                red, green, blue = bytes.fromhex(pair["colour"])

                                # Make a request to change this cell's background colour to the one requested
                                requests.append(
                                    {
                                        "repeatCell": {
                                            "range": {
                                                "startColumnIndex": column_start,
                                                "endColumnIndex": column_end,
                                                "startRowIndex": row_start,
                                                "endRowIndex": row_end,
                                                "sheetId": 0,
                                            },
                                            "cell": {
                                                "userEnteredFormat": {
                                                    "backgroundColor": {
                                                        "red": red / 255,
                                                        "green": green / 255,
                                                        "blue": blue / 255,
                                                    }
                                                }
                                            },
                                            "fields": "UserEnteredFormat(BackgroundColor)",
                                        }
                                    }
                                )

            # Clear the sheets queue
            sheets_queued_changes.clear()
    except Exception as e:
        print(GOOGLE_PREFIX + "Unexpected error while searching sheet: " + str(e))
        return
    finally:
        if locked:
            # Release queue lock
            sheets_queue_lock.release()

    if requests:
        # Compile all pending changes into a request and send it off to the Sheets API
        try:
            sheets.batchUpdate(
                spreadsheetId=google_sheet_id, body={"requests": requests}
            ).execute()

            print(
                GOOGLE_PREFIX
                + "Committed {} change{} to Sheets".format(
                    len(requests), "s" if len(requests) != 1 else ""
                )
            )
        except Exception as e:
            sheets_increment_retry_delay()
            print(
                GOOGLE_PREFIX
                + "Failed to commit sheets changes, will try again in {}. Original exception: {}".format(
                    sheets_next_request_after - datetime.now(), str(e)
                )
            )
            return

    # Reset the retry delay
    sheets_clear_retry_delay()


def sheets_increment_retry_delay():
    """
    Increments sheets_retries and sheets_next_request_after according to Google's recommended exponential backoff algorithm.
    """
    global sheets_retries, sheets_next_request_after

    # Increment retries
    sheets_retries = sheets_retries + 1

    # Calculate delay as 2^sheets_retries up to 32s
    delay = math.pow(2, min(sheets_retries, 5))
    sheets_next_request_after = datetime.now() + timedelta(seconds=delay)


def sheets_clear_retry_delay():
    """
    Clears sheets_retries and sheets_next_request_after.
    """
    global sheets_retries, sheets_next_request_after

    sheets_retries = 0
    sheets_next_request_after = datetime.now()


def google_token_timer_refresh():
    if not google_refresh_tokens():
        print(GOOGLE_PREFIX + "Stopping token timer since we cannot refresh anymore...")
        google_token_timer.stop()


@rolly_discord.event
async def on_ready():
    print(DISCORD_PREFIX + "Bot logged in!")


@rolly_discord.event
async def on_message(message):
    # Declare this globally here, since we use it early on, /and/ in a command
    # global discord_bot_channel

    # Make sure this is from the desired server
    if str(message.guild.id) != discord_bot_server:
        # print(DISCORD_PREFIX + 'Got a message from server {} channel {}, expected server {}'.format(message.guild.id, message.channel.id, discord_bot_server))
        return

    # If the bot owner is set, make sure this is from them
    if discord_bot_owners and str(message.author.id) not in discord_bot_owners:
        # print(DISCORD_PREFIX + 'Got a message from user {}, expected one of {}'.format(message.author.id, discord_bot_owners))
        return

    # Determine what prefix was used to address us, if any
    if message.content.startswith("#rolly"):
        prefix = "#rolly"
    elif message.content.startswith("<@" + str(discord_id) + ">"):
        prefix = "@{}#{}".format(
            rolly_discord.user.name, rolly_discord.user.discriminator
        )
    else:
        # We weren't addressed, we can stop here
        return

    # Split the command into arguments
    args = message.content.strip().split()[1:]

    # Return a message if no args are provided
    if not args:
        await message.channel.send(
            "Yo, I'm Rolly. Try `{} create` to start a roll call.".format(prefix),
            delete_after=30,
        )
    else:
        # Declare some globals
        global google_sheet_ranges
        global google_sheet_id

        # Filter what command came through
        if args[0] == "help":
            await message.channel.send(
                "Here's a list of commands you can give me:\n"
                + "\n"
                + "`help` - Shows this help text\n"
                + "`create` - Creates a new roll call message\n"
                + "`setsheet` - Sets the Google Sheet ID to update\n"
                + "`ranges` - Shows the current allowed ranges in the spreadsheet\n"
                + "`addranges` - Add one or more allowed ranges for the spreadsheet\n"
                + "`setranges` - Sets the ranges to update in the spreadsheet"
            )

        elif args[0] == "create":
            if len(args) > 1:
                await setup(message.channel, " ".join(args[1:]))
            else:
                await setup(message.channel)

        elif args[0] == "setsheet":
            # Make sure we've been given an ID
            if len(args) < 2:
                await message.channel.send(
                    "I need a sheet ID to do that.\nYou can get it from the URL e.g. `https://docs.google.com/spreadsheets/d/<sheet id>/`",
                    delete_after=30,
                )
            else:
                # Update and write config to file
                config["Google"]["sheet_id"] = args[1]
                google_sheet_id = args[1]
                with open(config_path, "w") as file:
                    config.write(file)

        elif args[0] == "ranges":
            # List off the current ranges
            message_str = (
                "Current allowed ranges are: `" + config["Google"]["sheet_ranges"] + "`"
            )

            await message.channel.send(message_str)

        elif args[0] == "addranges":
            # Make sure we've been given a range
            if len(args) < 2:
                await message.channel.send(
                    "I need at least one range to do that.\nBy range I mean something like `C2`, `A3:B7`, or `Sheet2!H13:AC139`. You can include several, just separate them with a space.",
                    delete_after=30,
                )
            else:
                # Update and write config to file
                new_ranges = config["Google"]["sheet_ranges"] + " " + " ".join(args[1:])
                config["Google"]["sheet_ranges"] = new_ranges
                google_sheet_ranges = new_ranges
                with open(config_path, "w") as file:
                    config.write(file)

        elif args[0] == "addranges":
            # Make sure we've been given a range
            if len(args) < 2:
                await message.channel.send(
                    "I need at least one range to do that.\nBy range I mean something like `C2`, `A3:B7`, or `Sheet2!H13:AC139`. You can include several, just separate them with a space.",
                    delete_after=30,
                )
            else:
                # Update and write config to file
                new_ranges = config["Google"]["sheet_ranges"] + " " + " ".join(args[1:])
                config["Google"]["sheet_ranges"] = new_ranges
                google_sheet_ranges = new_ranges
                with open(config_path, "w") as file:
                    config.write(file)

        elif args[0] == "setranges":
            # Make sure we've been given a range
            if len(args) < 2:
                await message.channel.send(
                    "I need at least one range to do that.\nBy range I mean something like `C2`, `A3:B7`, or `Sheet2!H13:AC139`. You can include several, just separate them with a space.",
                    delete_after=30,
                )
            else:
                # Update and write config to file
                config["Google"]["sheet_ranges"] = " ".join(args[1:])
                google_sheet_ranges = " ".join(args[1:])
                with open(config_path, "w") as file:
                    config.write(file)

    # Delete the command
    await message.delete()


@rolly_discord.event
async def on_raw_reaction_add(event):
    # Grab the channel and message that was reacted on
    channel = rolly_discord.get_channel(event.channel_id)
    message = await channel.fetch_message(event.message_id)

    # Make sure this is from the desired server
    if str(message.guild.id) != discord_bot_server:
        return

    # Check if it was on a message we sent
    if str(message.author.id) == discord_id:
        # Grab the user that reacted and the emoji
        user = await channel.guild.fetch_member(event.user_id)
        emoji = event.emoji.name

        # Ignore reacts that we make
        if str(user.id) == discord_id:
            return

        # Report what's happened
        global reaction_colours
        if emoji in reaction_colours:
            print(
                DISCORD_PREFIX
                + "{} reacted with '{}', changing their cell to {}".format(
                    user.display_name, emoji, reaction_colours[emoji]
                )
            )
        else:
            print(
                DISCORD_PREFIX
                + "{} reacted with unsupported emoji '{}'".format(
                    user.display_name, emoji
                )
            )
            return

        sheet_update_user(user.display_name, reaction_colours[emoji])


@rolly_discord.event
async def on_raw_reaction_remove(event):
    # Grab the channel and message that was reacted on
    channel = rolly_discord.get_channel(event.channel_id)
    message = await channel.fetch_message(event.message_id)

    # Make sure this is from the desired server
    if str(message.guild.id) != discord_bot_server:
        return

    # Check if it was on a message we sent
    if str(message.author.id) == discord_id:
        # Grab the user that reacted and the emoji
        user = await channel.guild.fetch_member(event.user_id)
        emoji = event.emoji.name

        # Ignore reacts that we make
        if str(user.id) == discord_id:
            return

        # Search through the other reacts
        global reaction_colours
        for reaction in message.reactions:
            # Skip the removed emoji and emojis we can't process
            if reaction.emoji == emoji or reaction.emoji not in reaction_colours.keys():
                continue

            # Iterate over each user that has reacted
            for react_user in await reaction.users().flatten():
                if react_user.id == user.id:
                    # Use this emoji instead
                    print(
                        DISCORD_PREFIX
                        + "{} removed their '{}' react, but they still have a '{}' react. Changing their cell to {}".format(
                            user.display_name,
                            emoji,
                            reaction.emoji,
                            reaction_colours[reaction.emoji],
                        )
                    )
                    sheet_update_user(
                        user.display_name, reaction_colours[reaction.emoji]
                    )
                    return

        # Didn't find an alternate emoji to use, clear their cell
        print(
            DISCORD_PREFIX
            + "{} removed their '{}' react, changing their cell to {}".format(
                user.display_name, emoji, "ffffff"
            )
        )
        sheet_update_user(user.display_name, "ffffff")


if __name__ == "__main__":
    # Read in args and config then apply accordingly
    init_from_args()
    init_from_config()

    # Load in Google credentials and/or prompt for OAuth authorisation
    google_load_credentials()

    # Start a timer to keep our Google tokens refreshed
    google_token_timer = RepeatingTimer(600, google_token_timer_refresh)
    google_token_timer.start()

    # Start a timer to flush through changes to Sheets
    sheets_commit_changes_timer = RepeatingTimer(3, sheets_commit_changes)
    sheets_commit_changes_timer.start()

    # Start the Discord bot
    print(DISCORD_PREFIX + "Starting Rolly")
    print(DISCORD_PREFIX + "Owner IDs: ")
    for id in discord_bot_owners:
        print(DISCORD_PREFIX + "    {}".format(id))
    rolly_discord.run(discord_bot_token)

    # Stop the token refresh timer
    google_token_timer.stop()
