#!/bin/python3
import datetime
import os
import re
import discord
import httplib2
import pickle
from argparse import ArgumentParser
from configparser import ConfigParser
from threading import Thread, Event
from googleapiclient.discovery import build
from oauth2client.client import OAuth2WebServerFlow

##### Define some constants ############################################################################################

DISCORD_PREFIX = '[Discord] '
GOOGLE_PREFIX = '[Google] '
CREDENTIALS_FILE = 'credentials.pkl'

##### Read in our ID and secret from config ############################################################################

config = ConfigParser()
config_path = 'rolly.conf'
config.read(config_path)

if not config.sections():
    print('No existing config was found')
    print('Copy the following blank template into ' + config_path + ' and fill in the blanks:')
    print('[Google]\n' +
          'client_id = \n' +
          'client_secret = \n' +
          'redirect_url = \n' +
          'sheet_id = \n' +
          'sheet_ranges = \n'
          '\n' +
          '[Discord]\n' +
          'client_id = \n' +
          'client_secret = \n' +
          'bot_token = \n' +
          'bot_owners = []\n' +
          'bot_server = \n')
    exit(1)

if 'Google' not in config:
    print('Failed to read config: \'Google\' section missing')
    exit(1)
if 'client_id' not in config['Google']:
    print('Failed to read config: \'client_id\' missing from section \'Google\'')
    exit(1)
if 'client_secret' not in config['Google']:
    print('Failed to read config: \'client_secret\' missing from section \'Google\'')
    exit(1)
if 'redirect_url' not in config['Google']:
    print('Failed to read config: \'redirect_url\' missing from section \'Google\'')
    exit(1)
if 'sheet_id' not in config['Google']:
    print('Failed to read config: \'sheet_id\' missing from section \'Google\'')
    exit(1)
if 'sheet_ranges' not in config['Google']:
    print('Failed to read config: \'sheet_ranges\' missing from section \'Google\'')
    exit(1)

if 'Discord' not in config:
    print('Failed to read config: \'Discord\' section missing')
    exit(1)
if 'client_id' not in config['Discord']:
    print('Failed to read config: \'client_id\' missing from section \'Discord\'')
    exit(1)
if 'bot_token' not in config['Discord']:
    print('Failed to read config: \'bot_token\' missing from section \'Discord\'')
    exit(1)
# if 'bot_owner' not in config['Discord']:
#     print('Failed to read config: \'bot_owner\' missing from section \'Discord\'')
#     exit(1)
if 'bot_server' not in config['Discord']:
    print('Failed to read config: \'bot_server\' missing from section \'Discord\'')
    exit(1)

google_id = config['Google']['client_id']
google_secret = config['Google']['client_secret']
google_redirect = config['Google']['redirect_url']
google_sheet_id = config['Google']['sheet_id']
google_sheet_ranges = config['Google']['sheet_ranges']

discord_id = config['Discord']['client_id']
discord_bot_token = config['Discord']['bot_token']
discord_bot_server = config['Discord']['bot_server']
try:
    discord_bot_owners = config['Discord']['bot_owners'].split()
except KeyError:
    discord_bot_owners = []

##### Parse arguments ##################################################################################################

parser = ArgumentParser()
args = parser.parse_args()

##### Authorise with Google ############################################################################################

google_credentials = None
def google_refresh_tokens():
    if google_credentials and not google_credentials.invalid:
        google_credentials.refresh(httplib2.Http())
        print(GOOGLE_PREFIX + 'Refreshed tokens')
        print(GOOGLE_PREFIX + 'New token expires in ' + str(datetime.datetime.now() - google_credentials.token_expiry))

        # Pickle the credentials object
        with open(CREDENTIALS_FILE, 'wb') as file:
            pickle.dump(google_credentials, file)

        return True
    else:
        print(GOOGLE_PREFIX + 'Failed to refresh tokens, credentials are invalid now')
        return False


if os.path.exists(CREDENTIALS_FILE):
    with open(CREDENTIALS_FILE, 'rb') as file:
        google_credentials = pickle.load(file)

if not google_refresh_tokens():
    flow = OAuth2WebServerFlow(client_id=google_id,
                               client_secret=google_secret,
                               scope='https://www.googleapis.com/auth/spreadsheets',
                               redirect_uri=google_redirect,
                               prompt='consent')

    google_auth_uri = flow.step1_get_authorize_url()

    print("Please authorise yourself at the below URL and paste the code here")
    print(google_auth_uri)
    google_auth_code = input('Code: ')
    google_credentials = flow.step2_exchange(google_auth_code)

    # Pickle the credentials object
    with open(CREDENTIALS_FILE, 'wb') as file:
        pickle.dump(google_credentials, file)

##### Set up the Google Sheets service #################################################################################

sheets_service = build('sheets', 'v4', credentials=google_credentials)
sheets = sheets_service.spreadsheets()

##### Define some classes ##############################################################################################

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

##### Define some functions ############################################################################################

async def setup(channel, message_content = ''):
    """
    Creates a roll call message in the given channel
    :param channel: Channel to send a message to
    :param message_content: Optional message content
    """
    message = await channel.send(message_content if message_content else 'Roll call!')
    await message.add_reaction('✅')
    await message.add_reaction('❔')
    await message.add_reaction('❌')

def contains_other(str_a, str_b):
    """
    Compares two strings to see if they contain each other, ignoring case
    :param str_a: String to compare
    :param str_b: String to compare
    :return: True if one contains the other, otherwise false
    """
    return (str_a.strip().lower() in str_b.strip().lower()) or (str_b.strip().lower() in str_a.strip().lower())

def parse_a1_coords(a1):
    """
    Parses an A1 string and returns the coordinates as integers
    Does not accept ranges
    :param a1: A1 string to parse
    :return: x and y coordinates as a list
    """
    if ':' in a1:
        raise ValueError('Cannot process an A1 range.')

    match = re.match(r"(?:\w+!)?(([A-Z]+)([0-9]+))", a1, re.I)
    if match:
        items = match.groups()
        if len(items) == 3:
            # return items[2:3]
            colNameToNum = lambda cn: sum([((ord(cn[-1 - pos]) - 64) * 26 ** pos) for pos in range(len(cn))])
            return colNameToNum(items[1]) - 1, int(items[2]) - 1
        else:
            raise ValueError('Got more parts than expected.')
    else:
        raise ValueError('Unable to process A1 string properly, are you sure it is valid?')


def sheet_update_user(name, colour_hex):
    """
    Updates the background colour of a user in Google Sheets
    :param name: Name of the user
    :param colour_hex: Colour to set
    """
    column_start = 0
    column_end = 0
    row_start = 0
    row_end = 0

    # Iterate over each range and find a cell that corresponds to name
    found = False
    for range in google_sheet_ranges.split():
        # Query our range from Sheets
        sheet_response = sheets.values().get(spreadsheetId=google_sheet_id, range=range).execute()
        sheet_values = sheet_response.get('values', [])

        # Make sure we got something
        if not sheet_values:
            print(GOOGLE_PREFIX + 'No data found')
            return

        # Get the x and y origin offset for this range
        x_offset, y_offset = parse_a1_coords(range.split(':')[0])

        # Try find the target name within the range
        for y, row in enumerate(sheet_values):
            for x, value in enumerate(row):
                if value and value.lower() in name.lower():
                    column_start = x_offset + x
                    column_end = x_offset + x + 1
                    row_start = y_offset + y
                    row_end = y_offset + y + 1
                    found = True
                    break

        # Stop iterating ranges if we found what we're looking for
        if found:
            break

    # Make sure we got a value
    if not found or column_end == 0 or row_end == 0:
        print(GOOGLE_PREFIX + 'Unable to find \'{}\' in ranges {}'.format(name, google_sheet_ranges))
        return

    # Convert the hex string to RGB values
    red, green, blue = bytes.fromhex(colour_hex)

    # Make an update request
    sheets.batchUpdate(spreadsheetId=google_sheet_id, body={
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "startColumnIndex": column_start,
                        "endColumnIndex": column_end,
                        "startRowIndex": row_start,
                        "endRowIndex": row_end,
                        "sheetId": 0
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": red/255,
                                "green": green/255,
                                "blue": blue/255
                            }
                        }
                    },
                    "fields": 'UserEnteredFormat(BackgroundColor)'
                }
            }
        ]
    }).execute()


def google_token_timer_refresh():
    if not google_refresh_tokens():
        print(GOOGLE_PREFIX + 'Stopping token timer since we cannot refresh anymore...')
        google_token_timer.stop()


##### Start a timer to keep our Google tokens refreshed ################################################################

google_token_timer = RepeatingTimer(600, google_token_timer_refresh)
google_token_timer.start()

##### Set up the Discord bot ###########################################################################################

rolly_discord = discord.Client()

reaction_colours = {'✅':'00ff00', '❔':'ffff00', '❌':'ff0000'}

@rolly_discord.event
async def on_ready():
    print(DISCORD_PREFIX + 'Bot logged in!')

@rolly_discord.event
async def on_message(message):
    # Declare this globally here, since we use it early on, /and/ in a command
    #global discord_bot_channel

    # Make sure this is from the desired server
    if str(message.guild.id) != discord_bot_server:
        # print(DISCORD_PREFIX + 'Got a message from server {} channel {}, expected server {}'.format(message.guild.id, message.channel.id, discord_bot_server))
        return

    # If the bot owner is set, make sure this is from them
    if discord_bot_owners and str(message.author.id) not in discord_bot_owners:
        # print(DISCORD_PREFIX + 'Got a message from user {}, expected one of {}'.format(message.author.id, discord_bot_owners))
        return

    # Determine what prefix was used to address us, if any
    prefix = ''
    if message.content.startswith('#rolly'):
        prefix = '#rolly'
    elif message.content.startswith('<@' + str(discord_id) + '>'):
        prefix = '@{}#{}'.format(rolly_discord.user.name, rolly_discord.user.discriminator)
    else:
        # We weren't addressed, we can stop here
        return

    # Split the command into arguments
    args = message.content.strip().split()[1:]

    # Return a message if no args are provided
    if not args:
        await message.channel.send('Yo, I\'m Rolly. Try `{} create` to start a roll call.'.format(prefix), delete_after=30)
    else:
        # Declare some globals
        global google_sheet_ranges
        global google_sheet_id

        # Filter what command came through
        if args[0] == 'help':
            await message.channel.send('Here\'s a list of commands you can give me:\n' +
                                       '\n' +
                                       '`help` - Shows this help text\n' +
                                       '`create` - Creates a new roll call message\n' +
                                       '`setsheet` - Sets the Google Sheet ID to update\n' +
                                       '`ranges` - Shows the current allowed ranges in the spreadsheet\n' +
                                       '`addranges` - Add one or more allowed ranges for the spreadsheet\n' +
                                       '`setranges` - Sets the ranges to update in the spreadsheet')

        elif args[0] == 'create':
            if len(args) > 1:
                await setup(message.channel, ' '.join(args[1:]))
            else:
                await setup(message.channel)

        elif args[0] == 'setsheet':
            # Make sure we've been given an ID
            if len(args) < 2:
                await message.channel.send('I need a sheet ID to do that.\nYou can get it from the URL e.g. `https://docs.google.com/spreadsheets/d/<sheet id>/`', delete_after=30)
            else:
                # Update and write config to file
                config['Google']['sheet_id'] = args[1]
                google_sheet_id = args[1]
                with open(config_path, 'w') as file:
                    config.write(file)

        elif args[0] == 'ranges':
            # List off the current ranges
            # message_str = 'Current allowed ranges are: `'
            # if len(config['Google']['sheet_ranges']) < 3:
            #     message_str = message_str + '`, and `'.join(config['Google']['sheet_ranges'])
            # else:
            #     message_str = message_str + '`, `'.join(config['Google']['sheet_ranges'][:-1]) + '`, and `' + config['Google']['sheet_ranges'][-1]
            # message_str = message_str + '`'
            message_str = 'Current allowed ranges are: `' + config['Google']['sheet_ranges'] + '`'

            await message.channel.send(message_str)

        elif args[0] == 'addranges':
            # Make sure we've been given a range
            if len(args) < 2:
                await message.channel.send('I need at least one range to do that.\nBy range I mean something like `C2`, `A3:B7`, or `Sheet2!H13:AC139`. You can include several, just separate them with a space.', delete_after=30)
            else:
                # Update and write config to file
                new_ranges = config['Google']['sheet_ranges'] + ' ' + ' '.join(args[1:])
                config['Google']['sheet_ranges'] = new_ranges
                google_sheet_ranges = new_ranges
                with open(config_path, 'w') as file:
                    config.write(file)

        elif args[0] == 'addranges':
            # Make sure we've been given a range
            if len(args) < 2:
                await message.channel.send('I need at least one range to do that.\nBy range I mean something like `C2`, `A3:B7`, or `Sheet2!H13:AC139`. You can include several, just separate them with a space.', delete_after=30)
            else:
                # Update and write config to file
                new_ranges = config['Google']['sheet_ranges'] + ' ' + ' '.join(args[1:])
                config['Google']['sheet_ranges'] = new_ranges
                google_sheet_ranges = new_ranges
                with open(config_path, 'w') as file:
                    config.write(file)

        elif args[0] == 'setranges':
            # Make sure we've been given a range
            if len(args) < 2:
                await message.channel.send('I need at least one range to do that.\nBy range I mean something like `C2`, `A3:B7`, or `Sheet2!H13:AC139`. You can include several, just separate them with a space.', delete_after=30)
            else:
                # Update and write config to file
                config['Google']['sheet_ranges'] = ' '.join(args[1:])
                google_sheet_ranges = ' '.join(args[1:])
                with open(config_path, 'w') as file:
                    config.write(file)

        # TODO: Add better range modification commands (i.e. ranges, addranges, removeranges)

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
        user = channel.guild.get_member(event.user_id)
        emoji = event.emoji.name

        # Ignore reacts that we make
        if str(user.id) == discord_id:
            return

        # Report what's happened
        global reaction_colours
        if emoji in reaction_colours:
            print(DISCORD_PREFIX + '{} reacted with \'{}\', changing their cell to {}'.format(user.display_name, emoji, reaction_colours[emoji]))
        else:
            print(DISCORD_PREFIX + '{} reacted with unsupported emoji \'{}\''.format(user.display_name, emoji))
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
        user = channel.guild.get_member(event.user_id)
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
                    print(DISCORD_PREFIX + '{} removed their \'{}\' react, but they still have a \'{}\' react. Changing their cell to {}'.format(user.display_name, emoji, reaction.emoji, reaction_colours[reaction.emoji]))
                    sheet_update_user(user.display_name, reaction_colours[reaction.emoji])
                    return

        # Didn't find an alternate emoji to use, clear their cell
        print(DISCORD_PREFIX + '{} removed their \'{}\' react, changing their cell to {}'.format(user.display_name, emoji, 'ffffff'))
        sheet_update_user(user.display_name, 'ffffff')

##### Start the Discord bot ############################################################################################

rolly_discord.run(discord_bot_token)

##### Stop the token refresh timer #####################################################################################

google_token_timer.stop()
