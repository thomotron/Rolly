#!/bin/python3
import re
import discord
from argparse import ArgumentParser
from configparser import ConfigParser
from googleapiclient.discovery import build
from oauth2client.client import OAuth2WebServerFlow

##### Define some constants ############################################################################################

DISCORD_PREFIX = '[Discord] '
GOOGLE_PREFIX = '[Google] '

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
          'bot_owner = \n' +
          'bot_server = \n' +
          'bot_channel = \n')
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
if 'bot_channel' not in config['Discord']:
    print('Failed to read config: \'bot_channel\' missing from section \'Discord\'')
    exit(1)

google_id = config['Google']['client_id']
google_secret = config['Google']['client_secret']
google_redirect = config['Google']['redirect_url']
google_sheet_id = config['Google']['sheet_id']
google_sheet_range = config['Google']['sheet_ranges']

discord_id = config['Discord']['client_id']
discord_bot_token = config['Discord']['bot_token']
discord_bot_owner = config['Discord']['bot_owner']
discord_bot_server = config['Discord']['bot_server']
discord_bot_channel = config['Discord']['bot_channel']

##### Parse arguments ##################################################################################################

parser = ArgumentParser()
args = parser.parse_args()

##### Authorise with Google ############################################################################################

flow = OAuth2WebServerFlow(client_id=google_id,
                           client_secret=google_secret,
                           scope='https://www.googleapis.com/auth/spreadsheets',
                           redirect_uri=google_redirect)

google_auth_uri = flow.step1_get_authorize_url()

print("Please authorise yourself at the below URL and paste the code here")
print(google_auth_uri)
google_auth_code = input('Code: ')
google_credentials = flow.step2_exchange(google_auth_code)

##### Set up the Google Sheets service #################################################################################

sheets_service = build('sheets', 'v4', credentials=google_credentials)
sheets = sheets_service.spreadsheets()

##### Define some functions ############################################################################################

async def setup(channel):
    """
    Creates a roll call message in the given channel
    """
    message = await channel.send('Roll call!')
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
    # Query our range from Sheets
    sheet_response = sheets.values().get(spreadsheetId=google_sheet_id, range=google_sheet_range).execute()
    sheet_values = sheet_response.get('values', [])

    # Make sure we got something
    if not sheet_values:
        print(GOOGLE_PREFIX + 'No data found')
        return

    # Try find the target name within the range
    column_start = 0
    column_end = 0
    row_start = 0
    row_end = 0
    for y, row in enumerate(sheet_values):
        for x, value in enumerate(row):
            if value and value in name:
                column_start = x
                column_end = x + 1
                row_start = y
                row_end = y + 1
                break

    # Make sure we got a value
    if column_end == 0 or row_end == 0:
        print(GOOGLE_PREFIX + 'Unable to find \'{}\' in range {}'.format(name, google_sheet_range))
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


##### Set up the Discord bot ###########################################################################################

rolly_discord = discord.Client()

reaction_colours = {'✅':'00ff00', '❔':'ffff00', '❌':'ff0000'}

@rolly_discord.event
async def on_ready():
    print(DISCORD_PREFIX + 'Bot logged in!')

@rolly_discord.event
async def on_message(message):
    # Make sure this is from the desired server and channel
    if str(message.guild.id) != discord_bot_server or str(message.channel.id) != discord_bot_channel:
        # print(DISCORD_PREFIX + 'Got a message from server {} channel {}, expected server {} channel {}'.format(message.guild.id, message.channel.id, discord_bot_server, discord_bot_channel))
        return

    # If the bot owner is set, make sure this is from them
    if discord_bot_owner and str(message.author.id) not in discord_bot_owner:
        # print(DISCORD_PREFIX + 'Got a message from user {}, expected user {}'.format(message.author.id, discord_bot_owner))
        return

    # Ignore messages not intended for us
    if not message.content.startswith('#rolly'):
        return

    # Filter what command came through
    if message.content.startswith('#rolly'):
        await setup(message.channel)
    else:
        return

@rolly_discord.event
async def on_raw_reaction_add(event):
    # Grab the channel and message that was reacted on
    channel = rolly_discord.get_channel(event.channel_id)
    message = await channel.fetch_message(event.message_id)

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
        try:
            print(DISCORD_PREFIX + '{} reacted with \'{}\', changing their cell to {}'.format(user.display_name, emoji, reaction_colours[emoji]))
        except KeyError:
            print(DISCORD_PREFIX + '{} reacted with unsupported emoji \'{}\''.format(user.display_name, emoji))
            return
        finally:
            sheet_update_user(user.display_name, reaction_colours[emoji])


@rolly_discord.event
async def on_reaction_add(reaction, user):
    if reaction.message.author.id != discord_id:
        # print(DISCORD_PREFIX + 'Got react for someone else\'s message')
        return

    global reaction_colours
    if reaction_colours[reaction.emoji]:
        print(DISCORD_PREFIX + '{} reacted with {}, changing their cell to {}'.format(user.name, reaction.emoji, reaction_colours[reaction.emoji]))
    else:
        print(DISCORD_PREFIX + '{} reacted with unsupported emoji {}'.format(reaction.user.name, reaction.emoji))

##### Start the Discord bot ############################################################################################

rolly_discord.run(discord_bot_token)
