import time
from datetime import datetime
from pytz import timezone

fmt = "%Y-%m-%d %H:%M:%S %Z%z"
NOW_TIME = datetime.now(timezone('US/Eastern'))
TODAY = NOW_TIME.strftime(fmt)  #Eastern Time, ie: 2022-05-19 09:36:02 EDT-0400
TIMESTR = time.strftime("%Y%m%d-%H%M%S")
TODAYSTR = time.strftime("%Y%m%d")
TODAY_SLASH = time.strftime("%m/%d/%Y")

BITLY_TOKEN = '1234abcd'

SHAREPOINT_CLIENT_URL = 'https://clientteam.sharepoint.com'
SHAREPOINT_BASE_URL = 'https://clientteam.sharepoint.com/sites/ClientDataTeam/'
SHAREPOINT_SUBFOLDER = '/sites/ClientDataTeam/Shared%20Documents/General/Database/Daily Data Sources/NN/'
SHAREPOINT_USER = 'xyz@email.com'
SHAREPOINT_PASSWORD = '1234'

EMAIL_USER = 'xyz@email.com'
EMAIL_PASSWORD = '1234'

SHAREPOINT_FILE_PATH = '/sites/ClientDataTeam/Shared%20Documents/General/NN/data/'

ISSUES_FILEPATH = r'data/Report Issues_' + TODAYSTR + '.xlsx'
ISSUES_FILENAME = r'Report Issues_' + TODAYSTR + '.xlsx'

RECIPIENT = 'recipient_1@email.com, recipient_2@email.com'

FOLDER_SP_BITLY = site.Folder('Shared%20Documents/General/Database/Daily Table Backups/Bitly')
FOLDER_SP_NN = site.Folder('Shared%20Documents/General/Database/Daily Table Backups/NN')

# API RATE LIMITER
ONE_MINUTE = 60
MAX_CALLS_PER_MINUTE = 30

BITLY_LIST = [
	'bit.ly/site_1',
	'bit.ly/site_2',
	'bit.ly/site_3',
	'bit.ly/site_4',
	'bit.ly/site_5',
	'bit.ly/site_6',
	'bit.ly/site_7',
	'bit.ly/site_8',
	'bit.ly/site_9',
	'bit.ly/site_10',
]