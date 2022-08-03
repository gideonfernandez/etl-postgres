import time
import psycopg2
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
from configparser import ConfigParser
from config import *

# sharepoint_mmg_url = 'https://montagemarketing.sharepoint.com/sites/MMGDataTeam'
sharepoint_mmg_url = 'https://montagemarketing.sharepoint.com'

sharepoint_base_url = SHAREPOINT_BASE_URL
sharepoint_subfolder = SHAREPOINT_SUBFOLDER
sharepoint_user = MMG_USER
sharepoint_password = MMG_PASSWORD
# sharepoint_user = 'gfernandez@montagemarketinggroup.com'
# sharepoint_password = 'gf*2100ea'


#Database Table Backup SP Locations 
authcookie = Office365(sharepoint_mmg_url, username=sharepoint_user, password=sharepoint_password).GetCookies()
# site = Site(sharepoint_base_url, version=Version.v365, authcookie=authcookie);


"""
Upload backup files to SHAREPOINT
"""
# import os.path
# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.sharepoint.files.file import File

# sharepoint_mmg_url = SHAREPOINT_MMG_URL
# sharepoint_base_url = SHAREPOINT_BASE_URL
# sharepoint_subfolder = SHAREPOINT_SUBFOLDER
# sharepoint_user = MMG_USER
# sharepoint_password = MMG_PASSWORD

# auth = AuthenticationContext(sharepoint_base_url)

# auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
# ctx = ClientContext(sharepoint_base_url, auth)
# web = ctx.web
# ctx.load(web)
# ctx.execute_query()
# print('Successfully connected to SharePoint: ', web.properties['Title'])


# folder_sp_bitly = web.Folder('Shared%20Documents/General/Database/Daily Table Backups/Bitly')
# folder_sp_nn = web.Folder('Shared%20Documents/General/Database/Daily Table Backups/NetworkNinja')



# folder_sp_bitly = site.Folder('Shared%20Documents/General/Database/Daily Table Backups/Bitly')
# folder_sp_nn = site.Folder('Shared%20Documents/General/Database/Daily Table Backups/NetworkNinja')

print('SUCCESS')

"""
SHAREPOINT END
"""


def config(filename='database.ini', section='postgresql'):
    # create a parser
    parser = ConfigParser()
    # read config file
    parser.read(filename)

    # get section, default to postgresql
    db = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))

    return db

def backup_nn_table():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        # print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)

        # create a cursor
        cur = conn.cursor()

        outputquery = 'copy mmg.networkninja to stdout with csv header'.format('select * from mmg.networkninja')

        timestamped = TIMESTR
        with open(f'data/db/nn_daily_backups/nn_bkp_' + timestamped + '.csv', 'wb') as f:
            nn_filename = 'nn_bkp_' + timestamped + '.csv'
            
            cur.copy_expert(outputquery, f)
        
        with open(f'data/db/nn_daily_backups/' + nn_filename, 'rb') as content_file:
            nn_content = content_file.read()

        folder_sp_nn.upload_file(nn_content, nn_filename)
        print('Uploaded '+ nn_filename + ' to Database > Daily Table Backups > NetworkNinja')

        # close the communication with the PostgreSQL
        cur.close()

    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Networkninja table backed up')

def clear_nn_table():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        # print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)

        # create a cursor
        cur = conn.cursor()

        print('Removing NN data')
        cur.execute(
            'DELETE FROM mmg.networkninja; COMMIT;'
            )

        # close the communication with the PostgreSQL
        cur.close()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed.')

def load_nn_table():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        # print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)

        # create a cursor
        cur = conn.cursor()

        copy_command = f"COPY mmg.networkninja FROM STDIN CSV HEADER;"
        cur.copy_expert(copy_command, open(r'data/db/db_NN_load.csv', "r"))
        cur.execute('COMMIT;')

        # close the communication with the PostgreSQL
        cur.close()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed.')

def backup_bitly_table():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        # print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)

        # create a cursor
        cur = conn.cursor()

        outputquery = 'copy mmg.bitly to stdout with csv header'.format('select * from mmg.bitly')

        timestamped = TIMESTR
        with open(f'data/db/bitly_daily_backups/bitly_bkp_' + timestamped + '.csv', 'wb') as f:
            bitly_filename = 'bitly_bkp_' + timestamped + '.csv'
            
            cur.copy_expert(outputquery, f)
        
        with open(f'data/db/bitly_daily_backups/' + bitly_filename, 'rb') as content_file:
            bitly_content = content_file.read()

        folder_sp_bitly.upload_file(bitly_content, bitly_filename)
        print('Uploaded '+ bitly_filename + ' to Database > Daily Table Backups > Bitly')

        # close the communication with the PostgreSQL
        cur.close()

    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Bitly table backed up')

def update_bitly_table():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        # print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)

        # create a cursor
        cur = conn.cursor()

        copy_command = f"COPY mmg.bitly FROM STDIN CSV HEADER;"
        cur.copy_expert(copy_command, open(r'data/db/db_bitly_load.csv', "r"))
        cur.execute('COMMIT;')

        # close the communication with the PostgreSQL
        cur.close()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed.')

# Run the SQL commands
# backup_nn_table()
# time.sleep(5)
# clear_nn_table()
# time.sleep(5)
# load_nn_table()
# time.sleep(5)
# backup_bitly_table()
# time.sleep(5)
# update_bitly_table()