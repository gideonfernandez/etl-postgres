import time
import psycopg2
import smtplib
from config import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from configparser import ConfigParser

username = MMG_USER
password = MMG_PASSWORD
mail_from = MMG_USER
mail_to = 'gideonfernandez@velocitysense.com'
# mail_subject = 'FAILURE ALERT - MMG Database Update Failed ' + TODAY
mail_body = 'The NN or Bitly table may have not successfully update at ' + TODAY

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

"""
NETWORKNINJA table check
"""
def check_nn_table():
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

        cur.execute('SELECT COUNT(*) FROM mmg.networkninja')
        nn_count = cur.fetchone()

        # close the communication with the PostgreSQL
        cur.close()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed.')
    return nn_count

results = check_nn_table()

for i in results:
    record_count = i

# If the NN table count is 0, then send an alert
if record_count == 0:
    try:
        mimemsg = MIMEMultipart()
        mimemsg['From']=mail_from
        mimemsg['To']=mail_to
        mimemsg['Subject']= 'FAILURE ALERT - NetworkNinja Update Failed ' + TODAY
        mimemsg.attach(MIMEText(mail_body, 'plain'))
        connection = smtplib.SMTP(host='smtp.office365.com', port=587)
        connection.starttls()
        connection.login(username,password)
        connection.send_message(mimemsg)
        connection.quit()
    except Exception as error:
        print(error)


"""
Bitly table check
"""
def check_bitly_table():
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

        cur.execute("SELECT COUNT(*) FROM mmg.bitly WHERE date = (current_date - INTERVAL '1 day')::date")
        bitly_count = cur.fetchone()

        # close the communication with the PostgreSQL
        cur.close()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed.')
    return bitly_count

results = check_bitly_table()

for i in results:
    print(i)
    bitly_record_count = i

# If the NN table count is 0, then send an alert
if bitly_record_count == 0:
    try:
        mimemsg = MIMEMultipart()
        mimemsg['From']=mail_from
        mimemsg['To']=mail_to
        mimemsg['Subject']= 'WARNING ALERT - Check Bitly Table ' + TODAY
        mimemsg.attach(MIMEText(mail_body, 'plain'))
        connection = smtplib.SMTP(host='smtp.office365.com', port=587)
        connection.starttls()
        connection.login(username,password)
        connection.send_message(mimemsg)
        connection.quit()
    except Exception as error:
        print(error)









