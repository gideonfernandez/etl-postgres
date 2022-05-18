import schedule
import time
import os
import psycopg2
from configparser import ConfigParser

def job():
    '''Run scheduled job.'''
    print('Starting Quant Report...')
    time.sleep(5)
    exec(open('quant_report.py').read())
    print('Quant Report Done!')

    print('Starting Target Market Report...')
    time.sleep(5)
    exec(open('target_market.py').read())
    print('Quant Report Done!')

    # print('Starting Bitly Gather Updated Metrics Report...')
    # time.sleep(5)
    # exec(open('bitly_metrics.py').read())
    # print('Bitly Metrics Captured!')

    print('Starting Postgresql jobs...')
    time.sleep(5)
    os.system('python db_run.py')
    print('Postgresql tables updated!')

    print('Sending email...')
    time.sleep(5)
    exec(open('email_push.py').read())
    print('Email sent')

# schedule.every().day.at('07:00').do(job)
schedule.every().day.at('17:00').do(job)
# schedule.every(5).seconds.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)