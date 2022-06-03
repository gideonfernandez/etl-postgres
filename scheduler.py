import schedule
import time
import os
import requests
from datetime import datetime

# Runtime start

def job():
    startTime = datetime.now()
    print('Date Start Time:', startTime)

    '''Run scheduled job.'''
    # print('Starting Quant Report...')
    # time.sleep(5)
    # exec(open('quant_report.py').read())
    # print('Quant Report Done!')

    # print('Starting Target Market Report...')
    # time.sleep(5)
    # exec(open('target_market.py').read())
    # print('Target Market Report Done!')

    # print('Starting Bitly Metrics Report...')
    # time.sleep(5)
    # exec(open('bitly_metrics.py').read())
    # print('Bitly Metrics Captured!')

    print('Starting Postgresql jobs...')
    time.sleep(5)
    os.system('python db_run.py')
    print('Postgresql tables updated!')

    # print('Sending email...')
    # time.sleep(5)
    # exec(open('email_push.py').read())
    # print('Email sent')

    print('Date End Time:', datetime.now())
    print('NetworkNinja & Bitly Jobs COMPLETED. Run time:', datetime.now() - startTime)

def alerts():
    print('NN table check...')
    time.sleep(5)
    os.system('python db_check_alerts.py')
    print('NN table check completed')

schedule.every().day.at('18:19').do(job)  #Time in UTC
# schedule.every().day.at('11:15').do(alerts)  #Time in UTC
# schedule.every(15).seconds.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)