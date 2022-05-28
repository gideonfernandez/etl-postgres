import schedule
import time
import os
from datetime import datetime

# Runtime start
startTime = datetime.now()
print('Date Start Time:', startTime)
def job():
    '''Run scheduled job.'''
    print('Starting Quant Report...')
    time.sleep(5)
    exec(open('quant_report.py').read())
    print('Quant Report Done!')

    print('Starting Target Market Report...')
    time.sleep(5)
    exec(open('target_market.py').read())
    print('Target Market Report Done!')

    # print('Starting Bitly Gather Updated Metrics Report...')
    time.sleep(5)
    exec(open('bitly_metrics.py').read())
    print('Bitly Metrics Captured!')

    print('Starting Postgresql jobs...')
    time.sleep(5)
    os.system('python db_run.py')
    print('Postgresql tables updated!')

    print('Sending email...')
    time.sleep(5)
    exec(open('email_push.py').read())
    print('Email sent')

schedule.every().day.at('14:14').do(job)  #Time in UTC
# schedule.every(30).seconds.do(job)

print('Date End Time:', datetime.now())
print('NetworkNinja & Bitly Jobs COMPLETED. Run time:', datetime.now() - startTime)

while True:
    schedule.run_pending()
    time.sleep(1)