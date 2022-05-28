import requests
import json
import pandas as pd
import itertools
from config import *
from itertools import product
from datetime import datetime, timedelta
from ratelimit import limits, RateLimitException, sleep_and_retry

# Runtime start
startTime = datetime.now()
DAY = timedelta(1)

# EDIT here for manual run
# START_DATE = datetime.strptime('2022-04-01 00:00:00', '%Y-%m-%d %H:%M:%S')

# START DATE is Yesterday. Multiply "DAY" times the # of days to go back
START_DATE = NOW_TIME.replace(tzinfo=None) - DAY
START_DATE = START_DATE.strftime("%Y-%m-%d")
START_DATE, YDAY_END_DATE = START_DATE + ' 00:00:00', START_DATE + ' 23:59:59'

YDAY_START_DATE = datetime.strptime(START_DATE, '%Y-%m-%d %H:%M:%S')
YDAY_END_DATE = datetime.strptime(YDAY_END_DATE, '%Y-%m-%d %H:%M:%S')

headers = {
    'Authorization': 'Bearer ' + BITLY_TOKEN,
}

@sleep_and_retry
@limits(calls=MAX_CALLS_PER_MINUTE, period=ONE_MINUTE)
def call_bitly_referrals(url, headers, params):
    response = requests.get(url, headers=headers, params=params)
    return response

@sleep_and_retry
@limits(calls=MAX_CALLS_PER_MINUTE, period=ONE_MINUTE)
def call_bitly_tags(url, headers):
    response_tag = requests.get(url, headers=headers)
    return response_tag

"""
Bitly Metrics
"""
bitly_date_list = []
while YDAY_START_DATE < YDAY_END_DATE:

    bitly_date = YDAY_START_DATE.strftime('%Y-%m-%d') + 'T15:00:00-0700'
    bitly_date_list.append(bitly_date)

    YDAY_START_DATE += DAY

iterate_list = list(map(list, product(BITLY_LIST, bitly_date_list)))

df = pd.DataFrame(iterate_list[0:],columns=['bitlink', 'date'])

dfs = []
for index, row in df.iterrows():

    try:
        params = (
            ('unit', 'day'),
            ('units', 1),
            ('unit_reference', row['date']),
        )

        # response = requests.get('https://api-ssl.bitly.com/v4/bitlinks/' + row['bitlink'] + '/referrers', headers=headers, params=params)
        response = call_bitly_referrals('https://api-ssl.bitly.com/v4/bitlinks/' + row['bitlink'] + '/referrers', headers, params)
        
        metrics = response.json()
        referrers_df = pd.json_normalize(metrics, record_path='metrics')

        referrers_df['Date'] = row['date']
        referrers_df['Bitly'] = row['bitlink']

        dfs.append(referrers_df)

    except KeyError:
        continue

bitly_df = pd.concat(dfs, ignore_index=True)

"""
Bitly TAGS
"""
dfs_tag = []
for i in BITLY_LIST:
    try:
        # response_tag = requests.get('https://api-ssl.bitly.com/v4/bitlinks/' + i, headers=headers)
        response_tag = call_bitly_tags('https://api-ssl.bitly.com/v4/bitlinks/' + i, headers)

        tags = response_tag.json()
        tags_df = pd.json_normalize(tags, record_path='tags')

        tags_df['Bitly'] = i

        dfs_tag.append(tags_df)

    except KeyError:
        continue

bitly_tags_df = pd.concat(dfs_tag, ignore_index=True)

bitly_tags_df = bitly_tags_df.rename(columns={
    0: 'Tags',
    })

# Group tags into a single record, delimited by slash
bitly_tags_df = bitly_tags_df.groupby(['Bitly'])['Tags'].apply(lambda x: ' / '.join(x)).reset_index()

"""
MERGE bitly metrics and tags
"""
bitly_metrics_tags_df = bitly_df.merge(bitly_tags_df, on='Bitly', how='left', indicator=False)

"""
CONSTRUCT final dataframe
"""
bitly_metrics_tags_df = bitly_metrics_tags_df.rename(columns={
    'value': 'Source',
    'clicks': 'Clicks',
    'tags': 'Tags',
    })

bitly_metrics_tags_df = bitly_metrics_tags_df[[
    'Date',
    'Bitly',
    'Source',
    'Tags',
    'Clicks',
    ]].copy()

bitly_metrics_tags_df['Date'] = pd.to_datetime(bitly_metrics_tags_df['Date']).dt.date

bitly_metrics_tags_df.to_csv(r'data/db/db_bitly_load.csv', index=False, header=True)

print('Bitly capture COMPLETED. Run time:', datetime.now() - startTime)




# bitly_df.to_excel(r'tmp/debug.xlsx', index=False)