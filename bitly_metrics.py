import requests
import json
import pandas as pd
import itertools
from config import *
from itertools import product
from datetime import datetime, timedelta
from tzlocal import get_localzone

# Runtime start
startTime = datetime.now()

DAY = timedelta(1)
local_tz = get_localzone()   # get local timezone
now = datetime.now(local_tz) # get timezone-aware datetime object
naive = now.replace(tzinfo=None) - DAY # same time
yesterday = now - DAY # exactly 24 hours ago, time may differ

# EDIT here for manual run
# START_DATE = datetime.strptime('2022-04-01 00:00:00', '%Y-%m-%d %H:%M:%S')

# START DATE is 2 days minus the current day. Change the 2 to how many days back you want to go
START_DATE = now.replace(tzinfo=None) - (DAY * 2)
START_DATE = START_DATE.strftime("%Y-%m-%d")
START_DATE = START_DATE + ' 00:00:00'
START_DATE = datetime.strptime(START_DATE, '%Y-%m-%d %H:%M:%S')

# Naive End of Day is yesterday at 11:59:59PM ET
naive_eod = naive.strftime("%Y-%m-%d")
naive_eod = naive_eod + ' 23:59:59'
naive_eod = datetime.strptime(naive_eod, '%Y-%m-%d %H:%M:%S')

headers = {
    'Authorization': 'Bearer ' + BITLY_TOKEN,
}

"""
Bitly Metrics
"""
bitly_date_list = []
while START_DATE < naive_eod:

    bitly_date = START_DATE.strftime('%Y-%m-%d') + 'T15:00:00-0700'
    bitly_date_list.append(bitly_date)

    START_DATE += DAY

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

        response = requests.get('https://api-ssl.bitly.com/v4/bitlinks/' + row['bitlink'] + '/referrers', headers=headers, params=params)

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
        response_tag = requests.get('https://api-ssl.bitly.com/v4/bitlinks/' + i, headers=headers)

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

bitly_metrics_tags_df.to_csv(r'postgres/db_data/db_BITLY_load.csv', index=False, header=True)

print('COMPLETED. Run time:', datetime.now() - startTime)




# bitly_df.to_excel(r'tmp/debug.xlsx', index=False)