import pandas as pd
import numpy as np
import xlrd
import glob
import re
from config import *

APPROVE_RECAP = 'NEED EC Review \nChange Status\nRECAP needs to be approved'
NO_RECAP = 'NEED EC Review \nNo recap\nCancel?'
TEST_RECAP = 'Event is not approved and contains word "TEST". Confirm and Delete?'

def findWholeWord(w):
    return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search

# Load the PRI report
# for i in glob.glob('data/pri*.csv'):
#     pri_filename = i

# pri_full_df = pd.read_csv(pri_filename, skiprows=2, thousands=r',')

"""
0 - Pull PRI from SHAREPOINT
Gets latest pri_auto file from MMG Data Team > Network Ninja > data
"""
import os.path
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

sharepoint_base_url = SHAREPOINT_BASE_URL
sharepoint_subfolder = SHAREPOINT_SUBFOLDER
sharepoint_user = MMG_USER
sharepoint_password = MMG_PASSWORD

# Constructing Details For Authenticating SharePoint

auth = AuthenticationContext(sharepoint_base_url)

auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
ctx = ClientContext(sharepoint_base_url, auth)
web = ctx.web
ctx.load(web)
ctx.execute_query()
print('Successfully connected to SharePoint: ', web.properties['Title'])

# Constructing Function for getting file details in SharePoint Folder
def folder_details(ctx, sharepoint_subfolder):
    folder = ctx.web.get_folder_by_server_relative_url(sharepoint_subfolder)
    fold_names = []
    sub_folders = folder.files
    ctx.load(sub_folders)
    ctx.execute_query()
    for s_folder in sub_folders:
        fold_names.append(s_folder.properties["Name"] + ',' + s_folder.properties["TimeCreated"])
    return fold_names

# Getting folder details
file_list = folder_details(ctx, sharepoint_subfolder)

# Printing list of files from sharepoint folder and write to dataframe
# print(file_list)
file_list_df = pd.DataFrame([sub.split(",") for sub in file_list], columns=['File', 'Date'])

# Sort dataframe by descending to get the latest file
file_list_df = file_list_df.sort_values(by='Date',ascending=False)

# Get filename of first row
pri_filename = file_list_df['File'].iloc[0]
print(pri_filename)

# Reading File from SharePoint Folder
sharepoint_file = '/sites/MMGDataTeam/Shared%20Documents/General/Database/Daily Data Sources/NetworkNinja/' + pri_filename

file_response = File.open_binary(ctx, sharepoint_file)

# Saving file data directory where quant_report will pull pri data from
with open(f'data/sharepoint_nn_data/{pri_filename}', 'wb') as output_file:
    output_file.write(file_response.content)

pri_full_df = pd.read_csv(f'data/sharepoint_nn_data/{pri_filename}', skiprows=2, thousands=r',')

# Filter PRI report by only looking at CPGI events
NON_CPGI = [
            'CEP'
            ,'Comms'
            ,'DRC'
            ,'DV'
            ,'FQHC'
            ,'HPO'
            ,'Legislative'
            ,'NLM'
            ,'Visibility'
            ]

pri_df = pri_full_df.copy()
pri_df = pri_df[~pri_df['Partner'].isin(NON_CPGI)]

# Delete text from NN2.0 Venue Name field
pri_df['Venue Name'] = pri_df['Venue Name'].str.replace(r" \(you must add id from venues_profiles for this to link\)","", regex=True)

# Add the row for quant columns to check if a recap exists
pri_df['sum'] = pri_df[[
        'Digital Recap | Digital Event Engagement',
        'Digital Recap | # of Digital Event Attendees',
        'Digital Recap | # of Tweets',
        'Digital Recap | Twitter Impressions',
        'Digital Recap | Twitter Engagement',
        'Digital Recap | # of Facebook posts',
        'Digital Recap | Facebook Reach',
        'Digital Recap | Facebook Engagement',
        'Digital Recap | # of Instagram posts',
        'Digital Recap | Instagram Reach',
        'Digital Recap | Instagram Impressions',
        'Digital Recap | Instagram Engagement',
        'Digital Recap | # of Twitter Chat Attendees',
        'Digital Recap | Twitter Chat Engagement',
        'Digital Recap | # of Facebook Live Attendees',
        'Digital Recap | Facebook Live Engagement',
        'Digital Recap | WhatsApp Reach',
        'Digital Recap | LinkedIn Reach',
        'Digital Recap | Blog Engagement',
        'Digital Recap | Google Analytics Page Views',
        'In-Person Recap | Est # of attendees',
        'In-Person Recap | Est # of  Engagement',
        'In-Person Recap | Reach of Tactic',
        'Hybrid Recap | Est # of attendees',
        'Hybrid Recap | # of In-Person Attendees',
        'Hybrid Recap | # of Virtual Attendees',
        'Hybrid Recap | Est # of Engagement',
        'Non-Digital Non-Event Recap | Reach',
        'Non-Digital Non-Event Recap | If printed, how many were distributed?',
        'Legacy Recap | Est # of attendees',
        'Legacy Recap | Est # of Engagement 1',
        'Legacy Recap | Est # of Engagement 2',
        'Legacy Recap | # of YouTube videos',
        'Legacy Recap | YouTube Reach',
        'Legacy Recap | YouTube Engagement',
        'Legacy Recap | # of Tweets',
        'Legacy Recap | Twitter Impressions',
        'Legacy Recap | Twitter Engagement',
        'Legacy Recap | # of Facebook posts',
        'Legacy Recap | Facebook Reach',
        'Legacy Recap | Facebook Engagement',
        'Legacy Recap | # of Instagram posts',
        'Legacy Recap | Instagram Reach',
        'Legacy Recap | Instagram Impressions',
        'Legacy Recap | Instagram Engagement',
        'Legacy Recap | Instagram Exp Audience Reach',
    ]].sum(axis=1)

# Aggregate columns to reduce report columns
pri_df['#-Total Social Media Posts'] = (
                                            pri_df['Digital Recap | # of Tweets']
                                            + pri_df['Digital Recap | # of Facebook posts']
                                            + pri_df['Digital Recap | # of Instagram posts']
                                            + pri_df['Legacy Recap | # of YouTube videos']
                                            + pri_df['Legacy Recap | # of Tweets']
                                            + pri_df['Legacy Recap | # of Facebook posts']
                                            + pri_df['Legacy Recap | # of Instagram posts']
    )

pri_df['#-Total Social Media Impressions'] = (
                                            pri_df['Digital Recap | Twitter Impressions']
                                            + pri_df['Digital Recap | Instagram Impressions']
                                            + pri_df['Legacy Recap | Twitter Impressions']
                                            + pri_df['Legacy Recap | Instagram Impressions']
    )

pri_df['#-Total Social Media Engagement'] = (
                                            pri_df['Digital Recap | Facebook Engagement']
                                            + pri_df['Digital Recap | Twitter Engagement']
                                            + pri_df['Digital Recap | Instagram Engagement']
                                            + pri_df['Digital Recap | Blog Engagement']
                                            + pri_df['Legacy Recap | YouTube Engagement']
                                            + pri_df['Legacy Recap | Twitter Engagement']
                                            + pri_df['Legacy Recap | Facebook Engagement']
                                            + pri_df['Legacy Recap | Instagram Engagement']
    )

pri_df['#-Total Social Media Reach'] = (
                                            pri_df['Digital Recap | Facebook Reach']
                                            + pri_df['Digital Recap | Instagram Reach']
                                            + pri_df['Digital Recap | LinkedIn Reach']
                                            + pri_df['Digital Recap | WhatsApp Reach']
                                            + pri_df['Legacy Recap | YouTube Reach']
                                            + pri_df['Legacy Recap | Facebook Reach']
                                            + pri_df['Legacy Recap | Instagram Reach']
                                            + pri_df['Legacy Recap | Instagram Exp Audience Reach']
    )

pri_df['#-Total Digital Attendees'] = (
                                            pri_df['Digital Recap | # of Digital Event Attendees']
                                            + pri_df['Digital Recap | # of Facebook Live Attendees']
                                            + pri_df['Digital Recap | # of Twitter Chat Attendees']
                                            + pri_df['Hybrid Recap | # of Virtual Attendees']
    )

pri_df['#-Total Digital Engagement'] = (
                                            pri_df['Digital Recap | Digital Event Engagement']
                                            + pri_df['Digital Recap | Facebook Live Engagement']
                                            + pri_df['Digital Recap | Twitter Chat Engagement']
    )

pri_df['#-Total Attendees'] = (
                                            pri_df['In-Person Recap | Est # of attendees']
                                            + pri_df['Legacy Recap | Est # of attendees']
                                            + pri_df[['Hybrid Recap | Est # of attendees', 'Hybrid Recap | # of In-Person Attendees']].max(axis=1)
    )

pri_df['#-Total Engagement'] = (
                                            pri_df['In-Person Recap | Est # of  Engagement']
                                            + pri_df['Hybrid Recap | Est # of Engagement']
                                            + pri_df[['Legacy Recap | Est # of Engagement 1', 'Legacy Recap | Est # of Engagement 2']].max(axis=1)
    )

pri_df['#-Total Reach'] = (
                                            pri_df['In-Person Recap | Reach of Tactic']
                                            + pri_df['Non-Digital Non-Event Recap | Reach']
    )

pri_df['#-Total Page Views'] = pri_df['Digital Recap | Google Analytics Page Views']
pri_df['#-Total Printed Material Distributed'] = pri_df['Non-Digital Non-Event Recap | If printed, how many were distributed?']

pri_df['Accomplishments'] = pri_df['Admin Recap | Accomplishments'].fillna(
                                            pri_df['Legacy Recap | Accomplishments'])

pri_df['Challenges'] = pri_df['Admin Recap | Challenges'].fillna(
                                            pri_df['Legacy Recap | Challenges'])

pri_df['Lessons Learned'] = pri_df['Admin Recap | Lessons Learned'].fillna(
                                            pri_df['Legacy Recap | Lessons Learned'])

pri_df['Barriers to Participation'] = pri_df['Admin Recap | Barriers to participation'].fillna(
                                            pri_df['Legacy Recap | Barriers to Participation'])

pri_df['Best-resonating message'] = pri_df['Admin Recap | Best-resonating message'].fillna(
                                            pri_df['Legacy Recap | Best-resonating Messages'])

pri_df['Concerns expressed'] = pri_df['Admin Recap | Concerns expressed'].fillna(
                                            pri_df['Legacy Recap | Concerns Expressed'])

pri_df['Other notes'] = pri_df['Admin Recap | Other notes'].fillna(
                                            pri_df['Legacy Recap | Other Notes'])

pri_df['Governance Activities Summary'] = pri_df['Admin Recap | Governance Activities Summary'].fillna(
                                            pri_df['Legacy Recap | Governance Activities Summary'])

"""
START: RECAPPED DATFRAME
"""
pri_recapped_df = pri_df[
    (pri_df['Accomplishments'].notnull()) |
    (pri_df['Challenges'].notnull()) |
    (pri_df['Lessons Learned'].notnull()) |
    (pri_df['Barriers to Participation'].notnull()) |
    (pri_df['Best-resonating message'].notnull()) |
    (pri_df['Concerns expressed'].notnull()) |
    (pri_df['Other notes'].notnull()) |
    (pri_df['Governance Activities Summary'].notnull()) |
    (pri_df['sum'] > 0)
    ]

"""
END: RECAPPED DATFRAME
"""

"""
START: NO RECAP DATFRAME
"""
# Check quant and qual recap fields to see if if there's a recap written in
pri_no_recap_df = pri_df[
    (pri_df['Accomplishments'].isnull()) &
    (pri_df['Challenges'].isnull()) &
    (pri_df['Lessons Learned'].isnull()) &
    (pri_df['Barriers to Participation'].isnull()) &
    (pri_df['Best-resonating message'].isnull()) &
    (pri_df['Concerns expressed'].isnull()) &
    (pri_df['Other notes'].isnull()) &
    (pri_df['Governance Activities Summary'].isnull()) &
    (pri_df['sum'] == 0)
    ]

# Add next step values depending on issue type
pri_recapped_df = pri_recapped_df.copy()
pri_no_recap_df = pri_no_recap_df.copy()

pri_recapped_df.loc[:,'*'] = APPROVE_RECAP
pri_no_recap_df.loc[:,'*'] = NO_RECAP

issues_frames = [
        pri_recapped_df,
        pri_no_recap_df,
        ]

pyxis_report_issues_df = pd.concat(issues_frames)

excluded = ['lm cancelled', 'cancelled', 'approved', 'Finished', 'Cancelled']
pyxis_report_issues_df.loc[(pyxis_report_issues_df['Event Status'].isin(excluded)), '*'] = '*'

# If event name contains the word 'Test', review the test activity
# pyxis_report_issues_df.loc[((pyxis_report_issues_df['Name'].str.contains('test', case=False))), '*'] = TEST_RECAP
# Assign the same as above for approved/recapped Legacy events
for idx, row in pyxis_report_issues_df.iterrows():
    if pyxis_report_issues_df.loc[idx, 'Event Status'] != 'Finished':
        if (findWholeWord('test')(pyxis_report_issues_df.loc[idx,'Name']) != None):
            pyxis_report_issues_df.loc[idx, '*'] = 'THIS IS DEFINITELY A TEST EVENT'

# Set date column to date
# Format of this column needs to be set at yyyy-dd-mm so that
# You can properly filter the records by date
pyxis_report_issues_df['Date'] = pd.to_datetime(pyxis_report_issues_df['Date']).dt.date

# Move the last column indicating next step to be the first column
cols = list(pyxis_report_issues_df.columns)
cols = [cols[-1]] + cols[:-1]
pyxis_report_issues_df = pyxis_report_issues_df[cols]

# Select the columns to be used in the final report
# pyxis_report_issues_df = pyxis_report_issues_df.iloc[:, np.r_[0:9]]

pyxis_report_issues_df = pyxis_report_issues_df[[
    '*',
    'Date',
    'Name',
    'Event ID',
    'Legacy Id',
    'Event Status',
    'Partner',
    'Engagement Activity Type',
    '#-Total Social Media Posts',
    '#-Total Social Media Impressions',
    '#-Total Social Media Engagement',
    '#-Total Social Media Reach',
    '#-Total Digital Attendees',
    '#-Total Digital Engagement',
    '#-Total Attendees',
    '#-Total Engagement',
    '#-Total Reach',
    '#-Total Page Views',
    '#-Total Printed Material Distributed',
    ]].copy()

# Sort by the first column indicating next step
pyxis_report_issues_df = pyxis_report_issues_df.sort_values(["*", "Date"], ascending = (False, True))

# This file here writes the final issues report
with pd.ExcelWriter(r'data/Pyxis Report Issues_' + TODAYSTR + '.xlsx') as writer:  # doctest: +SKIP
    sheet_name = 'Pyxis Report Issues_' + TODAYSTR
    print(sheet_name)
    pyxis_report_issues_df.to_excel(writer, index=False, startrow = 0, sheet_name=sheet_name)

    workbook  = writer.book

    tab_1_sheet_name = sheet_name
    worksheet_1 = writer.sheets[tab_1_sheet_name]

    header_format = workbook.add_format({'bold': False, 'text_wrap': False, 'border': 1, 'align': 'center', 'valign': 'left'})

    for col_num, value in enumerate(pyxis_report_issues_df.columns.values):
        worksheet_1.write(0, col_num, value, header_format)

    worksheet_1.set_column(0, 0, 20)
    worksheet_1.set_default_row(20)

# pyxis_report_issues_df.to_excel(r'tmp/debug.xlsx', index=False)