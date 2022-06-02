import pandas as pd
from config import *

CANCELLED = ['lm cancelled', 'cancelled', 'Cancelled', 'Cancelled Last Minute', 'Missed']

# for i in glob.glob('data/pri*.csv'):
#     pri_filename = i
#
# # Load the PRI report
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

# Reading File from SharePoint Folder
sharepoint_file = '/sites/MMGDataTeam/Shared%20Documents/General/Network Ninja/data/' + pri_filename
file_response = File.open_binary(ctx, sharepoint_file)

# Saving file data directory where quant_report will pull pri data from
with open(f'data/sharepoint_nn_data/{pri_filename}', 'wb') as output_file:
    output_file.write(file_response.content)

pri_full_df = pd.read_csv(f'data/sharepoint_nn_data/{pri_filename}', skiprows=2, thousands=r',')

# Set date column to date
# Format of this column needs to be set at yyyy-dd-mm so that
# You can properly filter the records by date
pri_full_df['Date'] = pd.to_datetime(pri_full_df['Date']).dt.date

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

msr_df = pri_full_df.copy()
msr_df = msr_df[~msr_df['Partner'].isin(NON_CPGI)]

# Filter out cancelled events
msr_df = msr_df[~msr_df['Event Status'].isin(CANCELLED)]

# Filter to MSR reports
msr_df = msr_df[(msr_df['Engagement Activity Type'] == 'Monthly Summary Report')]

msr_df['Accomplishments'] = msr_df['Admin Recap | Accomplishments'].fillna(
                                            msr_df['Legacy Recap | Accomplishments'])

msr_df['Challenges'] = msr_df['Admin Recap | Challenges'].fillna(
                                            msr_df['Legacy Recap | Challenges'])

msr_df['Lessons Learned'] = msr_df['Admin Recap | Lessons Learned'].fillna(
                                            msr_df['Legacy Recap | Lessons Learned'])

msr_df['Barriers to Participation'] = msr_df['Admin Recap | Barriers to participation'].fillna(
                                            msr_df['Legacy Recap | Barriers to Participation'])

msr_df['Best-resonating message'] = msr_df['Admin Recap | Best-resonating message'].fillna(
                                            msr_df['Legacy Recap | Best-resonating Messages'])

msr_df['Concerns expressed'] = msr_df['Admin Recap | Concerns expressed'].fillna(
                                            msr_df['Legacy Recap | Concerns Expressed'])

msr_df['Other notes'] = msr_df['Admin Recap | Other notes'].fillna(
                                            msr_df['Legacy Recap | Other Notes'])

msr_df['Governance Activities Summary'] = msr_df['Admin Recap | Governance Activities Summary'].fillna(
                                            msr_df['Legacy Recap | Governance Activities Summary'])

# Drop rows where all recap responses are empty
msr_df = msr_df.dropna(subset=[
                                'Accomplishments', 
                                'Challenges', 
                                'Lessons Learned',
                                'Barriers to Participation',
                                'Best-resonating message',
                                'Concerns expressed',
                                'Other notes',
                                'Governance Activities Summary',
                                ], how='all')

msr_df = msr_df[[
    'Date',
    'Name',
    'Event ID',
    'Legacy Id',
    'Event Status',
    'Engagement Activity Type',
    'Partner',
    'NN2 Venue Market',
    'NN2 Consortia Partner Target Market',
    'Accomplishments',
    'Challenges',
    'Lessons Learned',
    'Barriers to Participation',
    'Best-resonating message',
    'Concerns expressed',
    'Governance Activities Summary',
    'Other notes',
    ]].copy()

msr_df = msr_df.sort_values(["Date"], ascending = (True))

# Write to file
with pd.ExcelWriter(r'data/FINAL_Partner Monthly Summary Report - MMM 2022.xlsx') as writer:  # doctest: +SKIP
    msr_df.to_excel(writer, index=False, startrow = 0, sheet_name = 'PMSR MMM 2022')

    workbook  = writer.book

    # TAB 1
    tab_1_sheet_name = 'PMSR MMM 2022'
    worksheet_1 = writer.sheets[tab_1_sheet_name]

    header_format = workbook.add_format({'bold': False, 'text_wrap': False, 'fg_color': '#4F81BD', 'color': '#FFFFFF', 'border': 0, 'align': 'center', 'valign': 'center'})
    for col_num, value in enumerate(msr_df.columns.values):
        worksheet_1.write(0, col_num, value, header_format)