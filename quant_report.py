import pandas as pd
import xlrd
import glob
import numpy as np
from config import *

from shareplum import Office365
from shareplum import Site
from shareplum.site import Version

APPROVED = ['approved', 'recapped', 'Finished', 'Recapped']
NON_APPROVED = ['scheduled', 'requested', 'Scheduled', 'Requested', 'Request Rejected', 'Staffed']
CANCELLED = ['lm cancelled', 'cancelled', 'Cancelled', 'Cancelled Last Minute', 'Missed']

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

'''
1. LOAD PRI (from Sharepoint)
'''
# ENTER HERE to use local pri file, otherwise it will use the Sharepoint file
# for i in glob.glob('data/pri*.csv'):
#     pri_filename = i

# pri_full_df = pd.read_csv(pri_filename, skiprows=2, thousands=r',')

nn1_full_df = pd.read_csv(r'target_market/build_tm_inputs/nn1_jan1_2017_mar31_2022.csv', thousands=r',')
nn1_full_df = nn1_full_df[[
    'Event ID',
    ]].copy()

pri_full_df = pd.read_csv(f'data/sharepoint_nn_data/{pri_filename}', skiprows=2, thousands=r',', low_memory=False)

pri_full_df = pri_full_df.merge(nn1_full_df.drop_duplicates(),
    on=['Event ID'], how='left', indicator=True)

pri_full_df = pri_full_df[pri_full_df['_merge'].isin(['left_only'])]

pri_full_df = pri_full_df.drop(['_merge'], axis=1)

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

pri_full_df = pri_full_df.copy()
# pri_full_df = pri_full_df[~pri_full_df['Partner'].isin(NON_CPGI)]  #Comment out for all partners

# Delete text from NN2.0 Venue Name field
pri_full_df['Venue Name'] = pri_full_df['Venue Name'].str.replace(r" \(you must add id from venues_profiles for this to link\)","", regex=True)

# Set date column to date
# Format of this column needs to be set at yyyy-dd-mm so that
# You can properly filter the records by date

# Clean up incorrect date for event ID 16054
pri_full_df.loc[(pri_full_df['Event ID'] == 16054), 'Date'] = '09/23/2020'
pri_full_df['Date'] = pd.to_datetime(pri_full_df['Date']).dt.date

'''
2.  Filter out the data we don't need 
'''
# PRI Report
quant_full_df = pri_full_df.copy()
# quant_full_df = pri_full_df[~pri_full_df['Event Status'].isin(CANCELLED)]

'''
CREATE REPORT DATAFRAMES START HERE
Event Recap Type Id Mapping
    2 - In-Person
    3 - Non-Digital
    4 - Digital
    5 - Communication Resource
    7 - Hybrid
    8 - Admin
'''

'''
3.  Break out the different dataframes that will be used to concatonate as the final report 
'''
# 0. Reports will look at approved events only
# Unless it's a Needs Review report
approved_df = quant_full_df.copy()
approved_df = quant_full_df.loc[quant_full_df['Event Status'].isin(APPROVED)]

# 1. Needs review
# All non-approved events go here

# THESE TWO BOOKMARKS IS DROPPING THE MISSING EVENTS
needs_review_df = quant_full_df.loc[quant_full_df['Event Status'].isin(NON_APPROVED)]
needs_review_df = needs_review_df.loc[(needs_review_df['Event Recap Type Id'] != 10)]
needs_review_df = needs_review_df.copy()
needs_review_df['TYPE'] = 'NEEDS REVIEW'

'''
2. ADMIN RECAPS
'''
# 2-Admin Reports, recap type id = 8
admin_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
    (approved_df['Engagement Activity Type'] != 'Monthly Summary Report') &
    (approved_df['Engagement Activity Type'] != 'MSRS') &
    (approved_df['Engagement Activity Type'] != 'Webinar') & 
    (approved_df['Engagement Activity Type'] != 'Other Digital') &
    (approved_df['Engagement Activity Type'] != 'Digital Asset Distribution') &
    (approved_df['Engagement Activity Type'] != 'Digital Collateral Distribution (email, newsletters, flyers)') &
    (approved_df['Engagement Activity Type'] != 'Digital Events') &
    (approved_df['Engagement Activity Type'] != 'Social Media') &
    (approved_df['Engagement Activity Type'] != 'VAP - Live Demo/Engagement') &
    (approved_df['Engagement Activity Type'] != 'Earn Media') &
    (approved_df['Engagement Activity Type'] != 'Other - Non-Digital/Non-Event') &
    (approved_df['Engagement Activity Type'] != 'Workshops') &
    (approved_df['Engagement Activity Type'] != 'CPGI Other') &
    (approved_df['Engagement Activity Type'] != 'Governance')
    ))]
admin_df = admin_df.copy()
admin_df['TYPE'] = 'ADMIN'

# 2-Admin - Digital
admin_digital_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
            (approved_df['Engagement Activity Type'] == 'Webinar') | 
            (approved_df['Engagement Activity Type'] == 'Other Digital') |
            (approved_df['Engagement Activity Type'] == 'Digital Asset Distribution') |
            (approved_df['Engagement Activity Type'] == 'Digital Collateral Distribution (email, newsletters, flyers)') | 
            (approved_df['Engagement Activity Type'] == 'Digital Events') |
            (approved_df['Engagement Activity Type'] == 'VAP - Live Demo/Engagement') |
            (approved_df['Engagement Activity Type'] == 'Earn Media')
            ))]
admin_digital_df = admin_digital_df.copy()
admin_digital_df['TYPE'] = 'DIGITAL'

# 2-Admin - Non-Digital
admin_nondigital_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
            (approved_df['Engagement Activity Type'] == 'Other - Non-Digital/Non-Event')
            ))]
admin_nondigital_df = admin_nondigital_df.copy()
admin_nondigital_df['TYPE'] = 'NON-DIGITAL'

# 2-Admin - In-Person
admin_inperson_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
            (approved_df['Engagement Activity Type'] == 'Workshops') |
            (approved_df['Engagement Activity Type'] == 'CPGI Other')
            ))]
admin_inperson_df = admin_inperson_df.copy()
admin_inperson_df['TYPE'] = 'IN-PERSON'

# 2-Admin - Social Media
admin_sm_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
            (approved_df['Engagement Activity Type'] == 'Social Media')
            ))]
admin_sm_df = admin_sm_df.copy()
admin_sm_df['TYPE'] = 'SOCIAL MEDIA'

# 2-Admin - Governance
admin_governance_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
            (approved_df['Engagement Activity Type'] == 'Governance')
            ))]
admin_governance_df = admin_governance_df.copy()
admin_governance_df['TYPE'] = 'GOVERNANCE'

'''
MSR Recaps
'''
# MSR
msr_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 8) & (
            (approved_df['Engagement Activity Type'] == 'Monthly Summary Report') |
            (approved_df['Engagement Activity Type'] == 'MSRS')
            ))]
msr_df = msr_df.copy()
msr_df['TYPE'] = 'MSR'

'''
3. DIGITAL RECAPS
'''

# 3B. Digital, recap type id = 4 and not Social Media
digital_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 4) & (approved_df['Engagement Activity Type'] != 'Social Media'))]
digital_df = digital_df.copy()
digital_df['TYPE'] = 'DIGITAL'

# 3C. Social Media, recap type id = 4 and is Social Media
social_media_df = approved_df.loc[((approved_df['Event Recap Type Id'] == 4) & (approved_df['Engagement Activity Type'] == 'Social Media'))]
social_media_df = social_media_df.copy()
social_media_df['TYPE'] = 'SOCIAL MEDIA'

# 4. In-Person, recap type id = 2
in_person_df = approved_df.loc[(approved_df['Event Recap Type Id'] == 2)]
in_person_df = in_person_df.copy()
in_person_df['TYPE'] = 'IN-PERSON'

# 5. Non-Digital/Non-Event, recap type id = 3
non_dig_df = approved_df.loc[(approved_df['Event Recap Type Id'] == 3)]
non_dig_df = non_dig_df.copy()
non_dig_df['TYPE'] = 'NON-DIGITAL'

# 6. Hybrid, recap type id = 7
hybrid_df = approved_df.loc[(approved_df['Event Recap Type Id'] == 7)]
hybrid_df = hybrid_df.copy()
hybrid_df['TYPE'] = 'HYBRID'

# 6. ALL Legacy Events - Not Tagged for an Activity
legacy_df = quant_full_df.loc[(quant_full_df['Event Recap Type Id'] == 10)]
legacy_df = legacy_df.copy()
legacy_df['TYPE'] = 'LEGACY'

# 7. CancelLed
cancelled_df = quant_full_df.loc[quant_full_df['Event Status'].isin(CANCELLED)]
cancelled_df = cancelled_df.copy()
cancelled_df['TYPE'] = 'CANCELLED'

'''
Compile all the dataframes here
'''
final_rpts_frames = [
        social_media_df,
        admin_sm_df,
        digital_df,
        admin_digital_df,
        msr_df,
        in_person_df,
        admin_inperson_df,
        non_dig_df,
        admin_nondigital_df,
        hybrid_df,
        admin_df,
        admin_governance_df,
        legacy_df,
        needs_review_df,
        cancelled_df,
        ]

final_qnt_report_df = pd.concat(final_rpts_frames)

# Catch Legacy Events with a recap so that they're counted in
# 1. Needs review correction with quant values
# 2. Venue name with 'nan' values
final_qnt_report_df = final_qnt_report_df.reset_index(drop=True)

# Add the row for quant columns to check if a recap exists
final_qnt_report_df['sum'] = final_qnt_report_df[[
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

# Check for Legacy MSR events to see if a recap was submitted. If yes, then assign a value of 1 to sum column
final_qnt_report_df.loc[(
    (final_qnt_report_df['Legacy Recap | Challenges'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Lessons Learned'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Accomplishments'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Target/Primary Population'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Lessons Learned'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Best-resonating Messages'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Governance Activities Summary'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Concerns Expressed'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Barriers to Participation'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Tactics and strategies'].notnull()) |
    (final_qnt_report_df['Legacy Recap | Other Notes'].notnull())
    ), 'sum'] = 1

# Assign Recap Type
for idx, row in final_qnt_report_df.iterrows():
    if final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 2:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'In-Person'
    elif final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 3:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Non-Digital'
    elif final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 4:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Digital'
    elif final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 5:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Communication Resource'
    elif final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 7:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Hybrid'
    elif final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 8:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Admin'
    elif final_qnt_report_df.loc[idx, 'Event Recap Type Id'] == 10:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Legacy'
    else:
        final_qnt_report_df.loc[idx, 'Recap Type Submitted'] = 'Not Recapped'

# Convert the fields to string to perform the event type evaluation
final_qnt_report_df['Venue Name'] = final_qnt_report_df['Venue Name'].astype(str)
final_qnt_report_df['Partner'] = final_qnt_report_df['Partner'].astype(str)

# If type is Needs Review and has a recap, then assign a Type to event so it's counted in
for idx, row in final_qnt_report_df.iterrows():
    if final_qnt_report_df.loc[idx, 'TYPE'] == 'NEEDS REVIEW' and final_qnt_report_df.loc[idx,'sum'] > 0:
        # In-Person
        if (final_qnt_report_df.loc[idx,'Partner'].find('In-Person') != -1) or (final_qnt_report_df.loc[idx,'Venue Name'].find('In-Person') != -1) or (final_qnt_report_df.loc[idx,'Partner'].find('Live') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'IN-PERSON'
        # Non-Digital
        elif (final_qnt_report_df.loc[idx,'Partner'].find('Non-Digital') != -1) or (final_qnt_report_df.loc[idx,'Partner'].find('Non Digital,Non Event') != -1) or (final_qnt_report_df.loc[idx,'Venue Name'].find('Non-Digital') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'NON-DIGITAL'
        # Social Media
        elif (final_qnt_report_df.loc[idx,'Engagement Activity Type'].find('Social Media') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'SOCIAL MEDIA'
        # Digital
        elif (final_qnt_report_df.loc[idx,'Partner'].find('Digital') != -1) or (final_qnt_report_df.loc[idx,'Partner'].find('Virtual') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'DIGITAL'
        # Monthly Summary Report
        elif (final_qnt_report_df.loc[idx,'Engagement Activity Type'].find('Monthly Summary Report') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'MSR'

# Assign the same as above for approved/recapped Legacy events
for idx, row in final_qnt_report_df.iterrows():
    if final_qnt_report_df.loc[idx, 'TYPE'] == 'LEGACY':
        # In-Person
        if (final_qnt_report_df.loc[idx,'Partner'].find('In-Person') != -1) or (final_qnt_report_df.loc[idx,'Venue Name'].find('In-Person') != -1) or (final_qnt_report_df.loc[idx,'Partner'].find('Live') != -1) or (final_qnt_report_df.loc[idx,'Engagement Activity Type'].find('Events') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'IN-PERSON'
        # Non-Digital
        elif (final_qnt_report_df.loc[idx,'Partner'].find('Non-Digital') != -1) or (final_qnt_report_df.loc[idx,'Partner'].find('Non Digital,Non Event') != -1) or (final_qnt_report_df.loc[idx,'Venue Name'].find('Non-Digital') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'NON-DIGITAL'
        # Social Media
        elif (final_qnt_report_df.loc[idx,'Engagement Activity Type'].find('Social Media') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'SOCIAL MEDIA'
        # Digital
        elif (final_qnt_report_df.loc[idx,'Partner'].find('Digital') != -1) or (final_qnt_report_df.loc[idx,'Partner'].find('Virtual') != -1) or (final_qnt_report_df.loc[idx,'Engagement Activity Type'].find('Webinar') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'DIGITAL'
        # Monthly Summary Report
        elif (final_qnt_report_df.loc[idx,'Engagement Activity Type'].find('Monthly Summary Report') != -1):
            final_qnt_report_df.loc[idx, 'TYPE'] = 'MSR'

# Aggregate columns to reduce report columns
final_qnt_report_df['#-Total Social Media Posts'] = (
                                            final_qnt_report_df['Digital Recap | # of Tweets']
                                            + final_qnt_report_df['Digital Recap | # of Facebook posts']
                                            + final_qnt_report_df['Digital Recap | # of Instagram posts']
                                            + final_qnt_report_df['Legacy Recap | # of YouTube videos']
                                            + final_qnt_report_df['Legacy Recap | # of Tweets']
                                            + final_qnt_report_df['Legacy Recap | # of Facebook posts']
                                            + final_qnt_report_df['Legacy Recap | # of Instagram posts']
    )

final_qnt_report_df['#-Total Social Media Impressions'] = (
                                            final_qnt_report_df['Digital Recap | Twitter Impressions']
                                            + final_qnt_report_df['Digital Recap | Instagram Impressions']
                                            + final_qnt_report_df['Legacy Recap | Twitter Impressions']
                                            + final_qnt_report_df['Legacy Recap | Instagram Impressions']
    )

final_qnt_report_df['#-Total Social Media Engagement'] = (
                                            final_qnt_report_df['Digital Recap | Facebook Engagement']
                                            + final_qnt_report_df['Digital Recap | Twitter Engagement']
                                            + final_qnt_report_df['Digital Recap | Instagram Engagement']
                                            + final_qnt_report_df['Digital Recap | Blog Engagement']
                                            + final_qnt_report_df['Legacy Recap | YouTube Engagement']
                                            + final_qnt_report_df['Legacy Recap | Twitter Engagement']
                                            + final_qnt_report_df['Legacy Recap | Facebook Engagement']
                                            + final_qnt_report_df['Legacy Recap | Instagram Engagement']
    )

final_qnt_report_df['#-Total Social Media Reach'] = (
                                            final_qnt_report_df['Digital Recap | Facebook Reach']
                                            + final_qnt_report_df['Digital Recap | Instagram Reach']
                                            + final_qnt_report_df['Digital Recap | LinkedIn Reach']
                                            + final_qnt_report_df['Digital Recap | WhatsApp Reach']
                                            + final_qnt_report_df['Legacy Recap | YouTube Reach']
                                            + final_qnt_report_df['Legacy Recap | Facebook Reach']
                                            + final_qnt_report_df['Legacy Recap | Instagram Reach']
                                            # + final_qnt_report_df['Legacy Recap | Instagram Exp Audience Reach']
    )

final_qnt_report_df['#-Total Digital Attendees'] = (
                                            final_qnt_report_df['Digital Recap | # of Digital Event Attendees']
                                            + final_qnt_report_df['Digital Recap | # of Facebook Live Attendees']
                                            + final_qnt_report_df['Digital Recap | # of Twitter Chat Attendees']
                                            + final_qnt_report_df['Hybrid Recap | # of Virtual Attendees']
    )

final_qnt_report_df['#-Total Digital Engagement'] = (
                                            final_qnt_report_df['Digital Recap | Digital Event Engagement']
                                            + final_qnt_report_df['Digital Recap | Facebook Live Engagement']
                                            + final_qnt_report_df['Digital Recap | Twitter Chat Engagement']
    )

final_qnt_report_df['#-Total Attendees'] = (
                                            final_qnt_report_df['In-Person Recap | Est # of attendees']
                                            + final_qnt_report_df['Legacy Recap | Est # of attendees']
                                            + final_qnt_report_df[['Hybrid Recap | Est # of attendees', 'Hybrid Recap | # of In-Person Attendees']].max(axis=1)
    )

final_qnt_report_df['#-Total Engagement'] = (
                                            final_qnt_report_df['In-Person Recap | Est # of  Engagement']
                                            + final_qnt_report_df['Hybrid Recap | Est # of Engagement']
                                            + final_qnt_report_df[['Legacy Recap | Est # of Engagement 1', 'Legacy Recap | Est # of Engagement 2']].max(axis=1)
    )

final_qnt_report_df['#-Total Reach'] = (
                                            final_qnt_report_df['In-Person Recap | Reach of Tactic']
                                            + final_qnt_report_df['Non-Digital Non-Event Recap | Reach']
    )

final_qnt_report_df['#-Total Page Views'] = final_qnt_report_df['Digital Recap | Google Analytics Page Views']
final_qnt_report_df['#-Total Material Distributed'] = final_qnt_report_df['Non-Digital Non-Event Recap | If printed, how many were distributed?']

final_qnt_report_df['Target Population'] = final_qnt_report_df['Digital Recap | Target/Primary Population'].fillna(
                                            final_qnt_report_df['In-Person Recap | Target/Primary Population']).fillna(
                                            final_qnt_report_df['Hybrid Recap | Target/Primary Population']).fillna(
                                            final_qnt_report_df['Legacy Recap | Target/Primary Population'])   

final_qnt_report_df['Partners Recapped'] = final_qnt_report_df[[
    'Partners | Digital - CPGI',
    'Partners | Digital - CEP',
    'Partners | Digital - FQHC',
    'Partners | Digital - DV',
    'Partners | In-Person - CPGI',
    'Partners | In-Person - CEP',
    'Partners | In-Person - DV',
    'Partners | In-Person - HPO',
    'Partners | In-Person - FQHC',
    'Partners | Hybrid - CPGI',
    'Partners | Hybrid - CEP',
    'Partners | Hybrid - DV',
    'Partners | Hybrid - HPO',
    'Partners | Hybrid - FQHC',
    'Partners | NonDig NonEvent - CPGI',
    'Partners | NonDig NonEvent - CEP',
    'Partners | NonDig NonEvent - DV',
    'Partners | NonDig NonEvent - HPO',
    'Partners | NonDig NonEvent - FQHC',
    ]].apply(lambda x: ','.join(x.dropna()), axis=1)

final_qnt_report_df['CTA'] = final_qnt_report_df[[
    'Digital Recap | CTA',
    'In-Person Recap | CTA',
    'Non-Digital Non-Event Recap | CTA',
    'Hybrid Recap | CTA',
    ]].apply(lambda x: ','.join(x.dropna()), axis=1)

final_qnt_report_df['CTA Text'] = final_qnt_report_df[[
    'Digital Recap | CTA Text 1',
    'Digital Recap | CTA Text 2',
    'Digital Recap | CTA Text 3',
    'Digital Recap | CTA Text 4',
    'In-Person Recap | CTA Text 1',
    'In-Person Recap | CTA Text 2',
    'In-Person Recap | CTA Text 3',
    'In-Person Recap | CTA Text 4',
    'Non-Digital Non-Event Recap | CTA Text 1',
    'Non-Digital Non-Event Recap | CTA Text 2',
    'Non-Digital Non-Event Recap | CTA Text 3',
    'Non-Digital Non-Event Recap | CTA Text 4',
    'Hybrid Recap | CTA Text 1',
    'Hybrid Recap | CTA Text 2',
    'Hybrid Recap | CTA Text 3',
    'Hybrid Recap | CTA Text 4',
    ]].apply(lambda x: ', '.join(x.dropna()), axis=1)

# Research Facing Field
for idx, row in final_qnt_report_df.iterrows():
    # If RF1 is not null and RF2 is null, then use RF1
    if final_qnt_report_df.loc[idx, 'Research Facing 1'] != ' ' and final_qnt_report_df.loc[idx, 'Research Facing 2'] == ' ':
        final_qnt_report_df.loc[idx, 'Research Facing'] = final_qnt_report_df.loc[idx, 'Research Facing 1']
    # If RF1 is null and RF2 is not null, then use RF2
    elif final_qnt_report_df.loc[idx, 'Research Facing 1'] == ' ' and final_qnt_report_df.loc[idx, 'Research Facing 2'] != ' ':
        final_qnt_report_df.loc[idx, 'Research Facing'] = final_qnt_report_df.loc[idx, 'Research Facing 2']
    # If both RF fields are not null then use RF1
    elif final_qnt_report_df.loc[idx, 'Research Facing 1'] != ' ' and final_qnt_report_df.loc[idx, 'Research Facing 2'] != ' ':
        final_qnt_report_df.loc[idx, 'Research Facing'] = final_qnt_report_df.loc[idx, 'Research Facing 1']
    # If both RF fields are null then null
    elif final_qnt_report_df.loc[idx, 'Research Facing 1'] == ' ' and final_qnt_report_df.loc[idx, 'Research Facing 2'] == ' ':
        final_qnt_report_df.loc[idx, 'Research Facing'] = ''

# Monthly Summary Reports
final_qnt_report_df['Accomplishments'] = final_qnt_report_df['Admin Recap | Accomplishments'].fillna(
                                            final_qnt_report_df['Legacy Recap | Accomplishments'])

final_qnt_report_df['Challenges'] = final_qnt_report_df['Admin Recap | Challenges'].fillna(
                                            final_qnt_report_df['Legacy Recap | Challenges'])

final_qnt_report_df['Lessons Learned'] = final_qnt_report_df['Admin Recap | Lessons Learned'].fillna(
                                            final_qnt_report_df['Legacy Recap | Lessons Learned'])

final_qnt_report_df['Barriers to Participation'] = final_qnt_report_df['Admin Recap | Barriers to participation'].fillna(
                                            final_qnt_report_df['Legacy Recap | Barriers to Participation'])

final_qnt_report_df['Best-resonating message'] = final_qnt_report_df['Admin Recap | Best-resonating message'].fillna(
                                            final_qnt_report_df['Legacy Recap | Best-resonating Messages'])

final_qnt_report_df['Concerns expressed'] = final_qnt_report_df['Admin Recap | Concerns expressed'].fillna(
                                            final_qnt_report_df['Legacy Recap | Concerns Expressed'])

final_qnt_report_df['Other notes'] = final_qnt_report_df['Admin Recap | Other notes'].fillna(
                                            final_qnt_report_df['Legacy Recap | Other Notes'])

final_qnt_report_df['Governance Activities Summary'] = final_qnt_report_df['Admin Recap | Governance Activities Summary'].fillna(
                                            final_qnt_report_df['Legacy Recap | Governance Activities Summary'])

# Null out Legacy Id for NN2.0 activities
final_qnt_report_df.loc[(final_qnt_report_df['Legacy Id'] == 0), 'Legacy Id'] = None

# Drop duplicates caused by ECs submitting multiple recaps
final_qnt_report_df = final_qnt_report_df.drop_duplicates(subset=['Event ID'])

# Copy the full quant report to Target Markets
final_qnt_report_df.to_csv(r'target_market/build_tm_inputs/nn2_apr1_2022.csv', index=False)

final_qnt_report_df = final_qnt_report_df[[
    'TYPE',
    'Date',
    'Name',
    'Event ID',
    'Legacy Id',
    'Event Status',
    'Recap Type Submitted',
    'Engagement Activity Type',
    'Partner',
    'Partners Recapped',
    'Collaboration Partner(s)',
    # 'NN2 Venue Market',
    # 'NN2 Consortia Partner Target Market',
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
    '#-Total Material Distributed',
    'Target Population',
    'Accomplishments',
    'Challenges',
    'Lessons Learned',
    'Barriers to Participation',
    'Best-resonating message',
    'Concerns expressed',
    'Governance Activities Summary',
    'Other notes',
    'Research Facing',
    # 'CTA',
    # 'CTA Text',
    ]].copy()

final_qnt_report_df = final_qnt_report_df.sort_values(["TYPE"], ascending = (False))

# This is the final quant report
with pd.ExcelWriter(r'data/FINAL_Pyxis_Quantitative - MONTH 2022.xlsx') as writer:  # doctest: +SKIP
    final_qnt_report_df.to_excel(writer, index=False, startrow = 0, sheet_name = 'Quantitative MMM 1-31 2022')

    workbook  = writer.book

    tab_1_sheet_name = 'Quantitative MMM 1-31 2022'
    worksheet_1 = writer.sheets[tab_1_sheet_name]

    header_format = workbook.add_format({'bold': False, 'text_wrap': False, 'fg_color': '#4F81BD', 'color': '#FFFFFF', 'border': 0, 'align': 'center', 'valign': 'center'})
    for col_num, value in enumerate(final_qnt_report_df.columns.values):
        worksheet_1.write(0, col_num, value, header_format)





# final_qnt_report_df.to_excel(r'tmp/debug.xlsx', index=False)