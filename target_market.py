import pandas as pd
import xlrd
import glob
import numpy as np
from datetime import datetime

nn1_full_df = pd.read_csv(r'target_market/legacy_data.csv', thousands=r',')
nn2_full_df = pd.read_csv(r'target_market/new_data.csv', thousands=r',')

'''
1. Add all the required fields in System_1 to line up with System_2
'''
nn1_full_df['#-Total Social Media Posts'] = (
                    nn1_full_df['Number of videos']
                    + nn1_full_df['Number of tweets ']
                    + nn1_full_df['Number of posts']
                    + nn1_full_df['Number of posts.1']
    )

nn1_full_df['#-Total Social Media Impressions'] = (
                    nn1_full_df['Total Impressions']
                    + nn1_full_df['Impressions']
    )

nn1_full_df['#-Total Social Media Engagement'] = (
                    nn1_full_df['Engagement']
                    + nn1_full_df['Total Engagement']
                    + nn1_full_df['Engagement.1']
                    + nn1_full_df['Engagements ']
    )

nn1_full_df['#-Total Social Media Reach'] = (
                    nn1_full_df['YT Reach']
                    + nn1_full_df['Facebook Reach']
                    + nn1_full_df['Instagram Reach']
    )

nn1_full_df['#-Total Digital Attendees'] = 0

nn1_full_df['#-Total Digital Engagement'] = 0

nn1_full_df['#-Total Attendees'] = (
                    nn1_full_df['Estimated number of attendees at the event:']
    )

nn1_full_df['#-Total Engagement'] = (
                    nn1_full_df[[
                                'Estimated number of attendees you have engaged with at the event:',
                                'Estimated number of attendees engaged with at the event'
                                ]].max(axis=1)

    )

nn1_full_df['#-Total Reach'] = 0

nn1_full_df['#-Total Page Views'] = nn1_full_df['Page views ']

nn1_full_df['#-Total Material Distributed'] = nn1_full_df['Expected audience reach.1']

nn1_full_df['Partners Recapped'] = nn1_full_df['Organization Name']
nn1_full_df['Collaboration Partner(s)'] = ''
nn1_full_df['NN2 Venue Market'] = ''
nn1_full_df['NN2 Consortia Partner Target Market'] = nn1_full_df['TM State']

nn1_full_df['Recap Type Submitted'] = 'Legacy'
nn1_full_df['Accomplishments'] = nn1_full_df['In regards to this project, please list any accomplishments experienced by your organizationthis month.']
nn1_full_df['Challenges'] = nn1_full_df['In regards to this project, please list any challenges experienced by your organization this month. How were these challenges addressed?']
nn1_full_df['Lessons Learned'] = nn1_full_df['In regards to this project, please list any lessons learned by your organization this month.']
nn1_full_df['Barriers to Participation'] = nn1_full_df['Barriers to Participation ']
nn1_full_df['Best-resonating message'] = nn1_full_df['Best-resonating Messages']
nn1_full_df['Concerns expressed'] = nn1_full_df['Concerns Expressed']
nn1_full_df['Other notes'] = nn1_full_df['Other Notes']
nn1_full_df['Governance Activities Summary'] = nn1_full_df['Please provide a one paragraph summary of the governance activities you participated in this month. (e.g. sub awardee management, monthly meetings, EC meetings, internal training, internal organizational management)']
nn1_full_df['Research Facing'] = ''
nn1_full_df['CTA'] = ''
nn1_full_df['CTA Text'] = ''
nn1_full_df['Source'] = 'NN1'

'''
2. Define the target market NN2 dataframe
'''
nn2_full_df['TM City'] = 'Legacy Field'
nn2_full_df['TM State'] = 'Legacy Field'
nn2_full_df['Source'] = 'NN2'

nn2_df = nn2_full_df[[
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
    'NN2 Venue Market',
    'NN2 Consortia Partner Target Market',
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
    'CTA',
    'CTA Text',
    'City',
    'State',
    'TM City',
    'TM State',
    'Source'
    ]].copy()

'''
3. Define the target market NN1 dataframe
'''
nn1_df = nn1_full_df[[
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
    'NN2 Venue Market',
    'NN2 Consortia Partner Target Market',
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
    'CTA',
    'CTA Text',
    'City',
    'State',
    'TM City',
    'TM State',
    'Source'
    ]].copy()

'''
4. Create the target market dataframe
'''
# Merge the NN2 dataframe with NN1
# Produces a full outer join dataframes that need to be further split
nn1_nn2_outer_df = nn2_df.merge(nn1_df.drop_duplicates(),
    on=['Legacy Id'], how='outer', indicator=False)

# Split the full outer join for NN1
nn1_tmp_df = nn1_nn2_outer_df[nn1_nn2_outer_df['Source_x'].isnull()]

nn1_tmp_df = nn1_tmp_df[[
    'TYPE_y',
    'Date_y',
    'Name_y',
    'Event ID_y',
    'Legacy Id',
    'Event Status_y',
    'Recap Type Submitted_y',
    'Engagement Activity Type_y',
    'Partner_y',
    'Partners Recapped_y',
    'Collaboration Partner(s)_y',
    'NN2 Venue Market_y',
    'NN2 Consortia Partner Target Market_y',
    '#-Total Social Media Posts_y',
    '#-Total Social Media Impressions_y',
    '#-Total Social Media Engagement_y',
    '#-Total Social Media Reach_y',
    '#-Total Digital Attendees_y',
    '#-Total Digital Engagement_y',
    '#-Total Attendees_y',
    '#-Total Engagement_y',
    '#-Total Reach_y',
    '#-Total Page Views_y',
    '#-Total Material Distributed_y',
    'Target Population_y',
    'Accomplishments_y',
    'Challenges_y',
    'Lessons Learned_y',
    'Barriers to Participation_y',
    'Best-resonating message_y',
    'Concerns expressed_y',
    'Governance Activities Summary_y',
    'Other notes_y',
    'Research Facing_y',
    'CTA_y',
    'CTA Text_y',
    'TM City_y',
    'TM State_y',
    # 'Source_y',
    ]].copy()

# Split the full outer join for NN2
nn2_tmp_df = nn1_nn2_outer_df[nn1_nn2_outer_df['Source_x'] == 'NN2']

# Define where to pull the values from - NN1 or NN2
nn2_tmp_df = nn2_tmp_df.copy()


for idx, row in nn2_tmp_df.iterrows():

    # If the record is a legacy event, then pull metrics for the field from NN1
    if nn2_tmp_df.loc[idx, 'Source_y'] == 'NN1':
        nn2_tmp_df.loc[idx, 'Partner'] = nn2_tmp_df.loc[idx,'Partner_y']
        nn2_tmp_df.loc[idx, 'Partners Recapped'] = nn2_tmp_df.loc[idx,'Partners Recapped_y']
        nn2_tmp_df.loc[idx, 'Target Population'] = nn2_tmp_df.loc[idx,'Target Population_y']
        nn2_tmp_df.loc[idx, 'City'] = nn2_tmp_df.loc[idx,'City_y']
        nn2_tmp_df.loc[idx, 'State'] = nn2_tmp_df.loc[idx,'State_y']
        nn2_tmp_df.loc[idx, 'TM City'] = nn2_tmp_df.loc[idx,'TM City_y']
        nn2_tmp_df.loc[idx, 'TM State'] = nn2_tmp_df.loc[idx,'TM State_y']
        nn2_tmp_df.loc[idx, 'Source'] = nn2_tmp_df.loc[idx,'Source_y']
        nn2_tmp_df.loc[idx, 'Recap Type Submitted'] = nn2_tmp_df.loc[idx,'Recap Type Submitted_y']

        # NN1 Quant Columns
        nn2_tmp_df.loc[idx, '#-Total Social Media Posts'] = nn2_tmp_df.loc[idx,'#-Total Social Media Posts_y']
        nn2_tmp_df.loc[idx, '#-Total Social Media Impressions'] = nn2_tmp_df.loc[idx,'#-Total Social Media Impressions_y']
        nn2_tmp_df.loc[idx, '#-Total Social Media Engagement'] = nn2_tmp_df.loc[idx,'#-Total Social Media Engagement_y']
        nn2_tmp_df.loc[idx, '#-Total Social Media Reach'] = nn2_tmp_df.loc[idx,'#-Total Social Media Reach_y']
        nn2_tmp_df.loc[idx, '#-Total Digital Attendees'] = nn2_tmp_df.loc[idx,'#-Total Digital Attendees_y']
        nn2_tmp_df.loc[idx, '#-Total Digital Engagement'] = nn2_tmp_df.loc[idx,'#-Total Digital Engagement_y']
        nn2_tmp_df.loc[idx, '#-Total Attendees'] = nn2_tmp_df.loc[idx,'#-Total Attendees_y']
        nn2_tmp_df.loc[idx, '#-Total Engagement'] = nn2_tmp_df.loc[idx,'#-Total Engagement_y']
        nn2_tmp_df.loc[idx, '#-Total Reach'] = nn2_tmp_df.loc[idx,'#-Total Reach_y']
        nn2_tmp_df.loc[idx, '#-Total Page Views'] = nn2_tmp_df.loc[idx,'#-Total Page Views_y']
        nn2_tmp_df.loc[idx, '#-Total Material Distributed'] = nn2_tmp_df.loc[idx,'#-Total Material Distributed_y']

    # Otherwise pull the field from NN2 
    else:
        nn2_tmp_df.loc[idx, 'Partner'] = nn2_tmp_df.loc[idx,'Partner_x']
        nn2_tmp_df.loc[idx, 'Partners Recapped'] = nn2_tmp_df.loc[idx,'Partners Recapped_x']
        nn2_tmp_df.loc[idx, 'Target Population'] = nn2_tmp_df.loc[idx,'Target Population_x']
        nn2_tmp_df.loc[idx, 'City'] = nn2_tmp_df.loc[idx,'City_x']
        nn2_tmp_df.loc[idx, 'State'] = nn2_tmp_df.loc[idx,'State_x']
        nn2_tmp_df.loc[idx, 'TM City'] = ''
        nn2_tmp_df.loc[idx, 'TM State'] = ''
        nn2_tmp_df.loc[idx, 'Source'] = nn2_tmp_df.loc[idx,'Source_x']
        nn2_tmp_df.loc[idx, 'Recap Type Submitted'] = nn2_tmp_df.loc[idx,'Recap Type Submitted_x']

        # NN2 Quant Columns
        nn2_tmp_df.loc[idx, '#-Total Social Media Posts'] = nn2_tmp_df.loc[idx,'#-Total Social Media Posts_x']
        nn2_tmp_df.loc[idx, '#-Total Social Media Impressions'] = nn2_tmp_df.loc[idx,'#-Total Social Media Impressions_x']
        nn2_tmp_df.loc[idx, '#-Total Social Media Engagement'] = nn2_tmp_df.loc[idx,'#-Total Social Media Engagement_x']
        nn2_tmp_df.loc[idx, '#-Total Social Media Reach'] = nn2_tmp_df.loc[idx,'#-Total Social Media Reach_x']
        nn2_tmp_df.loc[idx, '#-Total Digital Attendees'] = nn2_tmp_df.loc[idx,'#-Total Digital Attendees_x']
        nn2_tmp_df.loc[idx, '#-Total Digital Engagement'] = nn2_tmp_df.loc[idx,'#-Total Digital Engagement_x']
        nn2_tmp_df.loc[idx, '#-Total Attendees'] = nn2_tmp_df.loc[idx,'#-Total Attendees_x']
        nn2_tmp_df.loc[idx, '#-Total Engagement'] = nn2_tmp_df.loc[idx,'#-Total Engagement_x']
        nn2_tmp_df.loc[idx, '#-Total Reach'] = nn2_tmp_df.loc[idx,'#-Total Reach_x']
        nn2_tmp_df.loc[idx, '#-Total Page Views'] = nn2_tmp_df.loc[idx,'#-Total Page Views_x']
        nn2_tmp_df.loc[idx, '#-Total Material Distributed'] = nn2_tmp_df.loc[idx,'#-Total Material Distributed_x']

nn2_tmp_df = nn2_tmp_df[[
    'TYPE_x',
    'Date_x',
    'Name_x',
    'Event ID_x',
    'Legacy Id',
    'Event Status_x',
    'Recap Type Submitted_x',
    'Engagement Activity Type_x',
    'Partner',
    'Partners Recapped',
    'Collaboration Partner(s)_x',
    'NN2 Venue Market_x',
    'NN2 Consortia Partner Target Market_x',
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
    'Accomplishments_x',
    'Challenges_x',
    'Lessons Learned_x',
    'Barriers to Participation_x',
    'Best-resonating message_x',
    'Concerns expressed_x',
    'Governance Activities Summary_x',
    'Other notes_x',
    'Research Facing_x',
    'CTA_x',
    'CTA Text_x',
    'City',
    'State',
    'TM City',
    'TM State',
    # 'Source',
    ]].copy()

# Remove the _x and _y column indicators
nn1_tmp_df.columns = nn1_tmp_df.columns.str.replace('_y', '', regex=True)
nn2_tmp_df.columns = nn2_tmp_df.columns.str.replace('_x', '', regex=True)

nn_frames = [
        nn1_tmp_df,
        nn2_tmp_df,
        ]

target_market_df = pd.concat(nn_frames)

# Strip whitespaces in City, State fields
target_market_df['City'] = target_market_df['City'].str.rstrip()
target_market_df['State'] = target_market_df['State'].str.rstrip()

# Change Partner Monthly Summary Report to MSR
target_market_df.loc[(target_market_df['TYPE'] == 'Partner Monthly Summary Report'), 'TYPE'] = 'MSR'
target_market_df['TYPE'] = target_market_df['TYPE'].str.upper()

# Filter and sort the data set based on the request
target_market_df = target_market_df.sort_values(['Legacy Id'], ascending = (True))

# This is the final quant report
with pd.ExcelWriter(r'data/FINAL_Target_Market_Quant - MONTH 2022.xlsx') as writer:  # doctest: +SKIP
    target_market_df.to_excel(writer, index=False, startrow = 0, sheet_name = 'Target Market MMM 1-31 2022')

    workbook  = writer.book

    tab_1_sheet_name = 'Target Market MMM 1-31 2022'
    worksheet_1 = writer.sheets[tab_1_sheet_name]

    header_format = workbook.add_format({'bold': False, 'text_wrap': False, 'fg_color': '#4F81BD', 'color': '#FFFFFF', 'border': 0, 'align': 'center', 'valign': 'center'})
    for col_num, value in enumerate(target_market_df.columns.values):
        worksheet_1.write(0, col_num, value, header_format)

# Create the database file
database_df = target_market_df.copy()

for idx, row in database_df.iterrows():

    # If Event ID is empty then assign the Legacy Id
    if pd.isna(database_df.loc[idx, 'Event ID']) is True:
        database_df.loc[idx, 'Event ID'] = database_df.loc[idx, 'Legacy Id']

database_df = database_df.rename(columns={
    'TYPE': 'TYPE',
    'Date': 'Date',
    'Name': 'Name',
    'Event ID': 'Event_ID',
    'Legacy Id': 'Legacy_Id',
    'Event Status': 'Event_Status',
    'Recap Type Submitted': 'Recap_Type_Submitted',
    'Engagement Activity Type': 'Engagement_Activity_Type',
    'Partner': 'Partner',
    'Partners Recapped': 'Partners_Recapped',
    'Collaboration Partner(s)': 'Collaboration_Partners',
    'NN2 Venue Market': 'NN2_Venue_Market',
    'NN2 Consortia Partner Target Market': 'NN2_Consortia_Partner_Target_Market',
    '#-Total Social Media Posts': 'Total_Social_Media_Posts',
    '#-Total Social Media Impressions': 'Total_Social_Media_Impressions',
    '#-Total Social Media Engagement': 'Total_Social_Media_Engagement',
    '#-Total Social Media Reach': 'Total_Social_Media_Reach',
    '#-Total Digital Attendees': 'Total_Digital_Attendees',
    '#-Total Digital Engagement': 'Total_Digital_Engagement',
    '#-Total Attendees': 'Total_Attendees',
    '#-Total Engagement': 'Total_Engagement',
    '#-Total Reach': 'Total_Reach',
    '#-Total Page Views': 'Total_Page_Views',
    '#-Total Material Distributed': 'Total_Material_Distributed',
    'Target Population': 'Target_Population',
    'Accomplishments': 'Accomplishments',
    'Challenges': 'Challenges',
    'Lessons Learned': 'Lessons_Learned',
    'Barriers to Participation': 'Barriers_to_Participation',
    'Best-resonating message': 'Best_resonating_message',
    'Concerns expressed': 'Concerns_expressed',
    'Governance Activities Summary': 'Governance_Activities_Summary',
    'Other notes': 'Other_notes',
    'Research Facing': 'Research_Facing',
    'CTA': 'CTA',
    'CTA Text': 'CTA_Text',
    'TM City': 'TM_City',
    'TM State': 'TM_State',
    'City': 'City',
    'State': 'State',
    })

database_df = database_df[[
    'Event_ID',
    'Legacy_Id',
    'TYPE',
    'Date',
    'Name',
    'Event_Status',
    'Recap_Type_Submitted',
    'Engagement_Activity_Type',
    'Partner',
    'Partners_Recapped',
    'Collaboration_Partners',
    'NN2_Venue_Market',
    'NN2_Consortia_Partner_Target_Market',
    'Total_Social_Media_Posts',
    'Total_Social_Media_Impressions',
    'Total_Social_Media_Engagement',
    'Total_Social_Media_Reach',
    'Total_Digital_Attendees',
    'Total_Digital_Engagement',
    'Total_Attendees',
    'Total_Engagement',
    'Total_Reach',
    'Total_Page_Views',
    'Total_Material_Distributed',
    'Target_Population',
    'Accomplishments',
    'Challenges',
    'Lessons_Learned',
    'Barriers_to_Participation',
    'Best_resonating_message',
    'Concerns_expressed',
    'Governance_Activities_Summary',
    'Other_notes',
    'Research_Facing',
    'CTA',
    'CTA_Text',
    'TM_City',
    'TM_State',
    'City',
    'State',
    ]].copy()

database_df.to_csv(r'data/db/db_network_load.csv', index=False, header=True)