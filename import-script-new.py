# %%
import time
import pandas as pd
import numpy as np
import pytz
from datetime import datetime

# timer
t0 = time.time()

# %%
# For Australia date
australia_timezone = pytz.timezone('Australia/Sydney')
today = datetime.now(australia_timezone).date()

# Basic constants 
USER = ''
YEAR = 2023

# Funding bands
BAND_A = 2500
BAND_B = 1250

# Is exchange?
EXCHANGE = True

# For credit?
CREDITS = True
CREDIT_POINTS = 24

# Gov grants
GOV_GRANT = True
GOV_GRANT_AMT = 7000
GOV_GRANT_TYPE = 'NCP Mobility Program Grant'

# %%
file    = pd.ExcelFile('')
sheet_1 = pd.read_excel(file, 'Outgoing - students')
sheet_2 = pd.read_excel(file, 'MAP Report')

# %%
# Filter out rows with incorrect course codes
incorrect_course_codes = {'A0501','A0502',9413,3050}
mask = sheet_2['Course 1 Code'].isin(incorrect_course_codes)
sheet_2 = sheet_2[sheet_2['Monash ID number'].notnull()]

# Replace incorrect course codes with Course 2 values in filtered rows
sheet_2.loc[mask, 'Course 1 Code'] = sheet_2.loc[mask, 'Course2 Code']
sheet_2.loc[mask, 'Course 1 Title'] = sheet_2.loc[mask, 'Course 2 Title']
sheet_2.loc[mask, 'Course 1 Campus'] = sheet_2.loc[mask, 'Course 2 Campus']
sheet_2.loc[mask, 'Course 1 Managing Faculty'] = sheet_2.loc[mask, 'Course 2 Managing Faculty']
sheet_2.loc[mask, 'Course 1 Second Faculty'] = sheet_2.loc[mask, 'Course 2 Second Faculty']

# Remove unwanted status values
# sheet_2 = sheet_2[(sheet_2['Status'] != 'Unsuccessful') & (sheet_2['Status'] != 'Pending') & (sheet_2['Status'] != 'Submitted')]
sheet_2 = sheet_2[(sheet_2['Status'] != 'Unsuccessful') & (sheet_2['Status'] != 'Pending')]
sheet_1 = sheet_1[sheet_1.Campus != 'Malaysia']

# Capitalise proper names and update pre-departure field
sheet_2[['Last Name', 'First Name', 'Course 1 Campus']] = sheet_2[['Last Name', 'First Name', 'Course 1 Campus']].applymap(str.title)
sheet_2['Pre-departure Modules Completion'] = sheet_2['Pre-departure Modules Completion'].notna()

# Mark domestic or international students
sheet_2['INTERNATIONAL_STUDENT'] = sheet_2['INTERNATIONAL_STUDENT'].replace({'International': True, 'Domestic': False})

# %%
# Get values from sheet_2
monash_id = sheet_2['Monash ID number'].values
last_name = sheet_2['Last Name'].values
first_name = sheet_2['First Name'].values
dob = sheet_2['DOB'].values
gender = sheet_2['Gender'].values
email = sheet_2['Email'].values
course_1_faculty = sheet_2['Course 1 Managing Faculty'].values
course_1_2nd_faculty = sheet_2['Course 1 Second Faculty'].values
course_1_title = sheet_2['Course 1 Title'].values
course_1_code = sheet_2['Course 1 Code'].values
course_1_campus = sheet_2['Course 1 Campus'].values
level_of_course = sheet_2['Level of Course'].values
international_student = sheet_2['INTERNATIONAL_STUDENT'].values
prog_title = sheet_2['Program'].values
prog_start_date = sheet_2['Program Date Record: Start Date'].values
prog_end_date = sheet_2['Program Date Record: End Date'].values
prog_country = sheet_2['Program Currently Assigned Country'].values
status = sheet_2['Status'].values
pre_departure = sheet_2['Pre-departure Modules Completion'].values
term = sheet_2['Term'].values

# %%
# Set sheet 1 fields
sheet_1['Student ID'] = monash_id
sheet_1['Family'] = last_name
sheet_1['Given Name'] = first_name
sheet_1['Date of Birth'] = dob
sheet_1['Gender'] = gender
sheet_1['Email'] = email
sheet_1['Faculty Owner'] = course_1_faculty
sheet_1['Second Faculty'] = course_1_2nd_faculty
sheet_1['Course Title'] = course_1_title
sheet_1['Course Code'] = course_1_code
sheet_1['Campus'] = course_1_campus
sheet_1['Undergrad/Postgrad'] = level_of_course
sheet_1['International Student?'] = international_student
sheet_1['Direction'] = 'Outgoing'
sheet_1['Year abroad'] = YEAR
sheet_1['Program Title'] = prog_title
sheet_1['DateProgStart'] = prog_start_date
sheet_1['DateProgEnd'] = prog_end_date
sheet_1['Country of program'] = prog_country
sheet_1['AcceptedByMA'] = True

# Set credit points and semester studies
if CREDITS:
    if EXCHANGE:
        sheet_1['Credit points attaining'] = term
        sheet_1['Number of sems abroad'] = term
        sheet_1['Semester/s abroad'] = term
        sheet_1['Number of sems abroad'] = sheet_1['Credit points attaining'].replace({
            'Sem 2 & Sem 1': 2, 'Sem 1 & Sem 2': 2, 'Semester 1': 1, 'Semester 2': 1
        })
        sheet_1['Semester/s abroad'] = sheet_1['Credit points attaining'].replace({
            'Sem 2 & Sem 1': 2, 'Sem 1 & Sem 2': 1, 'Semester 1': 1, 'Semester 2': 2
        })
        sheet_1['Credit points attaining'] = sheet_1['Credit points attaining'].replace({
            'Sem 2 & Sem 1': 48, 'Sem 1 & Sem 2': 48, 'Semester 1': 24, 'Semester 2': 24
        })
        sheet_1['Semester Studies'] = True
        sheet_1['University'] = sheet_1['Program Title'].str.split('|', expand=True)[1].str.strip()
        sheet_1['Program'] = sheet_1['Program Title'].str.split('|', expand=True)[0].replace({
            'EXC ': 'Exchange', 'ISA ': 'Study Abroad'
        })
        sheet_1['Program Title'] = sheet_1['Program Title'].str.split('|', expand=True)[0].replace({
            'EXC ': 'Semester Exchange', 'ISA ': 'Independent Study Abroad'
        })
    else:
        sheet_1['Credit points attaining'] = CREDIT_POINTS
        sheet_1['Semester Studies'] = False
        try:
            sheet_1['Program Title'] = sheet_1['Program Title'].str.split('|', expand=True)[1].str.strip()
        except KeyError:
            pass
else:
    sheet_1['Credit points attaining'] = 'Not for Credit'
    sheet_1['Semester Studies'] = False
    sheet_1['MAGrantIneligible'] = True
    sheet_1['Program Title'] = sheet_1['Program Title'].str.split('|', expand=True)[1].str.strip()

sheet_1['A&CRcvd'] = True
sheet_1['MAP Status'] = status
sheet_1['AlertTraveler Activated'] = False
sheet_1['Enrolled in Overseas Study Unit(s)'] = False
sheet_1['Pre-departure Training'] = pre_departure
sheet_1['Initially Entered By'] = USER
sheet_1['Date Entered'] = today

# %%
enr = pd.read_excel('')
ovs = pd.read_excel('')

# %%
# Define the names of the unit columns in the dataframe
unit_cols = [
    'UNIT_1', 'UNIT_2', 'UNIT_3', 'UNIT_4', 'UNIT_5', 'UNIT_6', 'UNIT_7',
    'UNIT_8', 'UNIT_9', 'UNIT_10', 'UNIT_11', 'UNIT_12', 'UNIT_13',
    'UNIT_14', 'UNIT_15', 'UNIT_16', 'UNIT_17', 'UNIT_18', 'UNIT_19',
    'UNIT_20'
]

# Split each value in the unit columns by the first space character, and keep only the first part
enr[unit_cols] = enr[unit_cols].astype(str).apply(lambda x: x.str.split(" ").str[0])

# Create a mask that indicates which rows have at least one unit code that matches a code in the `ovs` dataframe
mask = pd.concat([enr[col].isin(ovs['UNIT_CD'].values) for col in unit_cols], axis=1)
mask = mask.astype(int)

# Replace the unit codes with the corresponding mask values (0 or 1)
enr[unit_cols] = mask

# Count the number of units each student is enrolled in
enr['enrolled'] = enr[unit_cols].sum(axis=1)

# Keep only the columns 'PERSON_ID' and 'enrolled'
enr = enr[['PERSON_ID','enrolled']]
enrolled = enr.loc[enr['enrolled'] > 0, 'PERSON_ID']
sheet_1.loc[sheet_1['Student ID'].isin(enrolled), 'Enrolled in Overseas Study Unit(s)'] = True

# %%
if EXCHANGE:
    if GOV_GRANT:
        ncp = sheet_1[['Student ID', 'Country of program', 'Date of Birth', 'DateProgStart', 'Undergrad/Postgrad', 'International Student?']].copy()
        
        host_locations = {'Bangladesh', 'Bhutan', 'Brunei Darussalam', 'Cambodia', 'China', 'Cook Islands', 
                    'Federated States of Micronesia', 'Fiji', 'French Polynesia', 'Hong Kong', 'India', 
                    'Indonesia', 'Japan', 'Kiribati', 'Laos', 'Malaysia', 'Maldives', 'Marshall Islands', 
                    'Mongolia', 'Myanmar', 'Nauru', 'Nepal', 'New Caledonia', 'Niue', 'Pakistan', 'Palau', 
                    'Papua New Guinea', 'Philippines', 'Republic of Korea', 'Samoa', 'Singapore', 'Solomon Islands', 
                    'Sri Lanka', 'Taiwan', 'Thailand', 'Timor-Leste', 'Tonga', 'Tuvalu', 'Vanuatu', 'Vietnam'}

        # convert 'Date of Birth' column to datetime format
        ncp['Date of Birth'] = pd.to_datetime(ncp['Date of Birth'])

        # calculate age in years and create a new column to store the result
        ncp['Age'] = (pd.to_datetime(ncp['DateProgStart']) - ncp['Date of Birth']) / np.timedelta64(1, 'Y')

        ncp['Eligibility'] = np.where((ncp['Age'] >= 18) & (ncp['Age'] <= 28)
                                        & (ncp['International Student?'] == False) 
                                        & (ncp['Undergrad/Postgrad'] == 'Undergraduate') 
                                        & (ncp['Country of program'].isin(host_locations)), 'Eligible', 'Not eligible')

        # merge ncp dataframe with sheet_1 on 'Student ID'
        sheet_1 = pd.merge(sheet_1, ncp[['Student ID', 'Eligibility']], on='Student ID', how='left')

        # update based on 'Eligibility' column
        sheet_1.loc[sheet_1['Eligibility'] == 'Eligible', 'GovGrantType'] = GOV_GRANT_TYPE
        sheet_1.loc[sheet_1['Eligibility'] == 'Eligible', 'GovGrantAmt'] = GOV_GRANT_AMT
        
        sheet_1.loc[sheet_1['Eligibility'] == 'Eligible', 'MAGrantIneligible'] = True
        sheet_1.loc[sheet_1['Eligibility'] == 'Not eligible', 'MAGrantIneligible'] = False
        
        # Create a boolean mask for the universities that are eligible for Band A
        band_a_universities = {'University of Padua', 'University of Warwick', 'Monash University Malaysia'}
        univ_mask = sheet_1['University'].isin(band_a_universities)

        # Assign grant amounts based on eligibility
        sheet_1['MAGrantAmt'] = np.where(univ_mask, 'BAND_A', 'BAND_B')

        # Create a boolean mask for students who are not eligible
        not_eligible_mask = sheet_1['Eligibility'] == 'Not eligible'

        # Select only the rows where the mask is True and assign grant amounts based on eligibility
        sheet_1.loc[not_eligible_mask, 'MAGrantAmt'] = np.where(sheet_1.loc[not_eligible_mask, 'University'].isin(band_a_universities), BAND_A, BAND_B)
        
        # independent study abroad and exchange mask
        study_abroad = {'Independent Study Abroad'}
        study_abroad_mask = sheet_1['Program Title'].isin(study_abroad)
        sheet_1['MAGrantIneligible'] = np.where(study_abroad_mask, True, sheet_1['MAGrantIneligible'])
        sheet_1['MAGrantAmt'] = np.where(study_abroad_mask, None, sheet_1['MAGrantAmt'])
            
        sheet_1.drop(columns=['Eligibility'],inplace=True)
        
if GOV_GRANT:
    ncp = sheet_1[['Student ID', 'Country of program', 'Date of Birth', 'DateProgStart', 'Undergrad/Postgrad', 'International Student?']].copy()
    
    host_locations = {'Bangladesh', 'Bhutan', 'Brunei Darussalam', 'Cambodia', 'China', 'Cook Islands', 
                  'Federated States of Micronesia', 'Fiji', 'French Polynesia', 'Hong Kong SAR', 'India', 
                  'Indonesia', 'Japan', 'Kiribati', 'Laos', 'Malaysia', 'Maldives', 'Marshall Islands', 
                  'Mongolia', 'Myanmar', 'Nauru', 'Nepal', 'New Caledonia', 'Niue', 'Pakistan', 'Palau', 
                  'Papua New Guinea', 'Philippines', 'Republic of Korea', 'Samoa', 'Singapore', 'Solomon Islands', 
                  'Sri Lanka', 'Taiwan', 'Thailand', 'Timor-Leste', 'Tonga', 'Tuvalu', 'Vanuatu', 'Vietnam'}

    # convert 'Date of Birth' column to datetime format
    ncp['Date of Birth'] = pd.to_datetime(ncp['Date of Birth'])

    # calculate age in years and create a new column to store the result
    ncp['Age'] = (pd.to_datetime(today) - ncp['Date of Birth']) / np.timedelta64(1, 'Y')

    ncp['Eligibility'] = np.where(((ncp['Age'] >= 18) | (ncp['Age'] <= 28)) 
                                    & (ncp['International Student?'] == False) 
                                    & (ncp['Undergrad/Postgrad'] == 'Undergraduate') 
                                    & (ncp['Country of program'].isin(host_locations)), 'Eligible', 'Not eligible')


    # merge ncp dataframe with sheet_1 on 'Student ID'
    sheet_1 = pd.merge(sheet_1, ncp[['Student ID', 'Eligibility']], on='Student ID', how='left')

    # update based on 'Eligibility' column
    sheet_1.loc[sheet_1['Eligibility'] == 'Eligible', 'GovGrantType'] = GOV_GRANT_TYPE
    sheet_1.loc[sheet_1['Eligibility'] == 'Eligible', 'GovGrantAmt'] = GOV_GRANT_AMT
    sheet_1.loc[sheet_1['Eligibility'] == 'Eligible', 'MAGrantIneligible'] = True
    sheet_1.drop(columns=['Eligibility'],inplace=True)
    
if CREDITS and (EXCHANGE is False):
    not_eligible_mask = sheet_1['GovGrantType'] == 'NCP Mobility Program Grant'
    sheet_1['MAGrantAmt'] = np.where(not_eligible_mask, None, BAND_B)
    sheet_1['MAGrantIneligible'] = np.where(not_eligible_mask, True, False)

# %%
sheet_1.to_excel(f'output.xlsx',index=False)

# %%
t1 = time.time()
total = t1-t0
print(f"Import script took {total} seconds")


