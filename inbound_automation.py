# ====================== Packages, Imports ============================#

import pandas as pd
import smartsheet

import warnings
warnings.filterwarnings("ignore") 

# ====================== Cleaning excel section ============================#

FINAL_PATH = ""
MRS_PATH = ""

# Load dataframes from excel files
final_units = pd.read_excel(FINAL_PATH)
MRS_units   = pd.read_excel(MRS_PATH)

# Function to drop unwanted locations from dataframe
def drop_unwanted_locations(dataframe, location_col):
    locations = ['PENINSULA', 'CLAYTON', 'CAULFIELD', 'MALAYSIA', 
                 'CITY', 'ALFRED', 'MMC', 'MEL-LAWCHM', 
                 'MMS-ALFRED', 'PARKVILLE', 'BERWICK', 'GIPPSLAND']
    dataframe = dataframe[dataframe[location_col].isin(locations)]
    return dataframe

# Apply function to both dataframes
final_units = drop_unwanted_locations(final_units, 'Location__c')
MRS_units = drop_unwanted_locations(MRS_units, 'LOCATION_CD')

# Keep only offered and active units
MRS_units = MRS_units[MRS_units["OFFERED_IND"] == 'Y']
MRS_units = MRS_units[MRS_units["UNIT_STATUS"] == 'ACTIVE']

# Split for Semester 1
semester_one_final_units = final_units[final_units['Calendar_Type__c'] == 'S1-01']
semester_one_MRS_units = MRS_units[MRS_units['CAL_TYPE'] == 'S1-01']

# Split for Semester 2
semester_two_final_units = final_units[final_units['Calendar_Type__c'] == 'S2-01']
semester_two_MRS_units = MRS_units[MRS_units['CAL_TYPE'] == 'S2-01']

# List of Trimesters
tri = ['T2-58', 'T3-58', 'T2-57','T3-57', 'T4-57','T1-57','T1-58']

# Filter final_units and MRS_units dataframes for trimesters
tri_final_units = final_units[final_units['Calendar_Type__c'].isin(tri)]
tri_MRS_units = MRS_units[MRS_units['CAL_TYPE'].isin(tri)]

def algo_finding_offered(dataframe_final, dataframe_MRS):    
    # Select columns from dataframe_final
    match_final = dataframe_final[['Academic_Year__c',	'Calendar_Type__c',	'Calendar_Year__c','Location__c','Managing_Faculty_Name__c','Unit_Code__c','TITLE__C','Unit_Class__c','UO_Unique_Key__c','MA STATUS','OFFERED','HANDBOOK LINKS']]
    
    # Create a new column 'combined_fields' by combining 'Location__c' and 'Unit_Code__c'
    match_final['combined_fields'] = match_final[['Location__c', 'Unit_Code__c']].agg('-'.join, axis=1)

    # Select columns from dataframe_MRS
    match_mrs = dataframe_MRS[['LOCATION_CD','UNIT_CD','OFFERED_IND']]
    
    # Create a new column 'combined_fields' by combining 'LOCATION_CD' and 'UNIT_CD'
    match_mrs['combined_fields'] = match_mrs[['LOCATION_CD', 'UNIT_CD']].agg('-'.join, axis=1)

    # Sort the match_mrs dataframe by 'combined_fields'
    match_mrs.sort_values("combined_fields", inplace=True)
    
    # Drop duplicate values in match_mrs dataframe
    match_mrs.drop_duplicates(subset="combined_fields", keep=False, inplace=True)

    # Create an empty list 'offered'
    offered = []
    
    # Loop through 'combined_fields' in match_final dataframe
    for i in match_final['combined_fields']:
        # If 'combined_fields' value is in the list of 'combined_fields' from match_mrs dataframe, append 'Y' to the 'offered' list
        if i in list(match_mrs['combined_fields']):
            offered.append('Y')
        else:
            # If not, append 'N' to the 'offered' list
            offered.append('N')
            
    # Add the 'offered' list as a new column 'OFFERED' to the match_final dataframe
    match_final['OFFERED'] = offered
    return match_final

# Call the function algo_finding_offered for semester 1, 2 and Trimesters
S1 = algo_finding_offered(semester_one_final_units, semester_one_MRS_units)
S2 = algo_finding_offered(semester_two_final_units, semester_two_MRS_units)
TRI = algo_finding_offered(tri_final_units, tri_MRS_units)
final_dataframe = pd.concat([S1,S2,TRI])

# Update 'MA STATUS' column based on 'OFFERED' column
final_dataframe.loc[final_dataframe['OFFERED'] == 'N', 'MA STATUS'] = 'GREY'

# Capitalise titles in 'TITLE__C' column
final_dataframe['TITLE__C'] = final_dataframe['TITLE__C'].str.title()

# Map 'MA STATUS' to comments and add a new column 'COMMENT'
status_comments = {
    'GREEN': 'Open for enrolment, no supporting documents required.',
    'RED': 'Not available until faculty approval has been given. Documentation is required.',
    'YELLOW': 'Monash Abroad is required to check. Documentation may be required for a language unit.'
}
final_dataframe['COMMENT'] = final_dataframe['MA STATUS'].map(status_comments)

# Fill missing values in 'COMMENT' with 'N/A'
final_dataframe['COMMENT'].fillna('N/A', inplace=True)

# Sort the DataFrame by 'Unit_Code__c' column in ascending order
final_dataframe.sort_values('Unit_Code__c', inplace=True, ascending=True)

# Save the DataFrame to an Excel spreadsheet file
with pd.ExcelWriter('Finalised_inbound.xlsx', engine='xlsxwriter') as file:
    final_dataframe.to_excel(file, sheet_name='Offered', index=False)

# ====================== Smartsheet API section (1) ============================#

# API token for accessing Smartsheet API
API_TOKEN = ""

# File path for the excel file to be attached
FILE_NAME = ""

# Defining the sheet ID for which attachments need to be added
SHEET_ID = ""

def update_smartsheet_attachment(SHEET_ID, API_TOKEN, FILE_NAME):
    """
    This function updates an attachment in a Smartsheet sheet with a new version.

    Parameters
    ----------
    SHEET_ID : int
        The ID of the sheet in Smartsheet that needs to be updated with a new attachment.
    API_TOKEN : str
        The API token required to access the Smartsheet API.
    FILE_NAME : str
        The path of the file to be attached as a new version.

    Returns
    -------
    None
        This function only prints a success message if the attachment is updated successfully.

    Examples
    --------
    >>> update_smartsheet_attachment(1234567890, "api_token", "Finalised_inbound.xlsx")
    Attachment updated successfully.
    """
    # Instantiating Smartsheet API object
    smart = smartsheet.Smartsheet(API_TOKEN)

    # Fetching all attachments of the specified sheet
    response = smart.Attachments.list_all_attachments(SHEET_ID, include_all=True)
    attachments = response.data

    # Extracting the ID of the attachment
    for i in attachments:
        attachment_id = f"{i.id}"
    print(f"Attachment ID: {attachment_id}")

    # Updating the attachment with a new version
    response = smart.Attachments.attach_new_version(
        SHEET_ID,
        attachment_id,
        (FILE_NAME, open(FILE_NAME, 'rb'))
    )

    # Printing success message
    print('Attachment updated successfully.')

update_smartsheet_attachment(SHEET_ID, API_TOKEN, FILE_NAME)

# ====================== Smartsheet API section (2) ============================#

# Constants
FINALISED_PATH = ""
API_TOKEN = ""

# Define the status values to be filtered
STATUS_VALUES = ['GREEN', 'RED', 'YELLOW']

# Define the sheet IDs for each status value
SHEET_IDS = [1234567890, 1234567890, 1234567890] 

# Connect to Smartsheet
smart = smartsheet.Smartsheet(API_TOKEN)

def update_smartsheet_status_attachments(STATUS_VALUES, SHEET_IDS, API_TOKEN, FINALISED_PATH):
    """
    This function updates attachments in Smartsheet sheets with filtered versions of a finalised data file.

    Parameters
    ----------
    STATUS_VALUES : list of str
        The list of status values that are used to filter the data.
    SHEET_IDS : list of int
        The list of sheet IDs that correspond to the status values.
    API_TOKEN : str
        The API token required to access the Smartsheet API.
    FINALISED_PATH : str
        The path of the finalised data file.

    Returns
    -------
    None
        This function only prints a success message for each status after the corresponding attachment is updated.

    Examples
    --------
    >>> update_smartsheet_status_attachments(["GREEN", "RED", "YELLOW"], [1234567890, 1234567890, 1234567890], "api_token", "finalised_data.xlsx")
    Process completed for GREEN...
    Process completed for RED...
    Process completed for YELLOW...
    """
    # Load the finalised data into a Pandas dataframe
    finalised = pd.read_excel(FINALISED_PATH)

    # Connect to the Smartsheet API using the API token
    smart = smartsheet.Smartsheet(API_TOKEN)
    
    # Loop through each status value
    for status, sheet_id in zip(STATUS_VALUES, SHEET_IDS):
        # Filter the finalised data for the current status
        filtered = finalised[finalised['status'] == status]

        # Save the filtered data to a temporary file
        filtered_path = f"{status}_finalised.xlsx"
        filtered.to_excel(filtered_path, index=False)

        # Get the current attachments for the sheet
        response = smart.Attachments.list_all_attachments(sheet_id, include_all=True)
        attachments = response.data

    # Check if there are any attachments for the sheet
    if attachments:
        # Get the attachment ID of the most recent attachment
        attachment_id = attachments[-1].id

        # Update the attachment with the filtered data file
        response = smart.Attachments.attach_new_version(
            sheet_id,
            attachment_id,
            ('Finalised_inbound.xlsx', open(filtered_path, 'rb'))
        )

        # Store the result of the attachment update
        updated_attachment = response.result

        # Print a success message
        print(f"Process completed for {status}...")

    else:
        # If there are no attachments for the sheet, create a new attachment with the filtered data file
        response = smart.Attachments.attach(
            sheet_id,
            ('Finalised_inbound.xlsx', open(filtered_path, 'rb'))
        )

        # Print a success message
        print(f"Process completed for {status}...")


update_smartsheet_status_attachments(STATUS_VALUES, SHEET_IDS, API_TOKEN, FINALISED_PATH)

