import openpyxl
import pandas as pd
import numpy as np

def clean_data(file_path: str, output_path: str):
    """
    Reads excel file, drops ineligible courses and calculates remaining CP, and saves to output path.
    
    Parameters:
        file_path (str): Path to input excel file.
        output_path (str): Path to save the output excel file.
    
    Returns:
        None
    """
    file = pd.ExcelFile(file_path)
    sheet_1 = pd.read_excel(file, 'Report')
    sheet_2 = pd.read_excel(file, 'Calculations')
    
    # dropping ineligible courses 
    incorrect_course_codes = ['A0501','A0502','9413','9401']
    sheet_1 = sheet_1[~sheet_1['COURSE_CD'].isin(incorrect_course_codes)]
    
    # remaining cp 
    sheet_1['CP_REMAINING'] = sheet_1['CP_REMAINING'] - sheet_1['CP_CURRENT_ENROLLED']
    
    with pd.ExcelWriter(output_path) as file:
        sheet_1.to_excel(file, 'Report')
        sheet_2.to_excel(file, 'MAP Calculations')

# Load the excel file using openpyxl
wb = openpyxl.load_workbook("OS-HELP - Eligibility check output.xlsx")

# Index and assign worksheets to variables for faster access
main_ws = wb.worksheets[0]
MAP_export = wb.worksheets[1]

# Delete the first column of both worksheets
main_ws.delete_cols(1)
MAP_export.delete_cols(1)

def copy_paste_to_template(copy_column, paste_sheet, paste_column):
    """
    Copies values from `copy_column` to `paste_column` in `paste_sheet`.

    Parameters:
    copy_column (list): The column to copy.
    paste_sheet (openpyxl.worksheet.worksheet.Worksheet): The sheet to paste to.
    paste_column (int): The column number to paste to.

    Returns:
    None
    """
    for cell in copy_column:
        paste_sheet.cell(row = cell[0].row, column = paste_column, value = cell[0].value)
        
# Copy "STUDENT ID" column
copy_paste_to_template(main_ws['A2:A2000'], MAP_export, 1) 

# Copy "CREDIT POINTS COMPLETED" column
copy_paste_to_template(main_ws['BC2:BC2000'], MAP_export, 2) 

# Copy "CREDIT POINTS REMAINING" column
copy_paste_to_template(main_ws['BB2:BB2000'], MAP_export, 3)

# Copy "FEE_CAT" column
copy_paste_to_template(main_ws['Q2:Q2000'], MAP_export, 4)  

# Copy "INTERNATIONAL?" column
copy_paste_to_template(main_ws['AD2:AD2000'], MAP_export, 5)  

# Save the modified workbook
wb.save('OS-HELP - Eligibility check output.xlsx')

file = pd.ExcelFile('OS-HELP - Eligibility check output.xlsx')
sheet_1 = pd.read_excel(file, 'Report')
sheet_2 = pd.read_excel(file, 'MAP Calculations')

# checking eligibility (prem checks)
sheet_2['Student Eligible?'] = np.where((sheet_2['INTERNATIONAL?'] == 'AUSTRALIAN CITIZEN (INCL AUS CITIZENS WITH DUAL CITIZENSHIP)') & (sheet_2['FEE_CAT'].isin(['AU_DOM_CP', 'AU_DOM_P21', 'AU_DOM_P21_CP', 'AU_DOM_PP', 'AU_DOM_SW', 'AU_DOM']))
& (sheet_2['Credit Points Completed'] >= 48) & (sheet_1['CP_REQUIRED'] - sheet_2['Credit Points Completed'] >= 6),
True, False)

# Specify the name of the excel file
file_name = 'OS-HELP - Eligibility check output.xlsx'
  
# Saving the excelsheet
with pd.ExcelWriter('OS-HELP - Eligibility check output.xlsx') as file:
    sheet_1.to_excel(file, 'Report',index=False)
    sheet_2.to_excel(file, 'MAP Calculations',index=False)
    