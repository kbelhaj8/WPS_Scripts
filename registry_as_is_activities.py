import os
import pandas as pd
import warnings
from openpyxl import load_workbook

# Suppress the warning
warnings.simplefilter("ignore", category=UserWarning)

def combine_as_is_activities(folder_path):
    # Initialize an empty DataFrame to store the consolidated activities
    as_is_activities_df = pd.DataFrame()

    # Iterate through files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "As-Is" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            # Read the "Activities list" tab from each To-Be excel file
            excel_data = pd.read_excel(excel_file_path, sheet_name="Process Steps",)
            # Concatenate the data to the consolidated DataFrame
            as_is_activities_df = pd.concat([as_is_activities_df, excel_data], ignore_index=True)
    

    # Replace NaN with "N/A" in the DataFrame
    as_is_activities_df = as_is_activities_df.fillna("N/A").astype(object)

    #re-order the columns and get rid of extra columns:
    as_is_activities_df = as_is_activities_df[['Process Scenario', 'Key', 'Integrated Activity List (IAL)', 'As-Is Activity Number', 'As-Is Activity Type', 'As-Is Activity Name', 'As-Is Activity Description', 'USA', 'USAF', 'USCG', 'USMC', 'USN', 'DCMA', 'DFAS', 'DLA']]

    #clean up Great 8 columns and format them:
    columns_to_check = ['USA', 'USAF', 'USCG', 'USMC', 'USN', 'DCMA', 'DFAS', 'DLA']
    for column in columns_to_check:
        as_is_activities_df[column] = as_is_activities_df[column].replace("N/A", "")

    # Sort the table o
    as_is_activities_df.sort_values(by=['Process Scenario'], inplace=True)

    print("To-As-Is_Stakeholders_Registry built successfully")
    
    return as_is_activities_df

