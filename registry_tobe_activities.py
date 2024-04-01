import os
import pandas as pd

def combine_activities(folder_path):
    # Initialize an empty DataFrame to store the consolidated activities
    activities_df = pd.DataFrame()

    # Iterate through files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "To-Be" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            # Read the "Activities list" tab from each To-Be excel file
            excel_data = pd.read_excel(excel_file_path, sheet_name="Activities list", na_values="")
            # Concatenate the data to the consolidated DataFrame
            activities_df = pd.concat([activities_df, excel_data], ignore_index=True)

    # Sort the table
    activities_df.sort_values(by=['Key'], inplace=True)
    
    # Replace NaN with "N/A" in the DataFrame
    activities_df = activities_df.fillna("N/A").astype(object)
    
    print("To-Be_Activities_Registry built successfully")
    
    return activities_df
