import os
import pandas as pd

def combine_systems (folder_path):
    # Initialize an empty DataFrame to store the consolidated performers
    systems_df = pd.DataFrame()

    # Iterate through files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "To-Be" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            # Read the "Systems Overview" tab from each To-Be excel file
            excel_data = pd.read_excel(excel_file_path, sheet_name="Systems Overview", na_values="")
            # Concatenate the data to the consolidated DataFrame
            systems_df = pd.concat([systems_df, excel_data], ignore_index=True)

    # Remove trailing spaces from column names
    systems_df.columns = systems_df.columns.str.strip()

    # Reorder the columns based on the specified order
    systems_df = systems_df[['System Short Name', 'System Full Name', 'To-Be Status', 'To-Be Status Description', 'Process Scenario']]

    # Sort the table on "Performers" and "Process Scenario" columns
    systems_df.sort_values(by=['System Short Name', 'Process Scenario'], inplace=True)

    # Combine values in "Process Scenario" column for rows with the same "Systems", "System Short Name", and "Descriptions"
    systems_df['Process Scenario'] = systems_df.groupby(['System Short Name', 'System Full Name', 'To-Be Status', 'To-Be Status Description'])['Process Scenario'].transform(lambda x: '\n'.join(x))

    # Drop duplicate rows
    systems_df.drop_duplicates(subset=['System Short Name', 'System Full Name', 'To-Be Status', 'To-Be Status Description'], inplace=True)

    # Replace NaN with "N/A" in the DataFrame
    systems_df = systems_df.fillna("N/A").astype(object)

    print("Systems_Registry built successfully")

    return systems_df

