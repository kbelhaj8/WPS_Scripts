import os
import pandas as pd

def combine_as_is_performers(folder_path):
    # Initialize an empty DataFrame to store the consolidated performers
    as_is_performers_df = pd.DataFrame()

    # Iterate through files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "As-Is" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            # Read the "Performers List" tab from each To-Be excel file
            excel_data = pd.read_excel(excel_file_path, sheet_name="As-Is_Performers", na_values="")
            # Concatenate the data to the consolidated DataFrame
            as_is_performers_df = pd.concat([as_is_performers_df, excel_data], ignore_index=True)

    # Remove trailing spaces from column names
    as_is_performers_df.columns = as_is_performers_df.columns.str.strip()

    # Reorder the columns based on the specified order
    as_is_performers_df = as_is_performers_df[['Performers', 'Description', 'Process Scenario']]

    # Sort the table on "Performers" and "Process Scenario" columns
    as_is_performers_df.sort_values(by=['Performers', 'Process Scenario'], inplace=True)

    # Combine values in "Process Scenario" column for rows with the same "Performers" and "Description"
    as_is_performers_df['Process Scenario'] = as_is_performers_df.groupby(['Performers', 'Description'])['Process Scenario'].transform(lambda x: '\n'.join(x))

    # Drop duplicate rows
    as_is_performers_df.drop_duplicates(subset=['Performers', 'Description'], inplace=True)

    # Replace NaN with "N/A" in the DataFrame
    as_is_performers_df = as_is_performers_df.fillna("N/A").astype(object)

    print("As-Is_Stakeholders_Registry built successfully")

    return as_is_performers_df