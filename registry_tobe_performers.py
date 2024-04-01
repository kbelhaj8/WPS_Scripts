import os
import pandas as pd

def combine_performers(folder_path):
    # Initialize an empty DataFrame to store the consolidated performers
    performers_df = pd.DataFrame()

    # Iterate through files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "To-Be" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            # Read the "Performers List" tab from each To-Be excel file
            excel_data = pd.read_excel(excel_file_path, sheet_name="Performers List", na_values="")
            # Concatenate the data to the consolidated DataFrame
            performers_df = pd.concat([performers_df, excel_data], ignore_index=True)

    # Remove trailing spaces from column names
    performers_df.columns = performers_df.columns.str.strip()

    # Reorder the columns based on the specified order
    performers_df = performers_df[['Performers', 'Description', 'Process Scenario']]

    # Sort the table on "Performers" and "Process Scenario" columns
    performers_df.sort_values(by=['Performers', 'Process Scenario'], inplace=True)

    # Combine values in "Process Scenario" column for rows with the same "Performers" and "Description"
    performers_df['Process Scenario'] = performers_df.groupby(['Performers', 'Description'])['Process Scenario'].transform(lambda x: '\n'.join(x))

    # Drop duplicate rows
    performers_df.drop_duplicates(subset=['Performers', 'Description'], inplace=True)

    # Replace NaN with "N/A" in the DataFrame
    performers_df = performers_df.fillna("N/A").astype(object)
    
    print("To-Be_Stakeholders_Registry built successfully")

    return performers_df

