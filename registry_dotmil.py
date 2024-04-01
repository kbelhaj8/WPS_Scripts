import os
import pandas as pd

def combine_dotmil_changes(folder_path):
    dotmil_changes_df = pd.DataFrame()

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "To-Be" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            excel_data = pd.read_excel(excel_file_path, sheet_name="DOTmLPF-P Changes", na_values="")
            dotmil_changes_df = pd.concat([dotmil_changes_df, excel_data], ignore_index=True)

    # Replace NaN with "N/A" in the DataFrame
    dotmil_changes_df = dotmil_changes_df.fillna("N/A").astype(object)

    # Sort the table
    dotmil_changes_df.sort_values(by=['Process Scenario'], inplace=True)

    print("DOTmLPF-P_Registry built successfully")

    return dotmil_changes_df
