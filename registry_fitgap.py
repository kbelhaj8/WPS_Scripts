import os
import pandas as pd

def combine_fitgap(folder_path):
    fitgap_df = pd.DataFrame()

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "To-Be" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            excel_data = pd.read_excel(excel_file_path, sheet_name="Fit-Gap", na_values="")
            fitgap_df = pd.concat([fitgap_df, excel_data], ignore_index=True)

    # Replace NaN with "N/A" in the DataFrame
    fitgap_df = fitgap_df.fillna("N/A").astype(object)

    # Sort the table
    fitgap_df.sort_values(by=['Process Scenario'], inplace=True)

    print("Fit-Gap_Registry built successfully")

    return fitgap_df
