import pandas as pd
import os

def process_table_data(folder_path):
    # Initialize empty DataFrames to store the consolidated data
    to_be_process_overview_df = pd.DataFrame()
    as_is_process_overview_df = pd.DataFrame()

    # Iterate through files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            excel_file_path = os.path.join(folder_path, file_name)
            if "To-Be" in file_name:
                # For To-Be files, read the "Process Overview" tab
                excel_data = pd.read_excel(excel_file_path, sheet_name="Process Overview", na_values="")
                # Concatenate the data to the consolidated DataFrame for To-Be files
                to_be_process_overview_df = pd.concat([to_be_process_overview_df, excel_data], ignore_index=True)
            elif "As-Is" in file_name:
                # For As-Is files, read the "Process Overview" tab
                excel_data = pd.read_excel(excel_file_path, sheet_name="Process Overview", na_values="")
                # Concatenate the data to the consolidated DataFrame for As-Is files
                as_is_process_overview_df = pd.concat([as_is_process_overview_df, excel_data], ignore_index=True)

    # Rename the DataFrames
    to_be_process_overview_df.name = "To-Be_pf_Process_Overview"
    as_is_process_overview_df.name = "As-Is_pf_Process_Overview"

    # Create a new DataFrame based on To-Be DataFrame and add As-Is Process Description
    new_df = to_be_process_overview_df.copy()  # Create a copy of To-Be DataFrame

    # Merge As-Is Process Description based on matching "Process Scenario"
    new_df["As-Is Process Description"] = new_df["Process Scenario"].map(
        as_is_process_overview_df.set_index("Process Scenario")["As-Is Process Description"]
    )

    # Replace NaN with "N/A" in the DataFrame
    new_df = new_df.fillna("N/A").astype(object)

    # Sort the table
    new_df.sort_values(by=['Process Scenario'], inplace=True)

    print("Process_Registry built successfully")


    return new_df
