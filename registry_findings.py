import os
import pandas as pd

def combine_findings(folder_path):
    findings_df = pd.DataFrame()

    # Read To-Be data
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "To-Be" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            excel_data = pd.read_excel(excel_file_path, sheet_name="Findings Summary", na_values="")
            findings_df = pd.concat([findings_df, excel_data], ignore_index=True)

    # Replace NaN with "N/A" in the DataFrame
    findings_df = findings_df.fillna("N/A").astype(object)

    # Add new columns to the end of the DataFrame
    new_columns = [
        'Finding Type',
        'Financial Statement Only',
        'Root Cause Analysis Summary',
        'As-Is Recommendation',
        'As-Is Recommendation Rationale'
    ]

    for column in new_columns:
        findings_df[column] = "N/A"

    # Filling in Finding types data
    finding_type_list = [
        ('BR', 'Business Rule'),
        ('FD', 'Financial Note'),
        ('FN', 'Financial Note'),
        ('OP', 'Opportunity'),
        ('PP', 'Pain Point'),
        ('GA', 'Gap'),
        ('DT', 'Data')
    ]

    # Populate the findings type value based on the findings types list
    for index, row in findings_df.iterrows():
        # Iterate over each finding type abbreviation and full name tuple in the list
        for finding_type, full_name in finding_type_list:
            # Check if the finding type abbreviation is in the "Finding Key" column
            if finding_type in row['Finding Key']:
                # Assign the corresponding full name to the "Finding Type" column
                findings_df.at[index, 'Finding Type'] = full_name
                # Break the loop once a match is found
                break
                
    
    # Read As-Is data
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") and "As-Is" in file_name:
            excel_file_path = os.path.join(folder_path, file_name)
            excel_data = pd.read_excel(excel_file_path, sheet_name="Findings List", na_values="")
            as_is_findings_df = excel_data.copy()
            as_is_findings_df = as_is_findings_df.fillna("N/A").astype(object)
            
            
            # Merge data from as_is_findings_df to findings_df based on "Finding Key"
            for idx, row in findings_df.iterrows():
                matching_row = as_is_findings_df[as_is_findings_df['Finding Key'] == row['Finding Key']]
                if not matching_row.empty:
                    findings_df.at[idx, 'Financial Statement Only'] = matching_row['Financial Statement Only'].iloc[0]
                    findings_df.at[idx, 'Root Cause Analysis Summary'] = matching_row['Root Cause Analysis Summary'].iloc[0]
                    findings_df.at[idx, 'As-Is Recommendation'] = matching_row['As-Is Recommendation'].iloc[0]
                    findings_df.at[idx, 'As-Is Recommendation Rationale'] = matching_row['As-Is Recommendation Rationale'].iloc[0]

    # Reorder columns
    desired_columns = [
        'Process Scenario',
        'Finding Key',
        'Finding Type',
        'Finding Number',
        'Related As-Is Activity Number',
        'Financial Statement Only',
        'Related Finding(s)',
        'Finding Description',
        'Great 8',
        'Root Cause Analysis Summary',
        'As-Is Recommendation',
        'As-Is Recommendation Rationale',
        'Resolved in the To-Be State?',
        'To-Be Recommendations and Rationale'
    ]

    findings_df = findings_df[desired_columns]

    
    # Sort the table on "Performers" and "Process Scenario" columns
    findings_df.sort_values(by=['Process Scenario', 'Finding Number'], inplace=True)

    print("Findings_Registry built successfully")

    return findings_df
