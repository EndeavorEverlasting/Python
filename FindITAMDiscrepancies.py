import os
import pandas as pd
from datetime import datetime

def find_duplicates_and_missing_data(input_file):
    # ... (rest of the code remains the same)

    # Define a function to determine flagging reason
    def determine_flagging_reason(row):
        reasons = []

        if not pd.notna(row['Flr Pln N']):
            reasons.append('Missing Flr Pln N')
        if pd.notna(row['Flr Pln D']) and pd.notna(row['Department']):
            reasons.append('Duplicate based on Flr Pln D and Department')
        if pd.notna(row['Flr Pln D']) and pd.notna(row['Flr Pln L']):
            reasons.append('Duplicate based on Flr Pln D and Flr Pln L')

        flr_pln_d = row['Flr Pln D String']

        if flr_pln_d.startswith(('L', 'W')) and (not pd.notna(row['WS_Mon_Make_1']) or not pd.notna(row['WS_Mon_Mod_1'])):
            reasons.append('Incorrect Monitor Info')

        if reasons:
            return ', '.join(reasons)
        else:
            return 'Unknown Reason'

    # Apply the function to create the flagging reason column
    combined_data['Flagging Reason'] = combined_data.apply(
        determine_flagging_reason, axis=1)

    # Output filtered records to file:
    # Get the file name without extension
    file_name = os.path.splitext(input_file)[0]

    # Output combined data with modified file name and time in the file name
    combined_data.to_csv(f'{file_name}_{current_time}.csv', index=False)


# Example usage
input_file = 'PHELPS-NOV3.xlsx'
find_duplicates_and_missing_data(input_file)
