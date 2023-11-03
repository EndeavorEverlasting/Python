import os
import pandas as pd
from datetime import datetime


def find_duplicates_and_missing_data(input_file):
  """
    This function identifies and processes
    duplicate records, missing data, and 
    incorrect designators in an Excel file.

    Args:
        input_file (str): Path to the input Excel file.

    Returns:
        None
    """
  # Get the current date
  current_date = datetime.now().strftime('%Y-%m-%d')

  # Get the current time in military format
  current_time = datetime.now().strftime('%H%M%S')

  # Read the Excel file
  df = pd.read_excel(input_file, skiprows=1)

  # Clean up 'Flr Pln D' column by removing leading and trailing spaces
  df['Flr Pln D'] = df['Flr Pln D'].str.strip()

  # Check for incorrect designators
  # old line with cool code:
  # incorrect_designators = df[~df['Flr Pln D'].str.match(r'^[A-Z]\d+$')]

  # Check for incorrect designators
  incorrect_designators = []
  for index, row in df.iterrows():

    designator = row['Flr Pln D']
    epic_loc = row['EPIC_LOC']

    if pd.notna(epic_loc) and 'remote' in epic_loc.lower():
      continue  # Skip records where 'remote' is present in 'EPIC_LOC'

    if not isinstance(
        designator,
        str) or not designator[0].isalpha() or not designator[1:].isdigit():
      incorrect_designators.append(index)

  incorrect_designators_df = pd.DataFrame(
      incorrect_designators, columns=['Incorrect Designator Index'])
  '''
  # Suggest corrections
  for idx in incorrect_designators:
    sequence = df.loc[(df['Flr Pln N'] == df.loc[idx, 'Flr Pln N'])
                      & (df['Flr Pln L'] == df.loc[idx, 'Flr Pln L'])]
    new_designator = (
        f'{sequence["Flr Pln D"].str.extract("([A-Z]+)")[0].max()}'
        f'{sequence["Flr Pln D"].str.extract("([0-9]+)")[0].astype(int).max() + 1}'
    )

    df.loc[idx, 'Suggested Designator'] = new_designator
  '''
  # Find records where both 'Flr Pln N' and 'Flr Pln D', New URL and Designator,
  # match on two or more occurrences
  duplicates_n_d = df.groupby(['Flr Pln N',
                               'Flr Pln D']).filter(lambda x: len(x) >= 2)

  # Find records where both 'Flr Pln D' and 'Department', Department and Designator,
  # match on two or more occurrences
  duplicates_d_id = df.groupby(['Flr Pln D',
                                'Department']).filter(lambda x: len(x) >= 2)

  # Find records where both 'Flr Pln D' and 'Flr Pln L', Old URL and Designator,
  # match on two or more occurrences
  dupe_d_l = df.groupby(['Flr Pln D',
                         'Flr Pln L']).filter(lambda x: len(x) >= 2)

  # Concatenate the results
  duplicates = pd.concat([duplicates_n_d, duplicates_d_id, dupe_d_l],
                         ignore_index=True)

  # Find records with missing data in the Floor Plan URL column (e.g., 'Flr Pln N')
  missing_data = df[df['Flr Pln N'].isna()]

  # Filter entries where "Flr Pln D" does not start with "P" or "S",
  # or is not in the format of a letter followed by a number
  entries_to_include = df[~(df['Flr Pln D'].str.startswith(
      ('P', 'S')) | df['Flr Pln D'].str.match(r'^[A-Z]\d+$'))]

  # Check if "WS_Mon_Make_1" and "WS_Mon_Mod_1", the monitor Make and Model, are blank
  entries_to_include = entries_to_include[
      (entries_to_include['WS_Mon_Make_1'].isna()
       | entries_to_include['WS_Mon_Make_1'].eq('')) |
      (entries_to_include['WS_Mon_Mod_1'].isna()
       | entries_to_include['WS_Mon_Mod_1'].eq(''))]

  # Check the length of each dataframe
  print(f'Incorrect Designators: {len(incorrect_designators_df)}')
  print(f'Duplicates: {len(duplicates)}')
  print(f'Missing Data: {len(missing_data)}')
  print(f'Entries to Include: {len(entries_to_include)}')

  # Concatenate all identified records
  combined_data = pd.concat([
      df.loc[incorrect_designators_df['Incorrect Designator Index']],
      duplicates,
      missing_data,
      entries_to_include,
  ], ignore_index=True)

  if not combined_data.empty:
    # Add a new column for unique identifier (assuming you have some identifier)
    combined_data.insert(0, 'Unique Identifier',
                         range(1,
                               len(combined_data) + 1))

    # Add a new column 'Flr Pln D String' to store 'Flr Pln D' as strings
    combined_data['Flr Pln D String'] = combined_data['Flr Pln D'].astype(str)

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

      if flr_pln_d.startswith(
          ('L', 'W')) and (not pd.notna(row['WS_Mon_Make_1'])
                           or not pd.notna(row['WS_Mon_Mod_1'])):
        reasons.append('Incorrect Monitor Info')

      if not isinstance(
          flr_pln_d,
          str) or not flr_pln_d[0].isalpha() or not flr_pln_d[1:].isdigit():
        reasons.append('Incorrect Designator Info')

      if reasons:
        return ', '.join(reasons)
      else:
        return 'Unknown Reason'

    # Apply the function to create the flagging reason column
    combined_data['Flagging Reason'] = combined_data.apply(
        determine_flagging_reason, axis=1)

    # Add a new column 'Reasons' to store flagging reasons
    combined_data.insert(0, 'Reasons to Correct Entry',
                         combined_data['Flagging Reason'])

    # Remove the 'Flagging Reason' column
    # combined_data = combined_data.drop(columns=['Flagging Reason'])

    # Output flagging reasons for debugging
    print(combined_data[['Flr Pln D', 'Flagging Reason']])

    # Output filtered records to file:
    # Get the file name without extension
    file_name = os.path.splitext(input_file)[0]

    # Output combined data with modified file name and time in the file name
    output_file_name = f'{file_name}_{current_time}.csv'
    combined_data.to_csv(output_file_name, index=False)
    # Inform the user about the generated CSV file
    print(f'CSV file "{output_file_name}" generated successfully.')


# Example usage
input_file = 'SIUHN-NOV3.xlsx'
find_duplicates_and_missing_data(input_file)
