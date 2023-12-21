import os
import pandas as pd
from datetime import datetime


def find_duplicates_and_missing_data(input_file, expected_count=None):
  """
    This function identifies and processes
    duplicate records, missing data, and 
    incorrect designators in an Excel file.

    To-Do:
    Find duplicate printers and special printers
    New designator standard and errors
    Sequential Designators
    New vs Old Records
    Obsolete Records
    Separate Incorrect Designators from Missing Designator
    Missing Department
    Refactoring
    

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
  # old line with cool code (worth looking into):
  # incorrect_designators = df[~df['Flr Pln D'].str.match(r'^[A-Z]\d+$')]
  # Doesn't work with empty entries. Implementation requires catching exceptions.

  # Check for incorrect designators
  incorrect_designators = []
  for index, row in df.iterrows():

    designator = row['Flr Pln D']
    epic_loc = row['EPIC_LOC']

    # Logic for checking incorrect designators:
    # if (not designator[0].isalpha() or not designator[1:].isdigit()):
    # Check if epic_loc is not n/a (i.e. it is a number or str)
    # Using the value will throw an error if it's n/a.
    # Skip the entry if EPIC has classified the device as remote in the epic_loc column
    #
    if pd.notna(epic_loc) and isinstance(epic_loc,
                                         str) and 'remote' in epic_loc.lower():
      continue  # Skip records where 'remote' is present in 'EPIC_LOC'

    # After skipping remote devices without a designator,
    # Check that the designator is not a n/a value.
    # If it's not a n/a value and doesn't start with
    # the letters of the five device types
    # or it doesn't proceed to be a number after the first character
    # then it's an incorrect designator.

    # Currently the script only looks for the first character to be a letter,
    # and it doesn't match the current number system:
    # Current number system: W2006, WOW2018
    # Old Design: W6, WOW18

    if pd.notna(designator) and (not isinstance(designator, str)
                                 or not designator[0].isalpha()
                                 or not designator[1:].isdigit()):
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

  def find_duplicate_id(row, df):
    # Find the ID of the record causing a duplicate.

    designator = row['Flr Pln D']
    department = row['Department']
    newFlrPln = row['Flr Pln N']
    oldFlrPln = row['Flr Pln L']

    # Check for duplicates based on 'Flr Pln D' and 'Flr Pln N'
    if pd.notna(designator) and pd.notna(newFlrPln):
      duplicate_row = df[(df['Flr Pln D'] == designator)
                         & (df['Flr Pln N'] == newFlrPln)].iloc[0]
      return duplicate_row['ID']

    # Check for duplicates based on 'Flr Pln D' and 'Flr Pln L'
    if pd.notna(designator) and pd.notna(oldFlrPln):
      duplicate_row = df[(df['Flr Pln D'] == designator)
                         & (df['Flr Pln L'] == oldFlrPln)].iloc[0]
      return duplicate_row['ID']

    # Check for duplicates based on 'Flr Pln D' and 'Department'
    if pd.notna(designator) and pd.notna(department):
      duplicate_row = df[(df['Flr Pln D'] == designator)
                         & (df['Department'] == department)].iloc[0]
      return duplicate_row['ID']

    return None

  # Find records where both 'Flr Pln N' and 'Flr Pln D', New URL and Designator,
  # match on two or more occurrences
  duplicates_n_d = df.groupby(['Flr Pln N',
                               'Flr Pln D']).filter(lambda x: len(x) >= 2)

  # Find records where both 'Flr Pln D' and 'Department', Department and Designator,
  # match on two or more occurrences
  duplicates_d_dept = df.groupby(['Flr Pln D',
                                  'Department']).filter(lambda x: len(x) >= 2)

  # Find records where both 'Flr Pln D' and 'Flr Pln L', Old URL and Designator,
  # match on two or more occurrences
  dupe_d_l = df.groupby(['Flr Pln D',
                         'Flr Pln L']).filter(lambda x: len(x) >= 2)

  # Concatenate the results
  duplicates = pd.concat([duplicates_n_d, duplicates_d_dept, dupe_d_l],
                         ignore_index=True)

  # Add a new column 'Duplicate_of_ID' to store the ID of the record causing a duplicate
  duplicates['Duplicate_of_ID'] = duplicates.apply(
      lambda row: find_duplicate_id(row, df), axis=1)

  # Find records with missing data in the Floor Plan URL column (e.g., 'Flr Pln N')
  missing_FlrPlnN = df[df['Flr Pln N'].isna()]

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
  print(f'Missing Data: {len(missing_FlrPlnN)}')
  print(f'Entries to Include: {len(entries_to_include)}')

  # Concatenate all identified records
  combined_data = pd.concat([
      df.loc[incorrect_designators_df['Incorrect Designator Index']],
      duplicates,
      missing_FlrPlnN,
      entries_to_include,
  ],
                            ignore_index=True)

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
      # There are 8 flagging reasons
      # 1) Blank Flr Pln N
      # 2) Blank Flr Pln L
      # 3) Dupe based on Department and Designator
      # 4) Dupe based on Flr Pln D and Old Floor Plan, Flr Pln L
      # 5) Dupe based on Flr Pln D and Old Floor Plan, Flr Pln N
      # 6) Incorrect Monitor Info for one reason or another
      # 7) Designator doesn't match given formats
      # 8) Flags remaining records as, "Correct"

      # 1)
      if not pd.notna(row['Flr Pln N']):
        reasons.append('Missing New Floor Plan')
      # 2)
      if not pd.notna(row['Flr Pln L']):
        reasons.append('Missing Old Floor Plan')
      # 3)
      if pd.notna(row['Flr Pln D']) and pd.notna(row['Department']):
        reasons.append('Duplicate based on Designator and Department')
      # 4)
      if pd.notna(row['Flr Pln D']) and pd.notna(row['Flr Pln L']):
        reasons.append('Duplicate based on Designator and Old Floor Plan')
      # 5)
      if pd.notna(row['Flr Pln D']) and pd.notna(row['Flr Pln N']):
        reasons.append('Duplicate based on Designator and New Floor Plan')

      flr_pln_d = row['Flr Pln D String']

      # 6)
      if flr_pln_d.startswith(
          ('L', 'W')) and (not pd.notna(row['WS_Mon_Make_1'])
                           or not pd.notna(row['WS_Mon_Mod_1'])):
        reasons.append('Incorrect Monitor Info')

      # 7)
      if not isinstance(
          flr_pln_d,
          str) or not flr_pln_d[0].isalpha() or not flr_pln_d[1:].isdigit():
        reasons.append('Incorrect Designator Info')

      # 8)
      if not isinstance(flr_pln_d, str):
        return 'moo'

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
input_file = 'PBMC EOD 12.19.23.xlsx'
find_duplicates_and_missing_data(input_file)
