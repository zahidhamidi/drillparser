import pandas as pd
import numpy as np
import re
from fractions import Fraction
import os
from tqdm.auto import tqdm
from dateutil.parser import parse
from datetime import datetime


# Replace 'your_directory' with the actual directory containing Excel files
excel_directory = r'C:\Users\User\Downloads\test\HD1'

# Initialize an empty list to store individual DataFrames
dfs = []

# Flag to track if the first adjustment has been made
first_adjustment = False

# Iterate over all Excel files in the directory
for file_name in tqdm(os.listdir(excel_directory)):
    print(f"Running : {file_name}")
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(excel_directory, file_name)

        # Read all sheets in the Excel file into a dictionary of DataFrames
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        num_sheets = len(all_sheets)
        print(f"Number of sheets : {num_sheets}")

        # Iterate over all sheets
        for sheet_name, sheet_df in all_sheets.items():
            print(f"Sheet Name : {sheet_name}")
            # Find indices of "FROM", "TO", "OPERATION", and "Next 24 Hour Projection"
            from_indices = np.where(sheet_df == 'From')
            to_indices = np.where(sheet_df == 'To')
            operation_indices = np.where(sheet_df == 'Details Of Operations In Sequence And Remarks')
            next_projection_index = np.where(np.logical_or(sheet_df == 'Formation / Lithology', sheet_df == 'Current Status'))
            date_cell = np.where(sheet_df.apply(lambda row: any('Date:' in str(cell) for cell in row)))[0]



            # Extract row indices
            from_row = from_indices[0][0] if from_indices[0].size > 0 else None
            to_row = to_indices[0][0] if to_indices[0].size > 0 else None
            operation_row = operation_indices[0][0] if operation_indices[0].size > 0 else None
            next_projection_row = next_projection_index[0][0] if next_projection_index[0].size > 0 else None

            # Create a new DataFrame with columns "START", "END", and "OPERATION"
            if from_row is not None or to_row is not None or operation_row is not None or next_projection_row is not None:
                data_subset = sheet_df.iloc[operation_row + 1:next_projection_row, :]

                new_df = pd.DataFrame({
                    'START': data_subset.iloc[:, 0].tolist(),
                    'END': data_subset.iloc[:, 1].tolist(),
                    'OPERATION': data_subset.iloc[:, 3].tolist(),
                    'DATE': [sheet_name] * len(data_subset)
                })

                # Check if the sheet name is in date format, if not, try to find the date in the sheet cells
                try:
                    # Try parsing the sheet name as a date
                    parsed_date = parse(sheet_name)
                    new_df['DATE'] = parsed_date.strftime('%m/%d/%y')
                except ValueError:

                    # If sheet name is not a date, find the date in the cells
                    # date_cell = np.where(sheet_df.apply(lambda row: any('Date:' in str(cell) for cell in row)))[0]
                    date_cell = np.where(sheet_df == 'Date:')[0]

                    if len(date_cell) > 0:
                        # Extract the date using a regular expression
                        date_match = re.search(r'\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\w+\s\d{1,2},?\s\d{2,4}|\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2})\b',
                               sheet_df.iloc[date_cell[0],:].to_string())

                        

                        if date_match:
                            # Attempt to parse the matched date
                            try:
                                parsed_date = parse(date_match.group())

                                new_df['DATE'] = parsed_date.strftime('%m/%d/%y')
                            except ValueError:
                                print("Unable to parse the date.")
                        else:
                            # date_test = parse(sheet_df.iloc[date_cell[0],: ].to_string())
                            print("Error")
                    else:
                        print("No 'Date:' cell found.")
                    
                    # print(date_cell)
                    parsed_date = parse(date_match.group(0))
                    new_df['DATE'] = parsed_date.strftime('%m/%d/%y')

                # Remove rows where "OPERATION" is NaN or empty
                new_df = new_df.dropna(subset=['OPERATION'])

                # Replace NaN or empty values in "START" and "END" columns with "0000"
                new_df[['START', 'END']] = new_df[['START', 'END']].fillna('0000')

                # Convert START and END time columns to the desired format (HH:MM)
                def convert_time(x):
                    if isinstance(x, datetime):
                        # Convert datetime.time object to string and then handle it
                        x = x.strftime('%H:%M:%S')
                    
                    if (x == '0000') or (x == '2400'):
                        return '00:00'
                    elif ':' in str(x):
                        # If ':' exists, remove seconds and return
                        return str(x).split(':')[0] + ':' + str(x).split(':')[1]
                    else:
                        # Replace '/' with ':' and remove seconds
                        parts = str(x).replace('/', ':').split(':')
                        if len(parts) >= 2:
                            return parts[0] + ':' + parts[1]
                        else:
                            return str(x)  # Return the input as is if it cannot be split


                    
                


                new_df['START'] = new_df['START'].apply(convert_time)
                new_df['END'] = new_df['END'].apply(convert_time)

                # print(new_df)

                # Convert "START" and "END" columns to datetime objects if not empty
                new_df['START_DATE'] = pd.to_datetime(new_df['DATE'] + ' ' + new_df['START'], format='%m/%d/%y %H:%M', errors='coerce')
                new_df['END_DATE'] = pd.to_datetime(new_df['DATE'] + ' ' + new_df['END'], format='%m/%d/%y %H:%M', errors='coerce')




                # Adjust end date if it is after 12 am only for subsequent rows
                if first_adjustment:
                    try:
                        midnight_index = new_df[new_df['END_DATE'].dt.hour < new_df['START_DATE'].dt.hour].index
                        new_df.loc[midnight_index[0]:, ['END_DATE']] += pd.DateOffset(days=1)
                        new_df.loc[midnight_index[0]+1:, ['START_DATE']] += pd.DateOffset(days=1)
                    except IndexError:
                        pass

                # Convert "START_DATE" and "END_DATE" columns to the desired format (MM/DD/YY HH:MM)
                new_df['START_DATE'] = new_df['START_DATE'].dt.strftime('%m/%d/%y %H:%M')
                new_df['END_DATE'] = new_df['END_DATE'].dt.strftime('%m/%d/%y %H:%M')

                # Update the flag after the first adjustment
                first_adjustment = True

                # Drop the original "START" and "END" columns
                new_df = new_df.drop(['START', 'END', 'DATE'], axis=1)

                # Define a regex pattern to extract depth information
                depth_pattern = r'(\d+)\s*[mM]\s*to\s*(\d+)\s*[mM]'

                # Convert 'OPERATION' column to strings
                new_df['OPERATION'] = new_df['OPERATION'].astype(str)


                # Extract START_DEPTH and END_DEPTH using regex
                depth_extraction = new_df['OPERATION'].str.extract(depth_pattern)

                # If no match is found, set NaN for both START_DEPTH and END_DEPTH
                new_df['START_DEPTH'] = depth_extraction[0].astype(float) if 0 in depth_extraction.columns else pd.NA
                new_df['END_DEPTH'] = depth_extraction[1].astype(float) if 1 in depth_extraction.columns else pd.NA

                # Define a regex pattern to extract hole size information
                hole_size_pattern = r'(\d+(?:-|\s)?(?:\d+/\d+|\d+)?(?:\.\d+)?)\s*"\s*hole'

                # Extract HOLE_SIZE using regex
                hole_size_extraction = new_df['OPERATION'].str.extract(hole_size_pattern, flags=re.IGNORECASE)
                new_df['HOLE_SIZE'] = hole_size_extraction[0].apply(lambda x: float(Fraction(x.split('-')[0]) + Fraction(x.split('-')[1])) if isinstance(x, str) and '-' in x else float(Fraction(x.replace(' ', ''))) if pd.notna(x) else 0)

                # Define class keywords
                productive_keywords = ['drill','survey','safety meeting','ream', 'no loss','circulate','circulating','wiper trip','pump','RIH','deviation tool','TDS','prepare','pressure test']
                npt_keywords = ['repair','balled','stuck','fail','leak','mud loss','break','inspection']

                # Function to determine the class type based on the text value
                def get_class_type(text):
                    if isinstance(text, str) and any(keyword in text.lower() for keyword in productive_keywords):
                        return 'Productive'
                    elif isinstance(text, str) and any(keyword in text.lower() for keyword in npt_keywords):
                        return 'Non Productive'
                    else:
                        return pd.NA

                # Create a new column "Class_Type" based on the text in "OPERATION"
                new_df['Class_Type'] = new_df['OPERATION'].apply(get_class_type)

                print(new_df)

                # Append the DataFrame to the list
                dfs.append(new_df)

print(dfs)

# Concatenate all DataFrames into a final DataFrame
final_df = pd.concat(dfs, ignore_index=True)

# print(final_df)


# Export the DataFrame to an Excel file
excel_output_path = r'C:\Users\User\Downloads\output\depth_extraction_output.xlsx'
final_df.to_excel(excel_output_path, index=False)
print(f'DataFrame exported to Excel file: {excel_output_path}')
