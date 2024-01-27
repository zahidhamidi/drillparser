import pandas as pd
import numpy as np
import re
from fractions import Fraction
import os
from tqdm.auto import tqdm
from dateutil.parser import parse
from datetime import datetime
import sys
from styleframe import StyleFrame, Styler
import streamlit as st

def main(excel_directory,wguid, wname, wbname, time_start_marker,time_start_index,time_end_marker,time_end_index,activity_start_marker,activity_end_marker,date_marker,activity_index):
    
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
            # all_sheets = pd.read_excel(file_path)
            
       
           
            num_sheets = len(all_sheets)
            print(f"Number of sheets : {num_sheets}")

            # Iterate over all sheets
            for sheet_name, sheet_df in all_sheets.items():
                print(f"Sheet Name : {sheet_name}")
                print(sheet_df)
                print(type(sheet_df))


                # Find indices of "FROM", "TO", "OPERATION", and "Next 24 Hour Projection"
                from_indices = np.where(sheet_df == time_start_marker)  # FROM
                to_indices = np.where(sheet_df == time_end_marker)  # TO

                operation_indices = np.where(sheet_df == activity_start_marker)  #  Details Of Operations In Sequence And Remarks
                next_projection_index = np.where(np.logical_or(sheet_df == activity_end_marker, sheet_df == activity_end_marker))    # Current Status 
                date_cell = np.where(sheet_df.apply(lambda row: any(date_marker in str(cell) for cell in row)))[0] # Date :
                # NPT_indices = np.where(sheet_df == NPT_code_marker)  # NPT code
                # act_code_indices = np.where(sheet_df == activity_code_marker)  # NPT code

 


                # Extract row indices
                from_row = from_indices[0][0] if from_indices[0].size > 0 else None
                to_row = to_indices[0][0] if to_indices[0].size > 0 else None
                operation_row = operation_indices[0][0] if operation_indices[0].size > 0 else None
                next_projection_row = next_projection_index[0][0] if next_projection_index[0].size > 0 else None

  
                # Create a new DataFrame with columns "START", "END", and "OPERATION"
                if from_row is not None or to_row is not None or operation_row is not None or next_projection_row is not None:

                    data_subset = sheet_df.iloc[operation_row + 1:next_projection_row, :] 
                      


                    new_df = pd.DataFrame({
                        'START': data_subset.iloc[:, int(time_start_index)].tolist(),
                        'END': data_subset.iloc[:, int(time_end_index)].tolist(),
                        'OPERATION': data_subset.iloc[:, int(activity_index)].tolist(),
                        # 'NPT CODE' : data_subset.iloc[:, int(NPT_code_marker)].tolist(),
                        # 'ACTIVITY CODE' : data_subset.iloc[:, int(activity_code_index)].tolist(),
                        'DATE': [sheet_name] * len(data_subset)
                    })

                    

                    # Check if the sheet name is in date format, if not, try to find the date in the sheet cells
                    try:

                        # Try parsing the sheet name as a date
                        parsed_date = parse(sheet_name)
                        new_df['DATE'] = parsed_date.strftime('%m/%d/%y')
                        

                    except ValueError:

                        # If sheet name is not a date, find the date in the cells
                        date_cell = np.where(sheet_df == date_marker)[0]
                        print(date_marker)

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
                                print("Error")
                        else:
                            print("No 'Date:' cell found.")
                            new_df['DATE'] = ''
                            
                        
                        

                        if new_df['DATE'].astype(str).str.strip().any() != '': 
                            parsed_date = parse(date_match.group(0))
                            new_df['DATE'] = parsed_date.strftime('%m/%d/%y')

                    # Remove rows where "OPERATION" is NaN or empty
                    new_df = new_df.dropna(subset=['OPERATION'])
                    

                    # Replace NaN or empty values in "START" and "END" columns with "0000"
                    new_df[['START', 'END']] = new_df[['START', 'END']].fillna('xxxx')

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

                    # Convert 'DATE' column to datetime objects
                    new_df['DATE'] = pd.to_datetime(new_df['DATE'], format='%m/%d/%y')

                                        # Function to apply styling
                    """
                    def adjust_dates(row):
                        if row['START'] == row['END']:
                            row['DATE'] += pd.DateOffset(days=1)

                        return row["DATE"]

                    # Apply the custom function to each row
                    new_df["DATE"] = new_df.apply(adjust_dates, axis=1)
                    print(new_df["DATE"])
                    """

                    
                    # Convert "START_DATE" and "END_DATE" columns to datetime objects if not empty
                    new_df['START_DATE'] = pd.to_datetime(new_df['DATE'].dt.strftime('%m/%d/%y') + ' ' + new_df['START'], format='%m/%d/%y %H:%M', errors='coerce')
                    new_df['END_DATE'] = pd.to_datetime(new_df['DATE'].dt.strftime('%m/%d/%y') + ' ' + new_df['END'], format='%m/%d/%y %H:%M', errors='coerce')


                    # new_df['END_DATE'] = pd.to_datetime(new_df['DATE'] + ' ' + new_df['END'], format='%m/%d/%y %H:%M', errors='coerce')

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

                    

                    # Convert 'EndDate' back to the desired format
                    # new_df['END_DATE'] = new_df['END_DATE'].dt.strftime('%m/%d/%y %H:%M')



                    # Define a regex pattern to extract depth information
                    # depth_pattern = r'(\d+)\s*[mM]\s*to\s*(\d+)\s*[mM]'
                    # depth_pattern = r'(\d+(?:\.\d{1,2})?)\s*[mM]\s*to\s*(\d+(?:\.\d{1,2})?)\s*[mM]'
                    # depth_pattern = r'(?:\b(\d+(?:\.\d{1,2})?)\s*[mM]\s*)?to\s*(\d+(?:\.\d{1,2})?)\s*[mM]'
                    depth_pattern = re.compile(r'(?:\b(\d+(?:\.\d{1,2})?)\s*[mM]\s*)?to\s*(\d+(?:\.\d{1,2})?)\s*[mM]', re.IGNORECASE)
                    # depth_pattern = re.compile(r'(?=.*\bdrill(?:ing)?\b)(?:\b(\d+(?:\.\d{1,2})?)\s*[mM]\s*)?to\s*(\d+(?:\.\d{1,2})?)\s*[mM]', re.IGNORECASE)

                    # Convert 'OPERATION' column to strings
                    new_df['OPERATION'] = new_df['OPERATION'].astype(str)

                    # Extract START_DEPTH and END_DEPTH using regex                  
                    depth_extraction = new_df['OPERATION'].str.extract(depth_pattern)

                    # If no match is found, set NaN for both START_DEPTH and END_DEPTH
                    new_df['START_DEPTH'] = depth_extraction[0].astype(float) if 0 in depth_extraction.columns else pd.NA
                    new_df['END_DEPTH'] = depth_extraction[1].astype(float) if 1 in depth_extraction.columns else pd.NA

                    # Replace None with NaN
                    new_df = new_df.where(pd.notna(new_df), np.nan)

                    # Fill NaN values in 'START_DEPTH' with the last observed value from 'END_DEPTH'
                    new_df['START_DEPTH'].fillna(method='ffill', inplace=True, limit=1, downcast='int')

                    # Fill remaining NaN values in 'END_DEPTH' with the last observed value from 'END_DEPTH'
                    new_df['END_DEPTH'].fillna(method='ffill', inplace=True)
                    
                    # Check if the word "drill" exists partially or as a subword in each row of 'OPERATION'
                    drill_mask = new_df['OPERATION'].str.contains(r'drill', case=False, na=False)

                    # Set START_DEPTH and END_DEPTH to pd.NA for rows where "drill" is not present
                    new_df.loc[~drill_mask, ['START_DEPTH', 'END_DEPTH']] = pd.NA




                    # Define a regex pattern to extract hole size information
                    hole_size_pattern = r'(\d+(?:-|\s)?(?:\d+/\d+|\d+)?(?:\.\d+)?)\s*"\s*hole'
                    # hole_size_pattern = re.compile(r'\b(\d+(?:\s*\d+/\d+)?(?:\.\d+)?)\s*"\s*hole', flags=re.IGNORECASE)

                    # Extract HOLE_SIZE using regex
                    hole_size_extraction = new_df['OPERATION'].str.extract(hole_size_pattern, flags=re.IGNORECASE)
                    new_df['HOLE_SIZE'] = hole_size_extraction[0].apply(lambda x: float(Fraction(x.split('-')[0]) + Fraction(x.split('-')[1])) if isinstance(x, str) and '-' in x else float(Fraction(x.replace(' ', ''))) if pd.notna(x) else '')

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

                    # Append the DataFrame to the list
                    dfs.append(new_df)


   
    # Concatenate all DataFrames into a final DataFrame
    final_df = pd.concat(dfs)



    # Add new columns to the final DataFrame
    final_df.insert(0, 'WellGUID', wguid)
    final_df.insert(1, 'WellName', wname)
    final_df.insert(2, 'WellboreName', wbname)
    final_df['ActivityStartDepthMD'] = ''
    final_df['ActivityEndDepthMD'] = ''
    final_df['TopDepthTVD'] = ''
    final_df['BottomDepthTVD'] = ''
    final_df['ActivityCodeL1'] = ''
    final_df['ActivityCodeL2'] = ''
    final_df['ActivityCodeL3'] = ''
    final_df['ActivityCodeL4'] = ''
    final_df['ActivityCodeL5'] = ''
    final_df['ActivityCodeL6'] = ''
    final_df['NPTReason'] = ''
    final_df['NPTDetail'] = ''
    final_df['UPDATE_DATE'] = ''

    # Define a dictionary to map old column names to new column names
    column_mapping = {'OPERATION': 'Description', 'START_DATE': 'StartDate', 'END_DATE': 'EndDate', 'START_DEPTH': 'TopDepthMD', 'END_DEPTH': 'BottomDepthMD',
                       'HOLE_SIZE': 'WellboreDiameter', 'Class_Type': 'TimeClassification','NPT CODE':'NPTDetail','ACTIVITY CODE':'ActivityCodeL1'}

    # Use the rename method to rename the columns
    final_df.rename(columns=column_mapping, inplace=True)

    # Now, reorder the columns to place the newly added ones at the beginning
    final_df = final_df[['WellGUID', 'WellName', 'WellboreName', 'WellboreDiameter','StartDate','EndDate','ActivityStartDepthMD','ActivityEndDepthMD','TopDepthMD',
                         'BottomDepthMD','TopDepthTVD','BottomDepthTVD','ActivityCodeL1','ActivityCodeL2','ActivityCodeL3','ActivityCodeL4','ActivityCodeL5','ActivityCodeL6',
                         'Description','TimeClassification','NPTReason','NPTDetail','UPDATE_DATE']]
  
    # Remove rows where Description is empty
    final_df = final_df[final_df['Description'].astype(str).str.strip() != '']

    output_directory = r"C:\Users\ZHamid2\OneDrive - SLB\scripts\data_ingestion_DDR\temp_files\output"

    

    # Function to check datestamp continuity and add a 'Discontinuity' column
    def mark_discontinuity(df):
        # Check if 'StartDate' and 'EndDate' columns exist in the DataFrame
        if 'StartDate' not in df.columns or 'EndDate' not in df.columns:
            raise ValueError("Columns 'StartDate' and 'EndDate' not found in the DataFrame.")

        # Initialize a new column 'Discontinuity' with 'No'
        df['Discontinuity'] = 'No'

        # Iterate through rows
        for i in range(len(df) - 1):
            current_end_date = df.iloc[i]['EndDate']
            next_start_date = df.iloc[i + 1]['StartDate']

            # Check for empty or null values
            if pd.isnull(current_end_date) or pd.isnull(next_start_date):
                continue

            # Convert dates to datetime objects for comparison
            current_end_date = pd.to_datetime(current_end_date)
            next_start_date = pd.to_datetime(next_start_date)

            # Check for continuity
            if current_end_date != next_start_date:
                # Mark the 'Discontinuity' column as 'Yes' for discontinuous rows
                df.at[i + 1, 'Discontinuity'] = 'Yes'

        return df  # Return the DataFrame with the 'Discontinuity' column

    # Call the function
    final_df = mark_discontinuity(final_df)

    # Remove rows where both StartDate and EndDate are empty or null
    final_df = final_df.dropna(subset=['StartDate', 'EndDate'], how='all')
    
    

    """

    # Create StyleFrame from DataFrame
    sf = StyleFrame(final_df)

    # Define a Styler object with the desired style (yellow background)
    style = Styler(bg_color='yellow')

    # Iterate through columns and apply style based on the condition
    for col_name in final_df.columns:
        condition = final_df['StartDate'] == final_df['EndDate']
        sf.apply_style_by_indexes(sf[condition], cols_to_style=col_name, styler_obj=style)

    final_df.style.apply(lambda row: ["background-color: yellow" if str(row['StartDate']) == str(row['EndDate']) else "background-color: white"] * len(row), axis=1)

    """


    
    excel_output_path = os.path.join(output_directory, 'DDR_extraction_output.xlsx')

    final_df.to_excel(excel_output_path, index=False)

    # Check if the directory exists, create it if not
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)


if __name__ == "__main__":
    if len(sys.argv) != 13:
        # print("Usage: python data_ingestion_DDR.py <excel_directory> <wguid> <wname> <wbname>")
        sys.exit(1)

    excel_directory = sys.argv[1]
    wguid = sys.argv[2]
    wname = sys.argv[3]
    wbname = sys.argv[4]
    time_start_marker = sys.argv[5]
    time_start_index = sys.argv[6]
    time_end_marker = sys.argv[7]
    time_end_index = sys.argv[8]
    activity_start_marker = sys.argv[9]
    activity_end_marker = sys.argv[10]
    date_marker = sys.argv[11]
    activity_index = sys.argv[12]
    # NPT_code_index = sys.argv[13]
    # NPT_code_marker = sys.argv[14]
    # activity_code_marker = sys.argv[15]
    # activity_code_index = sys.argv[16]

    main(excel_directory, wguid, wname, wbname,time_start_marker,time_start_index,time_end_marker,time_end_index,activity_start_marker,
         activity_end_marker,date_marker,activity_index)


