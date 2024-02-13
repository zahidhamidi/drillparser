import pandas as pd
import re
import numpy as np
from fractions import Fraction
import os
# from googletrans import Translator, constants
from pprint import pprint

# Replace 'your_file.xlsx' with the actual file path
file_path = r"C:\Users\ZHamid2\OneDrive - SLB\Downloads\BN-95-2-2-15H__Activity.xlsx"

# Read Excel file into a DataFrame
new_df = pd.read_excel(file_path)


"""
depth_pattern = re.compile(r'(?:\b(\d+(?:\.\d{1,2})?)\s*[mM]\s*)?hasta\s*(\d+(?:\.\d{1,2})?)\s*[mM]', re.IGNORECASE)


# Convert 'OPERATION' column to strings
new_df['Description'] = new_df['Description'].astype(str)

# Check the condition and filter rows where 'ActivityCodeL1' is equal to "Drilling"
drilling_mask = new_df['ActivityCodeL1'] == "Drilling"
drilling_rows = new_df.loc[drilling_mask]

if not drilling_rows.empty:

    # Extract START_DEPTH and END_DEPTH using regex                  
    depth_extraction = new_df['Description'].str.extract(depth_pattern)

    # If no match is found, set NaN for both START_DEPTH and END_DEPTH
    new_df['ActivityStartDepthMD'] = depth_extraction[0].astype(float) if 0 in depth_extraction.columns else pd.NA
    new_df['ActivityEndDepthMD'] = depth_extraction[1].astype(float) if 1 in depth_extraction.columns else pd.NA

    # Replace None with NaN
    new_df = new_df.where(pd.notna(new_df), np.nan)

    # # Fill NaN values in 'START_DEPTH' with the last observed value from 'END_DEPTH'
    # new_df['ActivityStartDepthMD'].fillna(method='ffill', inplace=True, limit=1, downcast='int')

    # # Fill remaining NaN values in 'END_DEPTH' with the last observed value from 'END_DEPTH'
    # new_df['ActivityEndDepthMD'].fillna(method='ffill', inplace=True)

    # # Check if the word "drill" exists partially or as a subword in each row of 'OPERATION'
    # drill_mask = new_df['Description'].str.contains(r'perforando', case=False, na=False)

    # # Set START_DEPTH and END_DEPTH to pd.NA for rows where "drill" is not present
    # new_df.loc[~drill_mask, ['ActivityStartDepthMD', 'ActivityEndDepthMD']] = pd.NA
"""

# Define depth pattern (assuming you have a valid regex pattern)
# depth_pattern = re.compile(r'(?:\b(\d+(?:\.\d{1,2})?)\s*[mM]\s*)?hasta\s*(\d+(?:\.\d{1,2})?)\s*[mM]', re.IGNORECASE)
# depth_pattern = re.compile(r'(\d+(?:\.\d{1,2})?)\s*[mM]?\s*(?:mts)?\s*hasta\s*(\d+(?:\.\d{1,2})?)\s*[mM]?\s*(?:mts)?', re.IGNORECASE)
depth_pattern = re.compile(r'(\d+(?:\.\d{1,2})?)\s*[mM]?\s*(?:mts)?\s*(?:\([^)]*\))?\s*hasta\s*(\d+(?:\.\d{1,2})?)\s*[mM]?\s*(?:mts)?', re.IGNORECASE)
# depth_pattern = re.compile(r'(?:\b(\d+(?:\.\d{1,2})?)\s*[mM]?\s*(?:mts)?\s*(?:\([^)]*\))?)?hasta\s*(\d+(?:\.\d{1,2})?)\s*[mM]?\s*(?:mts)?', re.IGNORECASE)






# Check the condition and filter rows where 'ActivityCodeL1' is equal to "Drilling"
drilling_mask = new_df['ActivityCodeL1'] == "Drilling"
drilling_rows = new_df.loc[drilling_mask]

if not drilling_rows.empty:

    # Extract START_DEPTH and END_DEPTH using regex                  
    depth_extraction = new_df['Description'].str.extractall(depth_pattern)

    # If no match is found, set NaN for both START_DEPTH and END_DEPTH
    start_depths = depth_extraction[0].astype(float).groupby(level=0).first()
    end_depths = depth_extraction[1].astype(float).groupby(level=0).last()

    # Set NaN for rows without depth ranges
    new_df['ActivityStartDepthMD'] = np.nan
    new_df['ActivityEndDepthMD'] = np.nan

    # Update start depth for rows with depth ranges
    new_df.loc[drilling_mask, 'ActivityStartDepthMD'] = start_depths
    new_df.loc[drilling_mask, 'ActivityEndDepthMD'] = end_depths

# Display the resulting DataFrame
print(new_df)




# Define a regex pattern to extract hole size information
hole_size_pattern = r'(\d+(?:-|\s)?(?:\d+/\d+|\d+)?(?:\.\d+)?)\s*"\s*hoyo'
# hole_size_pattern = r'\b(\d+(?:[-\s]\d+/\d+|\.\d+)?)\s*"\s*(?:hoyo\s*|\s*hoyo\b)'
# hole_size_pattern = re.compile(r'\b(\d+(?:\s*\d+/\d+)?(?:\.\d+)?)\s*"\s*hoyo', flags=re.IGNORECASE)
# hole_size_pattern = r'(\d+(?:[-\s]\d+/\d+|\.\d+)?)\s*"\s*(?:(?=.*\bhoyo\b)|(?<=\bhoyo\b\s*), flags=re.IGNORECASE)'


# Extract HOLE_SIZE using regex
hole_size_extraction = new_df['Description'].str.extract(hole_size_pattern, flags=re.IGNORECASE)
new_df['WellboreDiameter'] = hole_size_extraction[0].apply(lambda x: float(Fraction(x.split('-')[0]) + Fraction(x.split('-')[1])) if isinstance(x, str) and '-' in x else float(Fraction(x.replace(' ', ''))) if pd.notna(x) else '')


# Display the DataFrame with translated descriptions
print(new_df["WellboreDiameter"])


excel_output_path = r"C:\Users\ZHamid2\OneDrive - SLB\Downloads\extraction_depth_hole.xlsx"
new_df.to_excel(excel_output_path, index=False)