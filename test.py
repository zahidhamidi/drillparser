import pandas as pd

# Replace 'your_file.xlsx' with the actual file path
file_path = r"C:\Users\ZHamid2\OneDrive - SLB\scripts\data_ingestion_DDR\temp_files\DDR  -002 of 05 December  2018 ZHD A-111 - Copy.xlsx"

# Read Excel file into a DataFrame
df = pd.read_excel(file_path)

# Display the DataFrame
print(df.columns)
print(type(df))
