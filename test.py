import pandas as pd
import excel2img

# Replace 'your_file.xlsx' with the actual file path
file_path = r"C:\Users\ZHamid2\OneDrive - SLB\scripts\data_ingestion_DDR\temp_files\DDR  -002 of 05 December  2018 ZHD A-111 - Copy.xlsx"

# Read Excel file into a DataFrame
df = pd.read_excel(file_path)
excel2img.export_img(file_path,"image.png/image.bmp","sheet!B2:H22")

# Display the DataFrame
print(df.columns)
print(type(df))
