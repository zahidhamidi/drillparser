import pandas as pd

# Replace 'your_directory' with the actual directory containing Excel files
file_path = r'C:\Users\ZHamid2\Downloads\MPCL-Halini_Activity.xlsx'

# Read all sheets in the Excel file into a dictionary of DataFrames
df = pd.read_excel(file_path, engine='openpyxl')


maintenance = ['service','service','troubleshoot','fix','repair']
rig_up = ['r/u','rig up','rig-up','rigging up']
rig_down = ['r/d','rig down','rig-down','rigging down']
circ = ['circulate','circulation','pump']

def activity_map(row):

    if ('drill' and 'hole') in row['Description'].lower():
        return "Drilling"
    elif ('rih' or 'run in hole') in row['Description'].lower():
        return "Tripping in"
    elif ('pooh' or 'pull out of hole') in row['Description'].lower():
        return "Tripping Out"
    elif ('function test' or 'f/t') in row['Description'].lower():
        return "Function Test"
    elif 'cement' in row['Description'].lower():
        return "Cementing"
    
    for keywords in circ:
        if keywords in row['Description'].lower():
            return "Well pumping/circulation"

    for keywords in maintenance:
        if keywords in row['Description'].lower():
            return "Equipment Maintenance"
        
    for keywords in rig_up:
        if keywords in row['Description'].lower():
            return "Rigging-up Equipment"
        
    for keywords in rig_down:
        if keywords in row['Description'].lower():
            return "Rigging-down Equipment"
        
    else:
        return row["ActivityCodeL1"]
    
    

    
    
    
df["ActivityCodeL1"] = df.apply(activity_map, axis=1)

# df.loc[df["Description"].str.contains('drill', case = False, na=False), "ActivityCodeL1"] = "Drilling"
nan = df["ActivityCodeL1"].isna().sum()
rows = df.shape[0]

# Export the DataFrame to an Excel file
excel_output_path = r'C:\Users\ZHamid2\Downloads\Halini_Activity_test.xlsx'
df.to_excel(excel_output_path, index=False)
print(f'DataFrame exported to Excel file: {excel_output_path}')
print("Unmapped Row % :", (nan/rows)*100)



