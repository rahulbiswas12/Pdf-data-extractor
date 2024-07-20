# Script to marge multiple excel files into one single file
import os
import pandas as pd
import glob

# Directory containing the Excel files
input_directory = 'C:/Users/ebhsrai/Downloads/NEET_PDFs/NEET_Excel_Output'  # Adjust this path
output_file = 'C:/Users/ebhsrai/Downloads/NEET_Combined_Marks.xlsx'  # Adjust this path

# Get all Excel files in the directory
all_files = glob.glob(os.path.join(input_directory, "*_marks.xlsx"))

print(f"Found {len(all_files)} Excel files in the directory.")

if len(all_files) == 0:
    print("No Excel files found. Please check the input directory path.")
    exit()

# List to store dataframes
df_list = []

# Read each Excel file and append to the list
for file in all_files:
    try:
        df = pd.read_excel(file)
        if not df.empty:
            df_list.append(df)
            print(f"Successfully read {file}")
        else:
            print(f"Warning: {file} is empty")
    except Exception as e:
        print(f"Error reading {file}: {str(e)}")

if len(df_list) == 0:
    print("No valid data found in any of the Excel files.")
    exit()

# Combine all dataframes in the list
combined_df = pd.concat(df_list, ignore_index=True)

# Sort the combined dataframe by Centre and Srlno
combined_df = combined_df.sort_values(['Centre', 'Srlno'])

# Write the combined dataframe to a new Excel file
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    combined_df.to_excel(writer, sheet_name='Combined Marks', index=False)
    
    # Auto-adjust columns' width
    worksheet = writer.sheets['Combined Marks']
    for i, col in enumerate(combined_df.columns):
        column_len = max(combined_df[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, column_len + 2)

print(f"All Excel files have been combined into {output_file}")
print(f"Total rows in combined file: {len(combined_df)}")