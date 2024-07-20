# Script to extract data from the NEET pdf result file
import re
import pandas as pd
import PyPDF2
import os

def extract_data_from_pdf(pdf_path):
    # Extract centre number from filename
    filename = os.path.basename(pdf_path)
    centre_number = filename.split('.')[0]  # Assumes filename is like "461202.pdf"

    content = ""
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            content += page.extract_text()

    print(f"Processing centre number: {centre_number}")

    # Extract Srlno and Marks
    pattern = r'(\d+)\s+(\d+|-?\d+)'
    matches = re.findall(pattern, content)

    data = {'Centre': [], 'Srlno': [], 'Marks': []}
    for match in matches:
        data['Centre'].append(centre_number)
        data['Srlno'].append(int(match[0]))
        data['Marks'].append(int(match[1]))

    return data

def create_excel(data, output_path):
    df = pd.DataFrame(data)
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column('A:C', 15)

    print(f"Excel file '{output_path}' has been created.")

# Directory containing PDF files
pdf_directory = 'C:/Users/ebhsrai/Downloads/NEET_PDFs/'  # Adjust this path
output_directory = 'C:/Users/ebhsrai/Downloads/NEET_PDFs/NEET_Excel_Output/'  # Adjust this path

# Create output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Process all PDF files in the directory
for filename in os.listdir(pdf_directory):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(pdf_directory, filename)
        data = extract_data_from_pdf(pdf_path)
        
        # Create output Excel filename
        excel_filename = f"{filename[:-4]}_marks.xlsx"
        output_excel = os.path.join(output_directory, excel_filename)
        
        # Create Excel file for this PDF
        create_excel(data, output_excel)

print("All PDF files have been processed and Excel files created.")