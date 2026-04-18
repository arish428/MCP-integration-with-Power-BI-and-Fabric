import os
from openpyxl import Workbook

# Folder containing your txt files
folder_path = r"D:\MCP_Microsoft_Fabric"

# Output Excel file
output_file = "columns_output.xlsx"

# Create workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "Columns"

# Header row
ws.append(["File Name", "Total Columns", "Column Name"])

# Loop through all txt files
for file_name in os.listdir(folder_path):
    if file_name.endswith(".txt"):
        file_path = os.path.join(folder_path, file_name)
        
        with open(file_path, 'r', encoding='utf-8') as f:
            data = f.read()
        
        # Step 1: Split rows
        rows = data.split("~@^*^@~")
        
        # Step 2: Extract header row
        header = rows[0]
        
        # Step 3: Split columns
        columns = header.split("~##~")
        
        # Step 4: Clean column names
        clean_columns = []
        for col in columns:
            col = col.strip().replace('"', '')
            if col:
                clean_columns.append(col)
        
        # Step 5: Write to Excel
        total_cols = len(clean_columns)
        for col in clean_columns:
            ws.append([file_name, total_cols, col])

# Save file
wb.save(output_file)

print(f"Output saved to {output_file}")