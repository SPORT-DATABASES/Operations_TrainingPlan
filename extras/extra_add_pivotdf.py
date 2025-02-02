
# Define the file paths
template_path = "Excel_template.xlsx"
output_filename = f"{next_sunday.strftime('%d%b')}_{next_saturday.strftime('%d%b')}.xlsx"
output_path = output_filename

# Delete the old workbook if it exists
if os.path.exists(output_path):
    os.remove(output_path)
    print(f"Old workbook '{output_path}' has been deleted.")

# Create a new workbook by copying the template
shutil.copy(template_path, output_path)
print(f"New workbook '{output_path}' has been created from the template.")

# Load the copied workbook
workbook = load_workbook(output_path)

# Insert unique dates into specific cells in the "Template" sheet
template_sheet = workbook['Template']
date_cells = ['C4', 'E4', 'G4', 'I4', 'K4', 'M4']
for date, cell in zip(unique_dates[:6], date_cells):  # Use up to 6 dates
    template_sheet[cell] = date.strftime('%a %d %b %Y')  # Format date

# Add or replace the "Data" sheet
if "Data" in workbook.sheetnames:
    del workbook["Data"]  # Remove the existing "Data" sheet to avoid duplication
worksheet = workbook.create_sheet("Data")

# Write column headers
headers = list(pivot_df.columns)
for col_idx, header in enumerate(headers, start=1):
    worksheet.cell(row=1, column=col_idx, value=header)
    worksheet.cell(row=1, column=col_idx).alignment = Alignment(horizontal="center", vertical="center")

# Write pivot DataFrame to the "Data" sheet
for row_idx, row in enumerate(pivot_df.itertuples(index=False), start=2):  # No index after reset
    for col_idx, cell_value in enumerate(row, start=1):
        cell = worksheet.cell(row=row_idx, column=col_idx, value=cell_value)
        if isinstance(cell_value, str):
            cell.alignment = Alignment(wrap_text=True, vertical="top")  # Enable wrap text for the cell

# Adjust column widths
for col_idx, header in enumerate(headers, start=1):
    column_letter = get_column_letter(col_idx)
    max_length = max(
        pivot_df.iloc[:, col_idx - 1].astype(str).map(len).max(),
        len(header)
    )
    worksheet.column_dimensions[column_letter].width = max_length + 2

# Save the workbook
workbook.save(output_path)
print(f"Pivot DataFrame successfully saved to the new workbook: {output_path}")
