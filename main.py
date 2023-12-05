import pandas as pd
from openpyxl import load_workbook

# File paths
file1_path = "Finance Ulti Employee Details Fernando 12.04.2023.xlsx"
file2_path = "2023.10 TM Count by PC - Formulas.xlsx"
file3_path = "path/to/your/final/pivot_table/summary.xlsx"

# Step 1: Load existing data (File 2)
try:
    df = pd.read_excel(file2_path, sheet_name="data")
except FileNotFoundError:
    # If File 2 doesn't exist, create an empty DataFrame
    df = pd.DataFrame()

# Step 2: Process the new data (File 1)
try:
    # Determine the total number of rows in File 1
    total_rows = pd.read_excel(file1_path, sheet_name="data").shape[0]

    # Read the new data dynamically
    df_new = pd.read_excel(file1_path, sheet_name="data", skiprows=df.shape[0], nrows=total_rows - df.shape[0])
except FileNotFoundError:
    # Handle the case where File 1 doesn't exist or is empty
    df_new = pd.DataFrame()

# Identify a column that uniquely identifies each employee (e.g., 'Employee Number')
unique_column = 'Employee Number'

# Identify criteria for removing employees
employees_to_remove = df[~df[unique_column].isin(df_new[unique_column])]

# Identify criteria for adding new employees
employees_to_add = df_new[~df_new[unique_column].isin(df[unique_column])]

# Update the existing DataFrame with new employees and remove those no longer present
df = pd.concat([df, employees_to_add], ignore_index=True)
df = df[~df[unique_column].isin(employees_to_remove[unique_column])]

# Save the updated data to File 2
df.to_excel(file2_path, sheet_name="data", index=False)

# Step 3: Refresh the Pivot Table in File 2 (assuming the PivotTable is in 'PivotTableSheet')
book = load_workbook(file2_path)
writer = pd.ExcelWriter(file2_path, engine='openpyxl') 
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# Access the PivotTable Analyze tools and include the new range of data
# Assuming your PivotTable is in cell 'A1', you can adjust the range accordingly
pivot_table_sheet = writer.sheets['PivotTableSheet']
pivot_table_sheet.pivotTables[0].table_ref = f"data!$A$1:$Z${df.shape[0]}"  # Adjust the range dynamically

# Refresh the Pivot Table
pivot_table_sheet.refresh_pivot_table()

# Save the changes to File 2
writer.save()

# Step 4: Extract data from the Pivot Table with selected filters
filtered_data = df[(df['Employee Type'].isin(['Hourly', 'Salary'])) & (df['Employment Status'].isin(['Active', 'Leave of Absence']))]

# Step 5: Create a new Excel file with the summary sheet (File 3)
with pd.ExcelWriter(file3_path, engine='openpyxl') as writer:
    # Write the filtered data to the 'SummarySheet'
    filtered_data.to_excel(writer, sheet_name='SummarySheet', index=False)

print("Script completed successfully.")
