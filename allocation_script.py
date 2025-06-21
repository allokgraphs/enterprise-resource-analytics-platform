import pandas as pd
import os
from datetime import datetime
import calendar
import re

def generate_resource_allocation_report(file_path):
    # Read the Excel file
    print(f"Reading file: {file_path}")
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Filter the dataframe to include only those with specific Working Status
    df = df[df['Working Status'].isin(['Working'])]
    
    # Get current month and year
    today = datetime.now()
    current_month = today.month
    current_year = today.year
    
    # Find date columns using multiple approaches
    date_columns = []
    date_pattern = re.compile(r'(\d{2}).-.-')  # Match DD-MM-YYYY or DD.MM.YYYY
    
    for col in df.columns:
        col_str = str(col)
        # Method 1: Column is a string with date format
        if date_pattern.search(col_str):
            date_columns.append(col)
        # Method 2: Column is a datetime object
        elif isinstance(col, datetime) or isinstance(col, pd.Timestamp):
            date_columns.append(col)
    
    print(f"\nIdentified date columns: {date_columns}")
    
    # Try to find our target date column - prioritize current month
    target_column = None
    
    # Method 1: Handle case where dates are parsed as datetime by pandas
    for col in df.columns:
        if isinstance(col, pd.Timestamp) or isinstance(col, datetime):
            if col.month == current_month:
                target_column = col
                print(f"Found matching month in datetime column: {col}")
                break
    
    # Method 2: Look for string representations
    if not target_column:
        month_pattern = f"-{current_month:02d}-" if current_month < 10 else f"-{current_month}-"
        for col in date_columns:
            col_str = str(col)
            if month_pattern in col_str:
                target_column = col
                print(f"Found column with current month pattern: {col}")
                break
    
    # Method 3: Specific approach for your Excel format - look for column with specific month number
    if not target_column:
        for col in df.columns:
            col_str = str(col)
            # Trying to match patterns like "31-05-2025" where 05 is the month
            match = date_pattern.search(col_str)
            if match and match.group(2) == f"{current_month:02d}":
                target_column = col
                print(f"Found exact date column via regex: {col}")
                break
    
    # Fallback: If current month not found, try to find any date column
    if not target_column and date_columns:
        target_column = date_columns[-1]  # Take the last date column as fallback
        print(f"Couldn't find current month column, using last date column: {target_column}")
    
    # Final fallback: Hardcoded column for testing (from your second image)
    if not target_column:
        # Look for a column that has "31-05-2025" or similar anywhere in its name
        for col in df.columns:
            col_str = str(col)
            if "31-05-2025" in col_str:
                target_column = col
                print(f"Found hardcoded fallback column: {col}")
                break
            
    if not target_column:
        print("WARNING: Could not automatically identify the date column.")
        print("Available columns are:", df.columns.tolist())
        # Ask user to specify the column
        col_idx = int(input("Please enter the index number of the availability column from the list above: "))
        target_column = df.columns[col_idx]
        print(f"Using user-selected column: {target_column}")

    
    # Select and rename columns
    cols = df[['Current Role', 'Region', 'Associate ID', 'Associate Name', target_column]].copy()
    cols.rename(columns={target_column: 'Current Availability'}, inplace=True)
    
    # Clean and standardize roles
    def clean_role(x):
        if pd.isna(x):
            return "UNDEFINED"
        
        x = str(x).strip().upper()
        
        if x in ('PROJECT MGR', 'PROJECT MGR/PGM'):
            return "PM"
        elif x == 'PROGRAM MGR' or x == 'PGM':
            return "PGM"
        elif x == 'TPDL' or x == 'TECHNICAL PDL':
            return "TPDL ROLE"
        elif x == 'SCRUM MASTER':
            return "SCRUM MASTER"
        return x
    
    cols['Current Role'] = cols['Current Role'].apply(clean_role)
    
    # Convert Current Availability to numeric format (0-100)
    def convert_availability(x):
        if pd.isna(x):
            return 0
        
        if isinstance(x, str):
            # Remove percentage sign if present
            clean_str = x.replace('%', '').strip()
            try:
                return float(clean_str)
            except:
                return 0
        
        # If the value is decimal (0.X), convert to percentage
        if isinstance(x, (int, float)):
            if 0 < x <= 1:
                return x * 100
            return x
        return 0

    
    cols['Current Availability'] = cols['Current Availability'].apply(convert_availability)
    
    # Add Avail Bucket column based on Current Availability
    def avail_bucket(x):
        if x >= 76:
            return '76 – 100%'
        elif x >= 51:
            return '51 – 75%'
        elif x >= 26:
            return '26 – 50%'
        else:
            return '0 – 25%'
    
    cols['Avail Bucket'] = cols['Current Availability'].apply(avail_bucket)
    
    # Reorder columns to match desired format
    cols = cols[['Current Role', 'Region', 'Avail Bucket', 'Associate ID', 'Associate Name', 'Current Availability']]
    
    # Sort by Current Availability in descending order
    cols = cols.sort_values('Current Availability', ascending=False)
    
    # Export to Excel
    current_month_name = datetime.now().strftime('%B')
    output_file = os.path.join(os.path.dirname(file_path), f'PMS_Dash_{current_month_name}.xlsx')
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write the data
        cols.to_excel(writer, sheet_name='PMS_PM_Resourse', index=False)
        
        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['PMS_PM_Resourse']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#4472C4',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # Apply formats to the header and cells
        for col_num, value in enumerate(cols.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        # Set column widths
        worksheet.set_column('A:A', 15)  # Current Role
        worksheet.set_column('B:B', 10)  # Region
        worksheet.set_column('C:C', 12)  # Avail Bucket
        worksheet.set_column('D:D', 12)  # Associate ID
        worksheet.set_column('E:E', 25)  # Associate Name
        worksheet.set_column('F:F', 18)  # Current Availability
        
        # Set alternate row colors (light blue for the data rows)
        for row_num in range(1, len(cols) + 1):
            cell_format.set_bg_color('#B8CCE4' if row_num % 2 == 0 else '#DCE6F1')
            for col_num in range(len(cols.columns)):
                worksheet.write(row_num, col_num, cols.iloc[row_num-1, col_num], cell_format)
    
    print(f"Report successfully generated: {output_file}")
    return output_file

# Usage
if __name__ == "__main__":
    file_path = input("Enter the path to your Excel file: ")
    try:
        output_path = generate_resource_allocation_report(file_path)
        print(f"Success! Output saved to: {output_path}")
    except Exception as e:
        print(f"Error: {str(e)}")