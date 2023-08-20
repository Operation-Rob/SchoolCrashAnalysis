import pandas as pd

def excel_to_csv_subsets(file_name, sheet_ranges):
    """
    Convert specific ranges from specific sheets of an Excel file to separate CSVs.

    Parameters:
    - file_name (str): The path to the Excel file.
    - sheet_ranges (dict): A dictionary where keys are sheet names and values are
                           tuples of (start_row, end_row, start_col, end_col).

    Example:
    sheet_ranges = {
        'Sheet1': (1, 10, 'A', 'E')
    }
    """
    
    xls = pd.ExcelFile(file_name)
    
    for sheet, (start_row, end_row, start_col, end_col) in sheet_ranges.items():
        data = pd.read_excel(xls, sheet_name=sheet, usecols=f"{start_col}:{end_col}", skiprows=start_row-1, nrows=end_row - start_row + 1)
        output_file_name = f"{sheet}_{start_row}_{end_row}_{start_col}_{end_col}.csv"
        data.to_csv(output_file_name, index=False)
        print(f"Saved {output_file_name}")

# Specification based on your requirements
sheet_ranges_to_export = {
    'BITRE_Fatality_Count_By_Date': (3, 12633, 'A', 'E'),
    'BITRE_Fatality': (5, 55052, 'A', 'W')
}

sheet_ranges_2 = {
    'BITRE_Fatal_Crash_Count_By_Date': (3, 12633, 'A', 'E'),
    'BITRE_Fatal_Crash': (5, 49624, 'A', 'T')
}


excel_to_csv_subsets('ardd_fatalities_jul2023_updated.xlsx', sheet_ranges_to_export)
excel_to_csv_subsets('ardd_fatal_crashes_jul2023_updated.xlsx', sheet_ranges_2)
