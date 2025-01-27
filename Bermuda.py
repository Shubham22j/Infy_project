import pandas as pd
import numpy as np
from datetime import time

# File paths
file_path = r'C:\Users\skshu\OneDrive\Desktop\Bermuda.xlsx'
output_file_path = r'C:\Users\skshu\OneDrive\Desktop\Cleaned_Bermuda_Final.xlsx'

# Load the Excel file
data = pd.read_excel(file_path, sheet_name="Sheet1")

# Drop the top 3 rows
data = data.iloc[2:].reset_index(drop=True)

# Drop rows where all values are NaN
data = data.dropna(how="all")

# Drop completely empty columns
data = data.dropna(how='all', axis=1)

# Rename the remaining columns (adjust as per your needs)
if len(data.columns) == 4:  # Ensure the number of columns matches
    data.columns = ['When', 'What', 'CardNum', 'Who']
else:
    print("The number of remaining columns doesn't match the names provided. Adjust accordingly.")

# Replace timestamps with the last valid date above
def standardize_and_replace(column):
    last_date = None
    for i, value in column.items():
        # Check if the value is a valid date (string format like '12/31/2024')
        try:
            if isinstance(value, str) and '/' in value:  # Parse date strings
                last_date = pd.to_datetime(value).date()
                column.at[i] = last_date
            elif isinstance(value, time):  # Check if it's a time
                column.at[i] = last_date  # Replace time with last valid date
            else:
                column.at[i] = last_date if last_date else None  # Default to the last date
        except Exception:
            column.at[i] = None  # Replace invalid values with None
    return column

data['When'] = standardize_and_replace(data['When'])

# Remove rows where the 'Who' column is empty or NaN
data = data.dropna(subset=['Who'])

# Add a new column 'Where' and populate it with 'Bermuda'
data['Where'] = 'Bermuda'

# Save the cleaned data to a new Excel file
data.to_excel(output_file_path, index=False)

print(f"Cleaned file saved to: {output_file_path}")
