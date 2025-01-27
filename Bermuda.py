import pandas as pd
import numpy as np
from datetime import time


file_path = r'C:\Users\skshu\OneDrive\Desktop\Bermuda.xlsx'
output_file_path = r'C:\Users\skshu\OneDrive\Desktop\Cleaned_Bermuda_Final.xlsx'

data = pd.read_excel(file_path, sheet_name="Sheet1")

data = data.iloc[2:].reset_index(drop=True)
data = data.dropna(how="all")

data = data.dropna(how='all', axis=1)

if len(data.columns) == 4:  
    data.columns = ['When', 'What', 'CardNum', 'Who']
else:
    print("The number of remaining columns doesn't match the names provided. Adjust accordingly.")

def standardize_and_replace(column):
    last_date = None
    for i, value in column.items():
        try:
            if isinstance(value, str) and '/' in value:  
                last_date = pd.to_datetime(value).date()
                column.at[i] = last_date
            elif isinstance(value, time): 
                column.at[i] = last_date  
            else:
                column.at[i] = last_date if last_date else None 
        except Exception:
            column.at[i] = None 
    return column

data['When'] = standardize_and_replace(data['When'])

data = data.dropna(subset=['Who'])

data['Where'] = 'Bermuda'

data.to_excel(output_file_path, index=False)

print(f"Cleaned file saved to: {output_file_path}")
