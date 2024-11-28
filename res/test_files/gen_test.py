import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Create test data
data = {
    'ID': [1, 2, 3, '4', 5, 6, 7, 8, 9, 10],
    'Date': [
        '2024-01-01',
        '01/02/2024',
        'Invalid Date',
        '2024-01-04',
        '05-01-2024',
        '2024/01/06',
        datetime.now(),
        '2024-01-08 ',  # Extra space
        '2024-01-09',
        None
    ],
    'Amount': [
        1000.50,
        '1,200.75',
        'abc',
        -999999999,  # Outlier
        1500.0000,  # Extra decimals
        1600.50,
        None,
        ' 1800.25 ',  # Leading/trailing spaces
        1900.50,
        2000
    ],
    'Category': [
        'Sales',
        'Sales ',  # Trailing space
        'SALES',  # Inconsistent case
        'Marketing',
        'Marketing',  # Duplicate
        None,
        'Finance',
        'HR',
        'HR',
        ''  # Empty string
    ],
    'Calculation': [
        '=A2/0',  # #DIV/0!
        '=VLOOKUP("missing",A1:B1,2)',  # #N/A
        '=INVALID_FUNCTION()',  # #NAME?
        '=A1',  # Valid
        None,
        '=REF!',  # #REF!
        '=VALUE("abc")',  # #VALUE!
        '=1+1',
        '=SUM()',  # Empty sum
        '=NULL'  # #NULL!
    ],
    'Percentage': [
        '50%',
        '0.75',  # Inconsistent format
        '80.00%',
        '.90',  # Missing leading zero
        '100%',
        None,
        '120%',  # Outlier
        '60.5%',
        '70%',
        '65'  # Missing % symbol
    ]
}

df = pd.DataFrame(data)

# Add duplicate column
df['Category_2'] = df['Category']

# Create Excel writer
writer = pd.ExcelWriter('test_data.xlsx', engine='xlsxwriter')

# Write visible sheet
df.to_excel(writer, sheet_name='Main', index=False)

# Write hidden sheet with duplicates
df.to_excel(writer, sheet_name='Hidden', index=False)

# Get workbook and worksheet
workbook = writer.book
worksheet = writer.sheets['Main']

# Add some formatting inconsistencies
format1 = workbook.add_format({'num_format': '#,##0.00'})
format2 = workbook.add_format({'num_format': '0.0'})

worksheet.set_column('C:C', 15, format1)
worksheet.write('C4', 1234.56, format2)

writer.close()