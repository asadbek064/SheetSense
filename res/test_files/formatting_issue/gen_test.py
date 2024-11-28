import pandas as pd

data = {
    'Employee': [' John Doe', 'Jane Smith ', '  Bob Jones', 'Alice Brown  ', 'Tom Wilson'],
    'Salary': ['$50,000', '60000', '$70,000.00', '80000.0', '90,000'],
    'Hire_Date': ['01-Jan-2024', '2024/01/02', '03.01.2024', 'Jan 4, 2024', '2024-01-05'],
    'Department': ['IT ', ' HR', 'Finance', '  Marketing ', 'Sales'],
    'Performance': ['95.00%', '87.5', '92', '88.75%', '90.0%']
}

df = pd.DataFrame(data)

writer = pd.ExcelWriter('formatting_issues.xlsx', engine='xlsxwriter')
df.to_excel(writer, index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add inconsistent formatting
formats = [
    workbook.add_format({'num_format': '#,##0'}),
    workbook.add_format({'num_format': '0.00'}),
    workbook.add_format({'num_format': '$#,##0.00'}),
    workbook.add_format({'num_format': '0%'})
]

for col, fmt in enumerate(formats):
    worksheet.set_column(col, col, 15, fmt)

writer.close()
