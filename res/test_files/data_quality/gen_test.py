import pandas as pd
import numpy as np

# Data quality issues dataset
data = {
    'Product_ID': ['A1', 'B2', 123, 'D4', None, 'F6', 'G7', np.nan, 'I9', 'J10'],
    'Stock_Count': [100, '200', 'invalid', 400, -999999, 600, '7OO', '800 ', None, 1000],
    'Price': [10.999, '20.00', 30, '40.0000', 50.5, '60', None, '80.0', 90.99999, '100.'],
    'Last_Updated': [None, '2024-13-01', '01/01/2024', '2024.01.01', 'yesterday', 
                     '01-01-2024 ', '2024/01/01', '1st Jan 2024', pd.Timestamp.now(), '2024-01-01']
}

df = pd.DataFrame(data)
df.to_excel('data_quality_issues.xlsx', index=False)

