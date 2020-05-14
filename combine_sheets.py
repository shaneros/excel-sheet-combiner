import pandas as pd

filename = 'input.xlsx'
df = pd.concat(pd.read_excel(filename, sheet_name=None), ignore_index=True)
df.to_excel(r'output.xlsx', index=False)