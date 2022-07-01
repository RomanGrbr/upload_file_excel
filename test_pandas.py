import pandas as pd
import json

sales = pd.read_excel('mytest.xlsx', sheet_name=None)

data = {}
for key, value in sales.items():
    data[key] = value

# print(pd.DataFrame.to_json(path_or_buf=data))
print(data)