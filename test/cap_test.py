# import pandas
import pandas as pd

data = {'Name': ['A', 'B', 'C', 'D'], 'Age': [10, 0, 30, 50]}

# create the initial DataFrame
df_initial = pd.DataFrame(data)

# add more data
new_data = {'Name': ['ㅇ', 'ㅇ'], 'Age': [25, 40]}
df = pd.concat([df_initial, pd.DataFrame(new_data)], ignore_index=True)

# create a Pandas Excel writer object using XlsxWriter as the engine
writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')

# write data to the Excel sheet
df.to_excel(writer, sheet_name='Sheet1', index=False)

# close file
writer.close()
