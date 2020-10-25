import pandas as pd

# Create a Pandas dataframe from the data.
df = pd.DataFrame({'Data': ['Its\\na bum\\nwrap', 20, 30, 'Test is test', 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

workbook  = writer.book
worksheet = writer.sheets['Sheet1']

wrap_format = workbook.add_format({'text_wrap': True})

worksheet.set_column('B:B', 20, wrap_format)

# Close the Pandas Excel writer and output the Excel file.
writer.save()