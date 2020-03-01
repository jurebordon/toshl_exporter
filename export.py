import pandas as pd
import functions as fnc

input_file_path = 'F:/Downloads/Toshl Export.xls'
output_file_path = 'F:/Downloads/ToshlExport_transformed.xlsx'
output_sheet_name = 'ToshlExport'

expenses = fnc.read_df(input_file_path, 'Toshl Expenses')
expenses = fnc.clean_df(expenses, 'ExpenseAmount', 'IncomeAmount')

incomes = fnc.read_df(input_file_path, 'Toshl Incomes')
incomes = fnc.clean_df(incomes, 'IncomeAmount', 'ExpenseAmount')

frames = [expenses, incomes]
result = pd.concat(frames)

# Create a Pandas Excel writer using XlsxWriter as the engine.
# Also set the default datetime and date formats.
writer = pd.ExcelWriter(output_file_path,
                        engine='xlsxwriter',
                        datetime_format='yyyy-mm-dd')

# Convert the dataframe to an XlsxWriter Excel object.
result.to_excel(writer, index=False, sheet_name=output_sheet_name)

# Get the xlsxwriter workbook and worksheet objects in order to set the column
# widths, to make the dates clearer.
workbook  = writer.book
worksheet = writer.sheets[output_sheet_name]
num_format = workbook.add_format({'num_format': '#,##0.00'})

worksheet.set_column('A:B', 15)
worksheet.set_column('C:D', 20)
worksheet.set_column('E:F', 15, num_format)
worksheet.set_column('G:G', 40)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
