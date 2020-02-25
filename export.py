import pandas as pd
import numpy as np

expenses = pd.read_excel('~/Downloads/ToshlExport.xls',
                         sheet_name='Toshl Expenses',
                         header=1,
                         usecols='A:D,G,I',
                         dtypeType={'B': object, 'C': object, 'D': object, 'I': object})

expenses.rename(columns={"In main currency": "ExpenseAmount"}, inplace=True)
expenses['Date'] = pd.to_datetime(expenses['Date'])
expenses['ExpenseAmount'] = pd.to_numeric(expenses['ExpenseAmount'])
expenses['IncomeAmount'] = np.float(0)

expenses = expenses[['Date', 'Account', 'Category', 'Tags', 'ExpenseAmount', 'IncomeAmount', 'Description']]
print(expenses.columns.to_list)

print(expenses.head)

print(sum(expenses.ExpenseAmount))

incomes = pd.read_excel('~/Downloads/ToshlExport.xls',
                        sheet_name='Toshl Incomes',
                        header=1,
                         usecols='A:D,G,I',
                        na_values='No value',
                        dtypeType={'B': object, 'C': object, 'D': object, 'I': object},
                        converters={'A': pd.to_datetime, 'E': pd.to_numeric})

incomes.rename(columns={"In main currency": "IncomeAmount"}, inplace=True)
incomes['ExpenseAmount'] = 0
incomes['Date'] = pd.to_datetime(incomes['Date'])

incomes = incomes[['Date', 'Account', 'Category', 'Tags', 'ExpenseAmount', 'IncomeAmount', 'Description']]

print(incomes.columns.to_list)

print(incomes.head)

print(sum(incomes.IncomeAmount))

frames = [expenses, incomes]

result = pd.concat(frames)



# Create a Pandas Excel writer using XlsxWriter as the engine.
# Also set the default datetime and date formats.
writer = pd.ExcelWriter('/Users/jure/Downloads/ToshlExport_transformed.xlsx',
                        engine='xlsxwriter',
                        datetime_format='yyyy-mm-dd')

# Convert the dataframe to an XlsxWriter Excel object.
result.to_excel(writer, index=False, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects in order to set the column
# widths, to make the dates clearer.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
num_format = workbook.add_format({'num_format': '#,##0.00'})

worksheet.set_column('A:B', 15)
worksheet.set_column('C:D', 20)
worksheet.set_column('E:F', 13, num_format)
worksheet.set_column('G:G', 40)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
