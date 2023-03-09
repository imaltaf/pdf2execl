import pandas as pd

# read the bank statement text file
with open('bank_statement.txt', 'r') as f:
    data = f.read().splitlines()

# create an empty DataFrame to store the data
df = pd.DataFrame(columns=['Date', 'Description', 'Amount'])

# loop through the data and extract information
for line in data:
    # check if the line starts with a date in the format of 'mm/dd/yyyy'
    if line[:10].count('/') == 2:
        # extract the date and description
        date = line[:10]
        description = line[10:].strip()
    else:
        # extract the amount
        amount = float(line.strip())
        # add a new row to the DataFrame
        df = df.append({'Date': date, 'Description': description, 'Amount': amount}, ignore_index=True)

# create a new column 'Amount Color' to store the color codes
df['Amount Color'] = ''
# loop through the DataFrame and assign color codes to credit amounts
for index, row in df.iterrows():
    if row['Amount'] > 0:
        df.loc[index, 'Amount Color'] = 'green'

# save the DataFrame to an Excel file
writer = pd.ExcelWriter('bank_statement.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
# add conditional formatting to highlight credit amounts in green
worksheet = writer.sheets['Sheet1']
worksheet.conditional_format('C2:C{}'.format(len(df)+1), {'type': 'cell',
                                                          'criteria': '>',
                                                          'value': 0,
                                                          'format': workbook.add_format({'bg_color': 'green'})})
writer.save()
