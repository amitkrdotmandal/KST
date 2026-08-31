import gspread
import pandas as pd

gc = gspread.service_account('kiran-order-take-ad21d600aed8.json')

sh = gc.open("kiran_order_take")
worksheet = sh.worksheet("Sheet1")
# Read entire sheet
data = worksheet.get_all_records()

# First row = column names
df = pd.DataFrame(data)



print(df['Status'].head(20))
df.to_excel('output.xlsx', index=False)