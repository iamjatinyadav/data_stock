import pandas as pd
import xlrd

symbol = str(input("Enter the Company Symbol Here:-"))
symbol = symbol.upper()
loc = "input.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
for i in range(sheet.nrows):
    check_var = str(sheet.cell_value(i,1))
    modify_var = check_var.strip()
    if modify_var == symbol:
        query_string = f"https://query1.finance.yahoo.com/v7/finance/download/{symbol}.NS?" \
                       f"period1=1626048000&period2=1633996800&interval=1d&events=history&" \
                       f"includeAdjustedClose=true"
        df = pd.read_csv(query_string)
        mod_df = df.drop(columns=['Open','High','Low','Adj Close','Volume'])
        mod_df.to_excel("output.xlsx", index=False)
        break
else:
    print("Please enter correct symbol")

