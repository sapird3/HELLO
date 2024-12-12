# Write a quick script that can read info from a single-column Excel sheet. 
# It would be ideal if the Python script could also store the Excel data in a list.
# DS 6/10/2024

# data from excel
#import xlwings as xw
 
# Specifying a sheet
#ws = xw.Book("Book2.xlsx").sheets['Sheet1']
 
# Selecting data from
#v1 = ws.range("A1:A12").value
#print("Result: ", v1)



import pandas as pd

df = pd.read_excel('Node_Health_Project_Log_Messages.xlsx', 'Only Error Messages')

print(df)