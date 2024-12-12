# Write a small program to read data from Excel file (some random ten rows, two columns of data), 
# store it into a list, and then print it.
# DS 6/10/2024

# data from excel
import xlwings as xw
 
# Specifying a sheet
ws = xw.Book("Book2.xlsx").sheets('Sheet1')

# Selecting data from
v1 = ws.range("A1:A12").value
v2 = ws.range("B1:B12").value

# for loop
for x in range(12):
    print(v1[x], v2[x])