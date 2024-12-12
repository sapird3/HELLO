# check 3: SMALL LOG TO EXCEL WORKS!!!

# one error log

key_value_store1 = {}

# specify data and split to lines
with open('one_error.log', 'r') as file1:
    lines1 = file1.readlines()
    
    for row1 in lines1:
        if '[info]' in row1:  # check if line contains "[info]" string
            index = row1.find('[info]') + 7
            res = row1[index:].strip()
            key = hash(res)  
            key_value_store1[key] = res  # remove leading/trailing whitespace

# output
print("\nONE_ERROR.LOG Key Value Store:")
for key, value in key_value_store1.items():
    print(key, value)


# excel

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages.xlsx")
ws = wb.sheets['Only Error Messages']

key_value_store2 = {}

last_row1 = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for num in range(1, last_row1 + 1):
    cell = "A" + str(num)

    # Selecting data
    x = ws.range(cell).value

    if x:  # Check if cell is not empty
        index2 = x.find("[info]")  # finding the index of the start of our output
        if index2 != -1:  # Check if "[info]" is found
            res2 = (x[index2 + 7:]).strip()  # Extract the substring after "[info]"
            key2 = hash(res2)  # Calculate SHA-256 hash
            key_value_store2[key2] = res2  # Remove leading/trailing whitespace
        
# Output
print("\nEXCEL Error Messages Key Value Store:")
for key, value in key_value_store2.items():
    print(key, value)



# check known errors !!!

print('\nKnown Errors will appear if applicable:')
for items2 in key_value_store2:
    for check in key_value_store1:
        if items2 == check:
            print(items2, key_value_store2[items2])
