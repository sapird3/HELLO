# PROJECT INTRO WORK PT 4
# DS 6/11/2024

# should be from file on system
# can't get to work, so working on what I can

# BIG LOG TO EXCEL WORKS!!! FINALLL

# timeout error log

key_value_store1 = {}

# specify data and split to lines
with open('timeout_error.log', 'r') as file1:
    lines1 = file1.readlines()
    #line_count = len(lines)  # determine number of rows
    #last_row = lines[-1]  # determine last row
    
    for row1 in lines1:
        if '/einstein' in row1:
            continue
        elif '[info]' in row1:  # check if line contains "[info]" string
            index1 = row1.find('[info]') + 7
            res1 = (row1[index1:]).strip()
            key1 = hash(res1)  
            key_value_store1[key1] = res1  # remove leading/trailing whitespace

# output
print("\nTIMEOUT_ERROR.LOG Key Value Store:")
for key1, value1 in key_value_store1.items():
    print(key1, value1)


# excel

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages.xlsx")
ws = wb.sheets['Only Error Messages']

key_value_store2 = {}

last_row2 = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for num in range(1, last_row2 + 1):
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