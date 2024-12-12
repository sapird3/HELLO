# PROJECT INTRO WORK PT 3
# DS 6/11/2024

# Try to read both the first and second workbooks and see if you can find a way to get the script 
# to detect if any hashes in the second workbook match the first.

# data from excel
import xlwings as xw

# Specifying sheets
wb1 = xw.Book("Node_Health_Project_Log_Messages.xlsx")
ws1 = wb1.sheets['Only Error Messages']
wb2 = xw.Book("Book1.xlsx")
ws2 = wb2.sheets['Sheet1']

# FIRST WORKBOOK

key_value_store1 = {}
full1 = ''

# Determine last row
last_row1 = ws1.range("A" + str(ws1.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for num in range(1, last_row1 + 1):
    cell = "A" + str(num)

    # Selecting data
    x = ws1.range(cell).value

    if x:  # Check if cell is not empty
        index = x.find("[info]") + 7  # finding the index of the start of our output
        res = ""  # declaring empty output var

        # loop to delete "fluff"
        for i in range(len(x)):
            if i >= index:
                res += x[i]

        key = abs(hash(res))  # Taking absolute value of hash
        key_value_store1[key] = res

        full1 = full1 + '\n' + res
        lines1 = full1.splitlines()

# output
count1 = 1
print("1ST WORKBOOK:")
print("\nFull list:", full1)
print("\nKey Value Store:")
for item in key_value_store1:
    print(item, lines1[count1])
    count1 += 1

# SECOND WORKBOOK

key_value_store2 = {}
full2 = ''

# Determine last row
last_row2 = ws2.range("A" + str(ws1.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for num in range(2, last_row2 + 1):
    cell = "A" + str(num)

    # Selecting data
    x = ws2.range(cell).value

    if x:  # Check if cell is not empty
        res = ""  # declaring empty output var

        # loop to delete "fluff"
        for i in range(len(x)):
            res += x[i]
        
        key = abs(hash(res))  # Taking absolute value of hash
        key_value_store2[key] = res

        full2 = full2 + '\n' + res
        lines2 = full2.splitlines()

# output
count2 = 1
print("\n\n2ND WORKBOOK:")
print("\nFull list:", full2)
print("\nKey Value Store:")
for item in key_value_store2:
    print(item, lines2[count2])
    count2 += 1


# check connections btwn two workbooks

print('\nKnown Errors will appear if applicable:')
for items in key_value_store2:
    for check in key_value_store1:
        if items == check:
            index = list(key_value_store2).index(items) + 1
            print(lines2[index])