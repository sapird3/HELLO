# PROJECT INTRO WORK PT 2
# DS 6/10/2024

# data from excel
import xlwings as xw

# Specifying a sheet
wb = xw.Book("Node_Health_Project_Log_Messages.xlsx")
ws = wb.sheets['Only Error Messages']

key_value_store = {}
full = ''

# Determine last row
last_row = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for num in range(1, last_row + 1):
    cell = "A" + str(num)

    # Selecting data
    x = ws.range(cell).value

    if x:  # Check if cell is not empty
        index = x.find("[info]") + 7  # finding the index of the start of our output
        res = ""  # declaring empty output var

        # loop to delete "fluff"
        for i in range(len(x)):
            if i >= index:
                res += x[i]

        key = abs(hash(res))  # Taking absolute value of hash
        key_value_store[key] = res

        full = full + '\n' + res

# output
print("Key Value Store:", key_value_store)
print("Full list:", full)