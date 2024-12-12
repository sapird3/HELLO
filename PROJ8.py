# PROJ 8
# DS

# Retireve the second column of collective action associated with the first column of errors

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages3.xlsx")
ws = wb.sheets['Error and Corrective Action']

actions_dict = {}

last_row = ws.range("B" + str(ws.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for value in range(1, last_row + 1):
    cell = "A" + str(value)  # Assuming the errors are in column A

    # Selecting data
    error = ws.range(cell).value
    action = ws.range("B" + str(value)).value

    if error and action:  # Check if both error and action are not empty
        actions_dict[error] = action  # Store the action associated with the error

# Output
print("\nList of Corrective Actions:")
for error, action in actions_dict.items():
    print(f"Error: {error}\n\tAction: {action}")