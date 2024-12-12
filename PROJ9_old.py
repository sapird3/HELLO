# PROJ 9
# DS

# output corrective action associated with inputed hash value from error message

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages3.xlsx")
ws = wb.sheets['Error and Corrective Action']

actions_dict = {}

last_row = ws.range("B" + str(ws.cells.last_cell.row)).end('up').row

# Store hash values from log file
hash_list = []

# Read data from the Excel sheet and generate hash values
for value in range(1, last_row + 1):
    error = ws.range("A" + str(value)).value  # Read error message
    action = ws.range("B" + str(value)).value  # Read corrective action

    if error and action:  # Check if both error and action are not empty
        # Extract the relevant part of the error message
        error = error.split(']')[-1].strip()  # Split the string at ']' and take the last part

        actions_dict[error] = action  # Store the action associated with the error

        # Generate hash values
        hash_value = hash(error)
        hash_list.append(hash_value)

# Output
i = 1
print("\nList of Corrective Actions:")
for error, action in actions_dict.items():
    print(f"Hash:{hash_list[i]} Error: {error}\n\tAction: {action}")
    i += 1