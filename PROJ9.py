# PROJ 9
# DS

# output corrective action associated with inputed hash value from error message
# NOT WORKING

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages3.xlsx")
ws = wb.sheets['Error and Corrective Action']

key_value_store = {}
actions_dict = {}

last_row = ws.range("B" + str(ws.cells.last_cell.row)).end('up').row

# Read data from the Excel sheet and generate hash values
for value in range(1, last_row + 1):
    error = ws.range("A" + str(value)).value  # Read error message
    action = ws.range("B" + str(value)).value  # Read corrective action

    if error and action:  # Check if both error and action are not empty
        # Extract the relevant part of the error message
        error = error.split(']')[-1].strip()  # Split the string at ']' and take the last part
        actions_dict[error] = action  # Store the action associated with the error

        cell = "A" + str(value)
        x = ws.range(cell).value

        if x:  # Check if cell is not empty
            index = x.find("[info]")  # finding the index of the start of our output
            if index != -1:  # Check if "[info]" is found
                res = (x[index + 7:]).strip()  # Extract the substring after "[info]"
                key = hash(res)  # Calculate hash
                key_value_store[key] = res  # Remove leading/trailing whitespace
            index = x.find("[error]")  # finding the index of the start of our output
            if index != -1:  # Check if "[error]" is found
                res = (x[index + 8:]).strip()  # Extract the substring after "[error]"
                key = hash(res)  # Calculate hash
                key_value_store[key] = res  # Remove leading/trailing whitespace
            index = x.find("[trace]")  # finding the index of the start of our output
            if index != -1:  # Check if "[trace]" is found
                res = (x[index + 8:]).strip()  # Extract the substring after "[trace]"
                key = hash(res)  # Calculate hash
                key_value_store[key] = res  # Remove leading/trailing whitespace
# Output
i = 1
print("\nList of Corrective Actions:")
for error, action in actions_dict.items():
    print(f"Error: {error}\n\tAction: {action}")
    i += 1