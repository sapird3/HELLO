# DS PROJECT 5
# print the next 3 lines surrounding the error

# error log

key_value_store1 = {}

# specify data and split to lines
with open('timeout_error.log', 'r') as file1:
    lines1 = file1.readlines()
    line_count1 = len(lines1)  # determine number of rows
    last_row1 = lines1[-1]  # determine last row
    
    for row1 in lines1:
        if '/einstein' in row1:
            continue
        elif '[info]' in row1:  # check if line contains "[info]" string
            index1 = row1.find('[info]') + 7
            res1 = (row1[index1:]).strip()
            key1 = hash(res1)  
            key_value_store1[key1] = res1  # remove leading/trailing whitespace

# output
print("\nERROR LOG Key Value Store:")
for key1, value1 in key_value_store1.items():
    print(key1, value1)


# UPDATED excel

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages3.xlsx")
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
            key2 = hash(res2)  # Calculate hash
            key_value_store2[key2] = res2  # Remove leading/trailing whitespace
        index2 = x.find("[error]")  # finding the index of the start of our output
        if index2 != -1:  # Check if "[error]" is found
            res2 = (x[index2 + 8:]).strip()  # Extract the substring after "[error]"
            key2 = hash(res2)  # Calculate hash
            key_value_store2[key2] = res2  # Remove leading/trailing whitespace
        index2 = x.find("[trace]")  # finding the index of the start of our output
        if index2 != -1:  # Check if "[trace]" is found
            res2 = (x[index2 + 8:]).strip()  # Extract the substring after "[trace]"
            key2 = hash(res2)  # Calculate hash
            key_value_store2[key2] = res2  # Remove leading/trailing whitespace
        
# Output
print("\nEXCEL Error Messages Key Value Store:")
for key, value in key_value_store2.items():
    print(key, value)



# check known errors !!!

keys = list(key_value_store1.keys())
c = 1
printed_items = set()

print('\nKnown Errors will appear if applicable:')
for items2 in key_value_store2:
    for check in key_value_store1:
        if items2 == check:

            key_index = keys.index(items2)
            print('Error #', str(c), ':')
            print(items2, key_value_store2[items2]) #  error
            
            for count in range(1, 4):  # Iterate over the next three keys
                next_index = key_index + count
                
                if next_index < len(keys):
                    next_key = keys[next_index]  # Get the next key
                    next_line = key_value_store1[next_key]  # Get the next line based on the next key
                    print(next_key, next_line)

                    # Check if an item has already been printed
                    if next_key in printed_items:
                        break

                    # Create an empty set to keep track of printed items
                    printed_items.add(next_key)

            c += 1
