# DS PROJECT 7
# look for the known errors in a process_manager.log from any node folder

import os
import json

# system log

hash_list1 = []  # List to store hash values from log file
log = 'C:\\Users\\sapird3\\Downloads\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10\\ORTI\\process_manager.log'

# specify data and split to lines
with open(log, 'r') as file1:
    lines1 = file1.readlines()
    
    for row1 in lines1:
        if '[info]' in row1:  # check if line contains "[info]" string
            index1 = row1.find('[info]') + 7
            res1 = (row1[index1:]).strip()
            hash_list1.append(hash(res1))  # Append hash value to the list
        elif '[trace]' in row1:  
            index1 = row1.find('[trace]') + 8
            res1 = (row1[index1:]).strip()
            hash_list1.append(hash(res1))  # Append hash value to the list
        elif '[error]' in row1: 
            index1 = row1.find('[error]') + 8
            res1 = (row1[index1:]).strip()
            hash_list1.append(hash(res1))  # Append hash value to the list

# output
print("\nList of Hash Values from Log File:")
print(hash_list1)


# Compare with key-value store


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

# Compare with key-value store
matching_errors = set()

for hash_value in hash_list1:
    if hash_value in key_value_store2:
        matching_errors.add(hash_value)

# Print matching errors
print("\nMatching Errors Found in Key-Value Store:")
for hash_value in matching_errors:
    print(hash_value, key_value_store2[hash_value])