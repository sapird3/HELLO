# PROJ 12
# DS
# 6/27/2024

# BRINGING IT ALL TGTHR (projects 7, 10, 11)

# import systems
import os

# system log
nodeFound = ''  # List to store hash values from log file
log = 'C:\\Users\\sapird3\\Downloads\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10\\ORTI\\6system_monitor.log'

# specify data and split to lines
with open(log, 'r') as file:
    lines = file.readlines()
    
    for row in lines:
        if 'MALFUNCTION_IRRECOVERABLE' in row:  # Corrected the spelling
            index = row.find('SubsystemName:')

            
            if index != -1:  # Check if 'SubsystemName:' is found
                nodeFound = row[index + len('SubsystemName:'):].strip()  # Adjusted the index to extract the node name
                nodeFound = nodeFound.replace('ProcessManagerState:MALFUNCTION_IRRECOVERABLE', '').strip()
        
# output
print("\nName of node where MALFUNCTINON_IRRECOVERABLE was found:", nodeFound)

# new log
log = 'C:\\Users\\sapird3\\Downloads\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10\\'+ nodeFound +'\\process_manager.log'
print(log)
hash_list = []  # List to store hash values from log file

# specify data and split to lines
with open(log, 'r') as file:
    lines = file.readlines()
    
    for row in lines:
        if '[info]' in row:  # check if line contains "[info]" string
            index = row.find('[info]') + 7
            res = (row[index:]).strip()
            hash_list.append(hash(res))  # Append hash value to the list
        elif '[trace]' in row:  
            index = row.find('[trace]') + 8
            res = (row[index:]).strip()
            hash_list.append(hash(res))  # Append hash value to the list
        elif '[error]' in row: 
            index = row.find('[error]') + 8
            res = (row[index:]).strip()
            hash_list.append(hash(res))  # Append hash value to the list

# output
print("\nList of Hash Values from Log File:")
print(hash_list)


# Compare with key-value store 

# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages3.xlsx")
ws = wb.sheets['Only Error Messages']

key_value_store = {}

last_row = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row

# Adjust the loop range to include duplicated cells
for num in range(1, last_row + 1):
    cell = "A" + str(num)

    # Selecting data
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
print("\nEXCEL Error Messages Key Value Store:")
for key, value in key_value_store.items():
    print(key, value)

# Compare with key-value store
matching_errors = set()

for hash_value in hash_list:
    if hash_value in key_value_store:
        matching_errors.add(hash_value)

# Print matching errors
print("\nMatching Errors Found in Key-Value Store:")
for hash_value in matching_errors:
    print(hash_value, key_value_store[hash_value])