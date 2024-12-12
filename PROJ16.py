# PROJ 16
# DS
# 7/2/2024

# IMPLEMENT USER INPUT into FINAL

# import systems
import os

# user input to extract folder path
print('Hi! So you have a BUNCH of logs to parse through for one little malfunction, huh? Well, you\'re in luck! Because I built this just for you!')
print(r'Please input the path of your log file below in the following syntax: C:\\Users\\(YOUR USERNAME)\\Downloads\\(ZIP FILE NAME)\\LOGS_(MM)_(DD)_(YR)_(HR)_(MIN)_(SEC)')
user_input = input('\nEnter node folder path to parse for MALFUNCTION_IRRECOVERABLE: ')
# C:\\Users\\sapird3\\Downloads\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10

# system log
log = user_input + '\\ORTI\\6system_monitor.log'
nodeFound = ''  # List to store hash values from log file

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
log = user_input +'\\'+ nodeFound +'\\process_manager.log'
hash_list1 = []  # List to store hash values from log file

# specify data and split to lines
with open(log, 'r') as file:
    lines = file.readlines()
    
    for row in lines:
        if '[info]' in row:  # check if line contains "[info]" string
            index = row.find('[info]') + 7
            res = (row[index:]).strip()
            hash_list1.append(hash(res))  # Append hash value to the list
        elif '[trace]' in row:  
            index = row.find('[trace]') + 8
            res = (row[index:]).strip()
            hash_list1.append(hash(res))  # Append hash value to the list
        elif '[error]' in row: 
            index = row.find('[error]') + 8
            res = (row[index:]).strip()
            hash_list1.append(hash(res))  # Append hash value to the list


# data from excel
import xlwings as xw

# Specify the Excel file
wb = xw.Book("Node_Health_Project_Log_Messages3.xlsx")
ws = wb.sheets['Error and Corrective Action']

actions_dict = {}

last_row = ws.range("B" + str(ws.cells.last_cell.row)).end('up').row

# Store hash values from log file
hash_list2 = []

# Read data from the Excel sheet and generate hash values
for value in range(1, last_row + 1):
    error = ws.range("A" + str(value)).value  # Read error message
    action = ws.range("B" + str(value)).value  # Read corrective action

    if error and action:  # Check if both error and action are not empty
        # Extract the relevant part of the error message
        error = error.split(']')[-1].strip()  # Split the string at ']' and take the last part

        # Generate hash values
        hash_value = hash(error)
        actions_dict[hash_value] = error
        hash_list2.append(hash_value)

# MATCHING

matching_errors = []

for hash_value in hash_list2:
    if hash_value in hash_list1:
        matching_errors.append(hash_value)

# Print matching errors
print("\n\nMatching Errors Found in Key-Value Store WITH CORRECTIVE ACTION:")
for hash_value in matching_errors:
    if hash_value in actions_dict:
        error_message = actions_dict[hash_value]
        corrective_action = ws.range("B" + str(list(actions_dict.values()).index(error_message) + 1)).value
        print("\nHash:\t", hash_value)
        print("Error:\t", error_message)
        print("Action:\t", corrective_action)

    