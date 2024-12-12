# PROJ 18
# DS
# 7/8/2024

# FINAL FINAL ??
# trying to add path to excel file - NOT WORKING

# import systems
import os

# user input to extract folder path
print('Hi! So you have a BUNCH of logs to parse through for one little malfunction, huh? Well, you\'re in luck! Because I built this just for you!')
print(r'Please input the path of your log file below in the following syntax: C:\\Users\\(YOUR USERNAME)\\Downloads\\(ZIP FILE NAME)\\LOGS_(MM)_(DD)_(YR)_(HR)_(MIN)_(SEC)')
user_input = input('\nEnter node folder path to parse for MALFUNCTION_IRRECOVERABLE: ')
# C:\\Users\\sapird3\\Downloads\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10
# C:\\Users\\sapird3\\Downloads\\LOGS_07_02_24_12_20_39\\LOGS_07_02_24_12_20_39

# system log
log = user_input + '\\ORTI\\6system_monitor.log' # Enter 6system_monitor.log file 
nodeList = []  # Use a set to store malfunctioning node names

# specify data and split to lines
with open(log, 'r') as file:
    lines = file.readlines()
    
    for row in lines:
        if 'MALFUNCTION_IRRECOVERABLE' in row:  # Corrected the spelling
            index = row.find('SubsystemName:')

            if index != -1:  # Check if 'SubsystemName:' is found
                nodeFound = row[index + len('SubsystemName:'):].strip()  # Adjusted the index to extract the node name
                nodeFound = nodeFound.replace('ProcessManagerState:MALFUNCTION_IRRECOVERABLE', '').strip()
                if nodeFound not in nodeList:  # Check if node is already in the list
                    nodeList.append(nodeFound)  # Add to set of nodes to search through any and all malfunctioning nodes
        
# output
print('\nName of node where MALFUNCTINON_IRRECOVERABLE was found:', nodeList) # print to visualize malfunctioning node(s)

# check each node file
node_errors = {}  # Dictionary to store errors per node
for node in nodeList:
    
    # new log
    log = user_input +'\\'+ node +'\\process_manager.log'

    # specify data and split to lines
    with open(log, 'r') as file:
        lines = file.readlines()
        
        hash_list1 = []  # List to store hash values from log file
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

        # Store the hash list for the current node
        node_errors[node] = hash_list1


# data from excel
import xlwings as xw

# Specify the Excel file
wb_name = input('\nEnter the name of Excel file: ') # Node_Health_Project_Log_Messages4
ws_name = input('Enter the sheet name: ') # Error and Corrective Action

# Determine the path to the desktop
onedrive_path = ('Web Sites\\https://medtronic-my.sharepoint.com\\personal\\sapird3_medtronic_com\\Documents')
# Define the path to the Excel file on the desktop
file_path = os.path.join(onedrive_path, wb_name + '.xlsx')

wb = xw.Book(file_path)
ws = wb.sheets[ws_name]

actions_dict = {}

last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row

# Store hash values from log file
hash_list2 = []

# Read data from the Excel sheet and generate hash values
for value in range(1, last_row + 1):
    error = ws.range('A' + str(value)).value  # Read error message
    action = ws.range('B' + str(value)).value  # Read corrective action

    if error and action:  # Check if both error and action are not empty
        # Extract the relevant part of the error message
        error = error.split(']')[-1].strip()  # Split the string at ']' and take the last part

        # Generate hash values
        hash_value = hash(error)
        actions_dict[hash_value] = (error, action)
        hash_list2.append(hash_value)

# MATCHING

print('\n\nMatching Errors Found in Key-Value Store WITH CORRECTIVE ACTION:')
node_matching_errors = {node: [] for node in nodeList}  # Dictionary to store matching errors per node

for node, hash_list1 in node_errors.items():
    for hash_value in hash_list2:
        if hash_value in hash_list1:
            node_matching_errors[node].append(hash_value)

# Print matching errors for each node
for node, matching_errors in node_matching_errors.items():
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
    print('Within '+ node +' node:')
    for hash_value in matching_errors:
        if hash_value in actions_dict:
            error_message, corrective_action = actions_dict[hash_value]
            print('\nHash:\t', hash_value)
            print('Error:\t', error_message)
            print('Action:\t', corrective_action)
print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')