# PROJ 19
# DS
# 7/8/2024

# FINAL FINAL FINAL ???
# printing next few lines for more info
# taken out unnesscary fluff and printing hash
# ADDED FEAUTRE: ask to look into the log of the prcoess that 'died'

# import systems
import os
import xlwings as xw

# user input to extract folder path
print('Hi! So you have a BUNCH of logs to parse through for one little malfunction, huh? Well, you\'re in luck! Because I built this just for you!')
print(r'Please input the full path of your log file using \\ to separate folder names.')
user_input = input('\nEnter node folder path to parse for MALFUNCTION_IRRECOVERABLE: ')
# C:\\Users\\sapird3\\Downloads\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10
# C:\\Users\\sapird3\\Downloads\\LOGS_07_02_24_12_20_39\\LOGS_07_02_24_12_20_39
# C:\\Users\\sapird3\\OneDrive - Medtronic PLC\\ORTI_Malfunction\\LOGS_03_13_24_15_26_10

# system log
log = os.path.join(user_input, 'ORTI', '6system_monitor.log')  # Enter 6system_monitor.log file 
nodeList = []  # Use a list to store malfunctioning node names

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
                    nodeList.append(nodeFound)  # Add to list of nodes to search through any and all malfunctioning nodes
        
# output
print('\nName of node where MALFUNCTION_IRRECOVERABLE was found:', nodeList)  # print to visualize malfunctioning node(s)

# check each node file
node_errors = {}  # Dictionary to store errors per node
node_lines = {}  # Dictionary to store log lines per node
for node in nodeList:
    
    # new log
    log = os.path.join(user_input, node, 'process_manager.log')

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

        # Store the hash list and log lines for the current node
        node_errors[node] = hash_list1
        node_lines[node] = lines


# Specify the Excel file
wb = xw.Book('Node_Health_Project_Log_Messages.xlsx')
ws = wb.sheets['Error and Corrective Action']

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

printed_items = set()
import re

# Print matching errors for each node 
for node, matching_errors in node_matching_errors.items():
    print('---------------------------------------------------------------------------------------')
    print('Within '+ node +' node:')
    for hash_value in matching_errors:
        if hash_value in actions_dict:
            error_message, corrective_action = actions_dict[hash_value]
            print('Error:\t', error_message)
            print('Action:\t', corrective_action)
            
            pattern = r"'(.*?)'"
            # Search for the pattern in the error message
            match = re.search(pattern, error_message)
            if match:
                # Extract the text within single quotes
                scanning = match.group(1)
                if scanning in error_message:
                    ask = input(f'\tWould you like to scan the {scanning} log (Y/N)? ')
                    if ask == 'Y' or ask == 'y' or ask == 'yes' or ask == 'Yes' or ask == 'YES':
                        print(f'\t\tScanning {scanning}.log...')
                        print('\t\t\tHERE')
                        print(fr'\t\t{user_input}\\{node}\\10{scanning}.log')
                    else:
                        print(f'\t\tDeclined further scanning of {scanning} log.')
            
            print('\t For more info, please refer to the succeeding lines after the found error:')
            key_index = node_errors[node].index(hash_value)
            for count in range(1, 4):  # Iterate over the next three lines
                next_index = key_index + count
                
                if next_index < len(node_lines[node]):
                    next_line = node_lines[node][next_index]  # Get the next log line
                    print('\t\t', next_line.strip())

                    # Check if an item has already been printed
                    if next_line in printed_items:
                        break

                    # Create an empty set to keep track of printed items
                    printed_items.add(next_line)
print('---------------------------------------------------------------------------------------')
