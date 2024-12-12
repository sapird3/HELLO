# PROJ 21
# DS
# 7/11/2024

# TRYING TO ADD NEWEST FEATURE: UNDEFINED INOPERABLE STATE

# import systems
import os
import xlwings as xw

# user input to extract folder path
print('System in UNDEFINED')
print('Please run the grab_logs.sh script, unzip the folder, and enter the full path to the folder.')
user_input = input('\nEnter node folder path to parse for UNDEFINED: ')
# C:\\Users\\sapird3\\Downloads\\Safety_Fail_No_App_Img\\LOGS_02_05_24_11_56_00

# system log
log = os.path.join(user_input, 'ORTI', '6system_monitor.log')  # Enter 6system_monitor.log file 
nodeList = []  # Use a list to store malfunctioning node names

# specify data and split to lines
with open(log, 'r') as file:
    lines = file.readlines()
    
    for row in lines:
        if 'UNDEFINED' in row:
            index = row.find('SubsystemName:')

            if index != -1:  # Check if 'SubsystemName:' is found
                nodeFound = row[index + len('SubsystemName:'):].strip()  # Adjusted the index to extract the node name
                nodeFound = nodeFound.replace('ProcessManagerState:UNDEFINED', '').strip()
                if nodeFound not in nodeList:  # Check if node is already in the list
                    nodeList.append(nodeFound)  # Add to list of nodes to search through any and all malfunctioning nodes
        
# output
print('Name of node where UNDEFINED was found:', nodeList)  # print to visualize malfunctioning node(s)
print()

# check each node file
node_errors = {}  # Dictionary to store errors per node
node_lines = {}  # Dictionary to store log lines per node
skip_node = []
for node in nodeList:
    
    # new log
    log = os.path.join(user_input, node, 'process_manager.log')

    # specify data and split to lines
    if os.path.exists(log):
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
    else:
        print(f"The file '{log}' does not exist.")
        skip_node.append(node)


# Specify the Excel file
wb = xw.Book('Node_Health_Project_Log_Messages.xlsx')
sheet_names = [sheet.name for sheet in wb.sheets]

for sheet_name in sheet_names:
    if sheet_name == 'Error and Corrective Action':
        ws = wb.sheets[sheet_name]

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

        print('\nKnown Errors Found and Suggested Corrective Actions')
        node_matching_errors = {node: [] for node in nodeList}  # Dictionary to store matching errors per node

        for node, hash_list1 in node_errors.items():
            for hash_value in hash_list2:
                if hash_value in hash_list1:
                    node_matching_errors[node].append(hash_value)

        printed_items = set()

        import re

        # Print matching errors for each node 
        for node, matching_errors in node_matching_errors.items():
            if node not in skip_node:
                print('---------------------------------------------------------------------------------------')
                print('Within '+ node +' node:')
                for hash_value in matching_errors:
                    if hash_value in actions_dict:
                        error_message, corrective_action = actions_dict[hash_value]
                        print('\nError:\t', error_message)
                        print('Action:\t', corrective_action)
                        
                        pattern = r"'(.*?)'"
                        # Search for the pattern in the error message
                        match = re.search(pattern, error_message)
                        if match:
                            # Extract the text within single quotes
                            scanning = match.group(1)
                            if scanning in error_message:
                                ask = input(f'\t Would you like to scan the {scanning} log (Y/N)? ')
                                if ask == 'Y' or ask == 'y' or ask == 'yes' or ask == 'Yes' or ask == 'YES':
                                    print(f'\t\t Scanning {scanning}...')
                                    
                                    def find_file_with_num(root_dir):
                                        num = 1

                                        # Walk through the directory and its subdirectories
                                        for root, dirs, files in os.walk(root_dir):
                                            while True:
                                                # Convert num to string for comparison
                                                num_str = str(num)
                                                
                                                # Check each file in the current directory
                                                for file in files:
                                                    if num_str in file:
                                                        # If num is found in the file name, return the full path
                                                        return os.path.join(root, file)
                                                
                                                # Increment num and continue the search
                                                num += 1
                                    
                                    matching_file = find_file_with_num(f'{user_input}\\{node}')

                                    with open(fr'{matching_file}', 'r') as file:

                                        ws_scanning = wb.sheets[scanning]
                                        actions_dict_scanning = {}
                                        last_row_scanning = ws_scanning.range('A' + str(ws_scanning.cells.last_cell.row)).end('up').row

                                        # Store hash values from log file
                                        hash_list_scanning = []

                                        # Read data from the Excel sheet and generate hash values
                                        for value in range(1, last_row + 1):
                                            error_scanning = ws_scanning.range('A' + str(value)).value  # Read error message
                                            action_scanning = ws_scanning.range('B' + str(value)).value  # Read corrective action

                                            if error_scanning and action_scanning:  # Check if both error and action are not empty
                                                # Extract the relevant part of the error message
                                                error_scanning = error_scanning.split(']')[-1].strip()  # Split the string at ']' and take the last part

                                                # Generate hash values
                                                hash_value_scanning = hash(error_scanning)
                                                actions_dict_scanning[hash_value_scanning] = (error_scanning, action_scanning)
                                                hash_list_scanning.append(hash_value_scanning)

                                                # matching
                                                node_matching_errors_scanning = {node_scanning: [] for node_scanning in nodeList}  # Dictionary to store matching errors per node

                                                for node_scanning, hash_list_scanning in node_errors.items():
                                                    for hash_value_scanning in hash_list_scanning:
                                                        if hash_value_scanning in hash_list_scanning:
                                                            node_matching_errors_scanning[node_scanning].append(hash_value_scanning)

                                                printed_items_scanning = set()

                                                # Print matching errors for each node
                                                for node_scanning, matching_errors_scanning in node_matching_errors_scanning.items():
                                                    for hash_value_scanning in matching_errors_scanning:
                                                        if hash_value_scanning in actions_dict_scanning:
                                                            error_message_scanning, corrective_action_scanning = actions_dict_scanning[hash_value_scanning]
                                                            print('\n\t\t\tError:\t', error_message_scanning)
                                                            print('\t\t\tAction:\t', corrective_action_scanning)
                                else:
                                    print(f'\t\tDeclined further scanning of {scanning} log.')
                        
                        print('\n\t For more info, please refer to the succeeding lines after the found error:')
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