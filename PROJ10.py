# PROJ 10
# DS

# extract name of malfunctioning node

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
print("\nName of node where MALFUNCTINON_IRRECOVERABLE was found:")
print(nodeFound)

# save to file in order to use in another script
with open('node_found.txt', 'w') as file:
    file.write(nodeFound)