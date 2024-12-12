# PROJ 11
# DS
# 6/27/2024

# scan process_manager.log for that malfunctioning node (from proj 10)

# retrieve malfucntioning node from saved file
with open('node_found.txt', 'r') as file:
    nodeFound = file.read()

print("Node found in the 6system_monitor.log file:", nodeFound)
