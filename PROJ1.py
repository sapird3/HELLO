# PROJECT INTRO WORK PT 1
# DS 6/10/2024

# data from excel
import xlwings as xw

# random number generator
import numpy as np 
 
# Specifying a sheet
ws = xw.Book("Node_Health_Project_Log_Messages.xlsx").sheets['Logs and Common Errors']

# Selecting data from
c1 = ws.range("A2:A18").value
c2 = ws.range("B2:B18").value
c3 = ws.range("C2:C18").value
c4 = ws.range("D2:D18").value

# for loop
y = np.random.randint(2,18)
for x in range(y, 0, -y):
    print("\n", c1[x], "\n", c2[x], "\n", c3[x], "\n", c4[x])