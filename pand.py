# PROJECT INTRO WORK PT 1 TRY PANDAS
# DS 6/10/2024

# data from excel
import pandas as pd

# random number generator
import numpy as np 
 
# Specifying a sheet
df = pd.read_excel('Node_Health_Project_Log_Messages.xlsx', 'Logs and Common Errors')
#mylist = df['A1:A2'].tolist()

print(df.head())

# Selecting data from
c1 = df.range("A2:A18").value
c2 = df.range("B2:B18").value
c3 = df.range("C2:C18").value
c4 = df.range("D2:D18").value

# for loop
y = np.random.randint(2,18)
for x in range(y, 0, -y):
    print("\n", c1[x], "\n", c2[x], "\n", c3[x], "\n", c4[x])