# check 2: LOG TO LOG WORKS!

# one error log

key_value_store1 = {}

# specify data and split to lines
with open('one_error.log', 'r') as file1:
    lines1 = file1.readlines()
    #line_count = len(lines)  # determine number of rows
    #last_row = lines[-1]  # determine last row
    
    for row1 in lines1:
        if '[info]' in row1:  # check if line contains "[info]" string
            index = row1.find('[info]') + 7
            res = (row1[index:]).strip()
            key = hash(res)  
            key_value_store1[key] = res  # remove leading/trailing whitespace

# output
print("\nONE_ERROR.LOG Key Value Store:")
for key, value in key_value_store1.items():
    print(key, value)


# timeout error log

key_value_store2 = {}

# specify data and split to lines
with open('timeout_error.log', 'r') as file2:
    lines2 = file2.readlines()
    #line_count = len(lines)  # determine number of rows
    #last_row = lines[-1]  # determine last row
    
    for row2 in lines2:
        if '/einstein' in row2:
            continue
        elif '[info]' in row2:  # check if line contains "[info]" string
            index2 = row2.find('[info]') + 7
            res2 = (row2[index2:]).strip()
            key2 = hash(res2)  
            key_value_store2[key2] = res2  # remove leading/trailing whitespace

# output
print("\nTIMEOUT_ERROR.LOG Key Value Store:")
for key2, value2 in key_value_store2.items():
    print(key2, value2)

    
# check known errors
print('\nKnown Errors will appear if applicable:')
for items2 in key_value_store2:
    for check in key_value_store1:
        if items2 == check:
            #index2 = list(key_value_store2).index(items2) + 1
            print(items2, key_value_store2[items2])
