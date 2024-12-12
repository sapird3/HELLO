# Project for Adam Calvert
# BY DANA SAPIR
# 07/08/2024

# Input an excel file and output an excel file in a different specific format
# Version 1) Basics of input to output excel

# old input: IRR_Medium_Term_NEW_240617_1621 - full list2.xls
# new input: IRR_Medium_Term_NEW_240710_1645.xls
# NEWEST input: IRR_Medium_Term_NEW_240711_1524.xlsx
# same output: 2024-06-18_19-58-51_Project_Import.xlsx

# import program to grab data from excel
import xlwings as xw

# input excel from user
file = input('Enter the name of the file you would like to input into this program: ') + '.xlsx' 
# IRR_Medium_Term_NEW_240711_1524

# specify excel file & sheet
wb1 = xw.Book(file)
ws1 = wb1.sheets[0]
wb2 = xw.Book()
ws2 = wb2.sheets

print('v/ 1')

# new naming
new_file_name = input('Enter the name you would like for the new excel file (without .xlsx): ') 
# 2024-06-18_19-58-51_Project_Import_DS_
ws2[0].name = 'Project'
ws2.add(name='Log', after=ws2[0])
ws2.add(name='Summary', after=ws2[1])

# find last row of input excel
last_row = ws1.range('A' + str(ws1.cells.last_cell.row)).end('up').row

print('v/ 2')

# SHEET [0] - Project

# setting up titles in new excel
ws2[0].range('A1').value = ['Selector', 'Unique ID', 'UniqueIDSuccessors', 'UniqueIDPredecessors', 'Index', 'Parent Feature', 
                            'Key', 'Task Name', 'Issue Type', 'Status', 'ART', 'Resource Names', 'Task Type', 'Source', 
                            'Σ Total SP', 'Σ Done SP', 'Focus Duration', 'Low-Risk Duration', 'Remaining SP', 
                            '% Complete', 'Σ Realistic Estimate', 'Σ Worst Case Estimate', 'RefinementLevel']
ws2[0].range('A1:W1').api.Font.Bold = True  # Apply bold font to the header row

# collect necessary info
sum_done_sp = 0
row = 0
for value in range(2, last_row + 1):  # Start from 2 to skip the header row
    row += 1
    selector = ''
    task_name = ''
    resource_names = ''
    unique_id = row
    unique_id_succ = ws1.range('C' + str(value)).value
    unique_id_pred = ws1.range('D' + str(value)).value
    index = ws1.range('A' + str(value)).value
    summary = ws1.range('E' + str(value)).value
    parent_feature = ws1.range('K' + str(value)).value
    key = ws1.range('B' + str(value)).value
    issue_type = ws1.range('F' + str(value)).value
    status = ws1.range('G' + str(value)).value
    art = ws1.range('J' + str(value)).value
    
    # steps to retrieve name from art scrum team column
    if art is not None:
        part1 = art.upper()
        part2 = part1.split('-')
        if len(part2) >= 2:
            part3 = part2[1].strip()
            resource_names = part3.replace(' ', '')
            selector = f'{key}_{resource_names}'
            task_name = f'[{selector}] | {summary}'
    
    #task_type = ws1.range('' + str(value)).value
    #source = ws1.range('' + str(value)).value
    refinement_lvl = ws1.range('H' + str(value)).value
    
    if status != 'Cancelled':
        sum_total_sp = ws1.range('N' + str(value)).value
        realistic_est = ws1.range('L' + str(value)).value
        worst_case_est = ws1.range('M' + str(value)).value
    
    if status == 'Done':
        sum_done_sp += sum_total_sp
    if sum_total_sp is not None:
        remaining_sp = sum_total_sp - sum_done_sp

    if issue_type == 'Feature':
        child_or_parent = 'Parent'
        selector = key
        task_name = f'[{selector}] | {summary}'

        if refinement_lvl == realistic_est:  # AM I USING THE CORRECT VALUES FOR THESE CALCULATIONS??
            focus_duration = realistic_est
        if refinement_lvl == sum_total_sp:
            low_risk_duration = sum_total_sp * 1.2  # 20% increase beyond duration

        if status == 'Done':
            perc_complete = '100%'
        else:
            perc_complete = '0%'
    
    else:
        child_or_parent = 'Child'

        focus_duration = '1 hr'
        low_risk_duration = '1 hr'

        if status == 'Done':
            perc_complete = '100%'
        elif refinement_lvl == realistic_est:
            perc_complete = '0%'
        elif status != 'Cancelled':
            if sum_total_sp != 0:
                perc_complete = sum_done_sp / sum_total_sp
                perc_complete = f'{perc_complete}%'


    # organize into new excel
    ws2[0].range('A' + str(value)).value = selector
    ws2[0].range('B' + str(value)).value = unique_id
    ws2[0].range('C' + str(value)).value = unique_id_succ
    ws2[0].range('D' + str(value)).value = unique_id_pred
    ws2[0].range('E' + str(value)).value = index
    ws2[0].range('F' + str(value)).value = parent_feature
    ws2[0].range('G' + str(value)).value = key
    ws2[0].range('H' + str(value)).value = task_name
    ws2[0].range('I' + str(value)).value = issue_type
    ws2[0].range('J' + str(value)).value = status
    ws2[0].range('K' + str(value)).value = art
    ws2[0].range('L' + str(value)).value = resource_names
    ws2[2].range('A' + str(value)).value = resource_names
    #Ws2[0].range('M' + str(value)).value = task_type
    #ws2[0].range('N' + str(value)).value = source
    ws2[0].range('O' + str(value)).value = sum_total_sp
    ws2[0].range('P' + str(value)).value = sum_done_sp
    ws2[0].range('Q' + str(value)).value = focus_duration
    ws2[0].range('R' + str(value)).value = low_risk_duration
    ws2[0].range('S' + str(value)).value = remaining_sp
    ws2[0].range('T' + str(value)).value = perc_complete
    ws2[0].range('U' + str(value)).value = realistic_est
    ws2[0].range('V' + str(value)).value = worst_case_est
    ws2[0].range('W' + str(value)).value = refinement_lvl

# Autofit columns and rows
ws2[0].autofit('c')  # Autofit columns
ws2[0].autofit('r')  # Autofit rows

print('v/ 3')

# SHEET [1] - Log

# setting up titles in new excel
ws2[1].range('A1').value = ['DataTimeStamp', 'Event', 'Data']
ws2[1].range('A1:C1').api.Font.Bold = True  # Apply bold font to the header row

# collect necessary info
for value in range(2, last_row + 1):
    data_time_stamp = ws1.range('U' + str(value)).value
    event = ws1.range('V' + str(value)).value
    data = ws1.range('W' + str(value)).value

    # organize into new excel
    ws2[1].range('A' + str(value)).value = data_time_stamp
    ws2[1].range('B' + str(value)).value = event
    ws2[1].range('C' + str(value)).value = data

# Autofit columns and rows
ws2[1].autofit('c')  # Autofit columns
ws2[1].autofit('r')  # Autofit rows

print('v/ 4')

# SHEET [2] - Summary

# setting up titles in new excel
ws2[2].range('A1').value = ['Resource Names', 'Sum FDR']
ws2[2].range('A1:B1').api.Font.Bold = True  # Apply bold font to the header row

# collect necessary info
for value in range(2, last_row + 1):
    resource_names = ws1.range('X' + str(value)).value
    sum_fdr = ws1.range('Y' + str(value)).value

    # organize into new excel
    #ws2[2].range('A' + str(value)).value = resource_names
    #ws2[2].range('B' + str(value)).value = sum_fdr

# Autofit columns and rows
ws2[2].autofit('c')  # Autofit columns
ws2[2].autofit('r')  # Autofit rows

print('v/ 5')

# save to file
wb2.save(new_file_name + '.xlsx')

# open new excel
wb2.app.visible = True

# output new excel to pdf
# ws2.api.PrintOut()

print('v/ 6')