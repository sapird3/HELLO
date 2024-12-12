# Project for Adam Calvert
# BY DANA SAPIR
# 07/25/2024

# Input an excel file and output an excel file in a different specific format
# Version 8) updating features and fixing calculations

# old input: IRR_Medium_Term_NEW_240617_1621 - full list2.xls
# new input: IRR_Medium_Term_NEW_240710_1645.xls
# new #2 input: IRR_Medium_Term_NEW_240711_1524.xlsx
# NEWEST input: IRR_Medium_Term_NEW_240718_1045.xls
# MAPPING input: Team Mapping for Jira Project converter.xlsx
# old team's output: 2024-06-18_19-58-51_Project_Import.xlsx

# import program to grab data from excel
import xlwings as xw

def process_excel(data_file, new_file_name, perc_est, mapping_file):
    # specify excel file & sheet
    wb1 = xw.Book(data_file)
    ws1 = wb1.sheets[0]
    wb2 = xw.Book()
    ws2 = wb2.sheets
    if mapping_file != '':
        wb3 = xw.Book(mapping_file)
        ws3 = wb3.sheets[0]

    print('v/ 1')

    # new naming
    ws2[0].name = 'Project'
    ws2.add(name='Log', after=ws2[0])
    ws2.add(name='Summary', after=ws2[1])

    # find last row of input excels
    data_last_row = ws1.range('A' + str(ws1.cells.last_cell.row)).end('up').row
    if mapping_file != '':
        map_last_row = ws3.range('A' + str(ws3.cells.last_cell.row)).end('up').row

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
    new_row = 2  # Start from 2 to skip the header row

    # Dictionary to track teams under each feature
    feature_teams = {}
    feature_key_list = []
    feature_unique_id_list = []
    feature_status = ''
    feature_ref_level = ''
    # Dictionary to track children and feature estimations under each feature
    feature_children = {}

    for row in range(2, data_last_row + 1):  # Start from 2 to skip the header row
        issue_type = ws1.range('F' + str(row)).value
        if issue_type is None:  # Skip empty rows
            continue
        elif issue_type != 'Feature':  # Skip child rows of Done Features
            if feature_status == 'Done':
                continue

        status = ws1.range('G' + str(row)).value
        if status == 'Cancelled':
            continue  # Skip the row if the status is 'Cancelled'

        selector = ''
        task_name = ''
        resource_names = ''
        unique_id = new_row - 1  # Adjust to start unique ID from 1
        unique_id_succ = ws1.range('C' + str(row)).value
        if unique_id_succ is not None:
            succ_list = [value.strip() for value in unique_id_succ.split(',')]
        unique_id_pred = ws1.range('D' + str(row)).value
        if unique_id_pred is not None:
            pred_list = [value.strip() for value in unique_id_pred.split(',')]
        index = ws1.range('A' + str(row)).value
        summary = ws1.range('E' + str(row)).value
        parent_feature = ws1.range('K' + str(row)).value
        key = ws1.range('B' + str(row)).value
        art = ws1.range('J' + str(row)).value
        refinement_lvl = ws1.range('H' + str(row)).value

        # Reset team names for new feature
        if issue_type == 'Feature':
            feature_teams = set()
            feature_key = key
            feature_key_list.append(feature_key)
            feature_unique_id_list.append(unique_id)
            feature_children[feature_key] = set()
            feature_status = status
            feature_ref_level = refinement_lvl

            sum_total_sp = ws1.range('N' + str(row)).value
            sum_done_sp = ws1.range('O' + str(row)).value
            realistic_est = ws1.range('L' + str(row)).value
            worst_case_est = ws1.range('M' + str(row)).value

        # steps to retrieve name from art scrum team column
        if art is not None:
            # Map the art to the consolidated team name
            if mapping_file != '':
                for map_row in range(2, map_last_row + 1):
                    original_team = ws3.range('A' + str(map_row)).value
                    mapped_team = ws3.range('B' + str(map_row)).value
                    if art == original_team:
                        art = mapped_team

            part1 = art.upper()
            if '-' in part1:
                part2 = part1.split('-')
                if len(part2) >= 2:
                    part3 = part2[1].strip()
                    resource_names = part3.replace(' ', '')
            else:
                resource_names = part1.replace(' ', '')

            # Check if the team already exists in this feature
            if resource_names not in feature_teams:
                feature_teams.add(resource_names)
            else:
                excel_value1 = ws1.range('N' + str(row)).value
                if sum_total_sp is not None and excel_value1 is not None:
                    sum_total_sp += excel_value1
                excel_value2 = ws1.range('O' + str(row)).value
                if sum_done_sp is not None and excel_value2 is not None:
                    sum_done_sp += excel_value2
                excel_value3 = ws1.range('L' + str(row)).value
                if realistic_est is not None and excel_value3 is not None:
                    realistic_est += excel_value3
                excel_value4 = ws1.range('M' + str(row)).value
                if worst_case_est is not None and excel_value4 is not None:
                    worst_case_est += excel_value4
                continue

            selector = f'{feature_key}_{resource_names}'
            task_name = f'[{selector}] | {summary}'
                
        if status == 'Done':
            sum_done_sp += sum_total_sp
        if sum_total_sp is not None:
            remaining_sp = sum_total_sp - sum_done_sp

        if issue_type == 'Feature':
            focus_duration = '1 hr'
            low_risk_duration = '1 hr'

            selector = key
            task_name = f'[{selector}] | {summary}'

            if status == 'Done':
                perc_complete = '100%'
            else:
                perc_complete = '0%'
        else:
            if issue_type == 'Feature Estimation':
                selector = f'{feature_key}_{resource_names}'
                task_name = f'[{selector}] | {summary}'
                feature_teams.add(resource_names)

            refinement_lvl = feature_ref_level
            if refinement_lvl == 'Feature Estimation':
                focus_duration = realistic_est
                low_risk_duration = worst_case_est
            elif refinement_lvl == 'Story Points':
                focus_duration = sum_total_sp
                low_risk_duration = sum_total_sp * perc_est  # 20% (or 1.2x) increase beyond duration

            if status == 'Done':
                perc_complete = '100%'
            elif refinement_lvl == realistic_est:
                perc_complete = '0%'
            elif status != 'Cancelled':
                if sum_total_sp != 0:
                    perc_complete = sum_done_sp / sum_total_sp
                    perc_complete = f'{perc_complete}%'
            
            # Add parent's Unique ID as a Successor
            if unique_id_succ is None:
                unique_id_succ = feature_key
                succ_list.append(feature_key)
            else:
                if feature_key not in unique_id_succ:
                    unique_id_succ = f'{unique_id_succ}, {feature_key}'
                    succ_list.append(feature_key)

        # Check if the child or feature estimation already exists under this feature
        if (issue_type, resource_names) not in feature_children[feature_key]:
            feature_children[feature_key].add((issue_type, resource_names))
            if issue_type == 'User Story':
                continue
        else:
            continue
        
        # Substitute Unique IDs for Successors and Predecessors
        unique_id_succ = ''
        for id in succ_list:
            if id in feature_unique_id_list:
                index = feature_unique_id_list.index(id)
                id = feature_key_list[index]
            else:
                unique_id_succ = f'{id} not found!!!'
                continue
            if unique_id_succ == '':
                unique_id_succ = id
            else:
                unique_id_succ = f'{unique_id_succ}, {id}'
        if unique_id_pred is not None:
            unique_id_pred = ''
            for id in pred_list:
                if id in feature_unique_id_list:
                    index = feature_unique_id_list.index(id)
                    id = feature_key_list[index]
                else:
                    unique_id_pred = f'{id} not found!!!'
                    continue
                if unique_id_pred == '':
                    unique_id_pred = id
                else:
                    unique_id_pred = f'{unique_id_succ}, {id}'

        # organize into new excel
        ws2[0].range('A' + str(new_row)).value = selector
        ws2[0].range('B' + str(new_row)).value = unique_id
        ws2[0].range('C' + str(new_row)).value = unique_id_succ
        ws2[0].range('D' + str(new_row)).value = unique_id_pred
        ws2[0].range('E' + str(new_row)).value = index
        ws2[0].range('F' + str(new_row)).value = parent_feature
        ws2[0].range('G' + str(new_row)).value = key
        ws2[0].range('H' + str(new_row)).value = task_name
        ws2[0].range('I' + str(new_row)).value = issue_type
        ws2[0].range('J' + str(new_row)).value = status
        ws2[0].range('K' + str(new_row)).value = art
        ws2[0].range('L' + str(new_row)).value = resource_names
        ws2[2].range('A' + str(new_row)).value = resource_names
        #ws2[0].range('M' + str(new_row)).value = task_type
        #ws2[0].range('N' + str(new_row)).value = source
        ws2[0].range('O' + str(new_row)).value = sum_total_sp
        ws2[0].range('P' + str(new_row)).value = sum_done_sp
        ws2[0].range('Q' + str(new_row)).value = focus_duration
        ws2[0].range('R' + str(new_row)).value = low_risk_duration
        ws2[0].range('S' + str(new_row)).value = remaining_sp
        ws2[0].range('T' + str(new_row)).value = perc_complete
        ws2[0].range('U' + str(new_row)).value = realistic_est
        ws2[0].range('V' + str(new_row)).value = worst_case_est
        ws2[0].range('W' + str(new_row)).value = refinement_lvl

        new_row += 1  # Increment the row counter only if the row is not skipped

    # Autofit columns and rows
    ws2[0].autofit('c')  # Autofit columns
    ws2[0].autofit('r')  # Autofit rows
    #ws2[0].api.Rows(2).Select()
    #ws2[0].api.Application.ActiveWindow.FreezePanes = True  # Freeze top row

    # Renaming columns into text that Microsoft Project understands
    # ws2[0].range('A1').value = ['Selector', 'Unique ID', 'UniqueIDSuccessors', 'UniqueIDPredecessors', 'Index', 'Parent Feature', 
    # 'Key', 'Task Name', 'Issue Type', 'Status', 'ART', 'Resource Names', 'Task Type', 'Source', 
    # 'Text19', 'Text20', 'Duration', 'Duration1', 'Remaining SP', 
    # '%Complete', 'Text22', 'Text23', 'Text24']

    print('v/ 3')

    # SHEET [1] - Log

    # setting up titles in new excel
    ws2[1].range('A1').value = ['DataTimeStamp', 'Event', 'Data']
    ws2[1].range('A1:C1').api.Font.Bold = True  # Apply bold font to the header row

    # collect necessary info
    for row in range(2, data_last_row + 1):
        data_time_stamp = ws1.range('U' + str(row)).value
        event = ws1.range('V' + str(row)).value
        data = ws1.range('W' + str(row)).value

        # organize into new excel
        ws2[1].range('A' + str(row)).value = data_time_stamp
        ws2[1].range('B' + str(row)).value = event
        ws2[1].range('C' + str(row)).value = data

    # Autofit columns and rows
    ws2[1].autofit('c')  # Autofit columns
    ws2[1].autofit('r')  # Autofit rows
    #ws2[1].api.Rows(2).Select()
    #ws2[1].api.Application.ActiveWindow.FreezePanes = True  # Freeze top row

    print('v/ 4')

    # SHEET [2] - Summary

    # setting up titles in new excel
    ws2[2].range('A1').value = ['Resource Names', 'Sum FDR']
    ws2[2].range('A1:B1').api.Font.Bold = True  # Apply bold font to the header row

    # collect necessary info
    for row in range(2, data_last_row + 1):
        sum_fdr = ws1.range('Y' + str(row)).value

        # organize into new excel
        ws2[2].range('B' + str(row)).value = sum_fdr

    # Autofit columns and rows
    ws2[2].autofit('c')  # Autofit columns
    ws2[2].autofit('r')  # Autofit rows
    #ws2[2].api.Rows(2).Select()
    #ws2[2].api.Application.ActiveWindow.FreezePanes = True  # Freeze top row

    print('v/ 5')

    # save to file
    wb2.save(new_file_name + '.xlsx')

    # open new excel
    wb2.app.visible = True

    # output new excel to pdf
    # ws2.api.PrintOut()

    print('v/ 6')

if __name__ == "__main__":

    # input DATA excel from user
    data_file = input('Enter the name of the data file you would like to input into this program: ') + '.xls' 
    # IRR_Medium_Term_NEW_240718_1045

    # new naming
    new_file_name = input('Enter the name you would like for the new excel file (without .xlsx): ') 
    # 2024-06-18_19-58-51_Project_Import_DS_

    # input percentage estimation
    perc_est = input('Enter the percentage of worst-case estimation from realistic (without ''%''): ')
    perc_est = int(perc_est)/100 + 1 # convert to decimal and add one to multiply duration time
    # 20

    # input MAPPING excel from user (optional)
    ask = input('Do you have a mapping team converter file to input? (Y/N) ')
    if ask == 'Y' or ask == 'y' or ask == 'yes' or ask == 'Yes' or ask == 'YES':
        mapping_file = input('Enter the name of the mapping team converter file you would like to input into this program: ') + '.xlsx' 
        # Team Mapping for Jira Project converter
    else:
        mapping_file = ''

    process_excel(data_file, new_file_name, perc_est, mapping_file)