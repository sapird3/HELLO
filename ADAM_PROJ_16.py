# Project #1 for Adam Calvert
# PROJECT IMPORT FROM JIRA
# BY DANA SAPIR
# 08/16/2024

# Input an excel file and output an excel file in a different specific format (Jira --> Microsoft Project)
# Transfer to Git

# Version 16) fix what my code needs to fix to identically match comparing output (for sentiniel):
    # - include child teams in project sheet
    # - fix spacing on unique id succ & pred v/
    # - add text 16 column (some?)
    # - text 17: parent/child task v/
    # - text 18: user stories (ALL) v/
    # - fix headings (text 19-21) and delete columns after % complete v/
    # - ignore enhancements and defects (RAS tickets) v/
    # DATA SHEET


# old input: IRR_Medium_Term_NEW_240617_1621 - full list2.xls
# new input: IRR_Medium_Term_NEW_240710_1645.xls
# new #2 input: IRR_Medium_Term_NEW_240711_1524.xlsx
# new #3 input: IRR_Medium_Term_NEW_240718_1045.xls
# new #4 input: IRR_Medium_Term_NEW_240730_1640.xls
# NEWEST input: IRR_Medium_Term_NEW_240731_1150.xls
# test input: Sentinel_Hugo_Produc_240802_1103 (Adam + with RefinementLevel).xls
# MAPPING input: Team Mapping for Jira Project converter - Copy.xlsx
# old team's output: 2024-06-18_19-58-51_Project_Import.xlsx

# import program to grab data from excel
import xlwings as xw
import time  # to time how long it takes script to run

def process_excel(data_file, new_file_name, perc_est, mapping_file):
    start_time = time.time()

    # Initialize variables with default values
    sum_total_sp = 0
    sum_done_sp = 0
    perc_complete = '0%'
    feature_key = ''
    realistic_est = 0
    worst_case_est = 0
    feature_status = ''
    feature_ref_level = ''
    remaining_sp = 0
    selector = ''
    task_name = ''
    unique_id_succ = ''
    unique_id_pred = ''
    task_type = ''
    
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

    # setting up titles in new excel
    ws2[0].range('A1').value = ['Selector', 'Unique ID', 'UniqueIDSuccessors', 'UniqueIDPredecessors', 'Index', 'Parent Feature', 
                                'Key', 'Task Name', 'Issue Type', 'Status', 'ART', 'Resource Names', 'Task Type', 'Source', 
                                'Σ Total SP', 'Σ Done SP', 'Focus Duration', 'Low-Risk Duration', 'Remaining SP', 
                                '% Complete']
    ws2[0].range('A1:W1').api.Font.Bold = True  # Apply bold font to the header row

    ws2[1].range('A1').value = ['DataTimeStamp', 'Event', 'Data', 'Selector', 'Unique ID Successors NOT FOUND', 'Unique ID Predecessors NOT FOUND']
    ws2[1].range('A1:F1').api.Font.Bold = True  # Apply bold font to the header row

    ws2[2].range('A1').value = ['Resource Names', 'Sum FDR']
    ws2[2].range('A1:B1').api.Font.Bold = True  # Apply bold font to the header row

    # First pass: collect necessary info
    data = []
    feature_teams = set()
    feature_key_list = []
    feature_unique_id_list = []
    feature_children = {}
    summary_resource_names = []

    current_value = 0
    updated_value = 0
    summary_row = 2
    summary_total = 0
    unique_id_counter = 1
    key_to_unique_id = {}

    for row in range(2, data_last_row + 1):  # Start from 2 to skip the header row
        issue_type = ws1.range('F' + str(row)).value
        if issue_type is None:  # Skip empty rows
            continue
        elif issue_type != 'Feature':  # Skip child rows of Done Features
            #if feature_status == 'Done':
            #    continue
            if issue_type == 'Enhancement':
                continue  # Skip row if 'Enhancment'
            if issue_type == 'Defect':
                continue  # Skip row if 'Defect' FX, Initiative, SFX, 


        status = ws1.range('G' + str(row)).value
        if status == 'Cancelled':
            continue  # Skip the row if the status is 'Cancelled'
        
        unique_id = unique_id_counter
        unique_id_succ = ws1.range('C' + str(row)).value
        unique_id_pred = ws1.range('D' + str(row)).value
        succ_list = [value.strip() for value in unique_id_succ.split(',')] if unique_id_succ else []
        pred_list = [value.strip() for value in unique_id_pred.split(',')] if unique_id_pred else []
        index = ws1.range('A' + str(row)).value
        summary = ws1.range('E' + str(row)).value
        parent_feature = ws1.range('L' + str(row)).value
        key = ws1.range('B' + str(row)).value
        art = ws1.range('K' + str(row)).value
        refinement_lvl = ws1.range('H' + str(row)).value
        resource_names = ''

        if issue_type == 'Feature':
            feature_teams = set()
            feature_key = key
            feature_key_list.append(feature_key)
            feature_unique_id_list.append(unique_id)
            feature_children[feature_key] = set()
            feature_status = status
            feature_ref_level = refinement_lvl

            sum_total_sp = ws1.range('O' + str(row)).value
            sum_done_sp = ws1.range('P' + str(row)).value
            realistic_est = ws1.range('M' + str(row)).value
            worst_case_est = ws1.range('N' + str(row)).value

            # Ensure numeric values
            sum_total_sp = sum_total_sp if sum_total_sp is not None else 0
            sum_done_sp = float(sum_done_sp) if sum_done_sp is not None else 0.0
            realistic_est = realistic_est if realistic_est is not None else 0
            worst_case_est = worst_case_est if worst_case_est is not None else 0

        if art is not None:
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

            if resource_names not in feature_teams:
                feature_teams.add(resource_names)
            else:
                add_total_sp = ws1.range('O' + str(row)).value
                if sum_total_sp is not None and add_total_sp is not None:
                    sum_total_sp += add_total_sp
                excel_value3 = ws1.range('M' + str(row)).value
                if realistic_est is not None and excel_value3 is not None:
                    realistic_est += excel_value3
                excel_value4 = ws1.range('N' + str(row)).value
                if worst_case_est is not None and excel_value4 is not None:
                    worst_case_est += excel_value4
                continue

        if sum_total_sp is not None:
            remaining_sp = sum_total_sp - sum_done_sp

        focus_duration = '1 hr'
        low_risk_duration = '1 hr'

        if issue_type == 'Feature':
            selector = key
            task_name = f'[{selector}] | {summary}'

            if status == 'Done':
                perc_complete = '100%'
            elif sum_total_sp == 0:
                perc_complete = '0%'
            else:
                perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
            
            task_type = 'Parent Task'
        
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
                low_risk_duration = sum_total_sp * perc_est
            
            if status == 'Done':
                perc_complete = '100%'
            elif sum_total_sp == 0:
                perc_complete = '0%'
            else:
                perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
            
            if int(float(perc_complete.rstrip('%'))) > 100: 
                perc_complete = '100%'
            
            if unique_id_succ is None:
                unique_id_succ = feature_key
                succ_list.append(feature_key)
            else:
                if feature_key not in unique_id_succ:
                    unique_id_succ = f'{unique_id_succ},{feature_key}'
                    succ_list.append(feature_key)
            
            task_type = 'Child Task'
        
        if resource_names != '':
            if resource_names not in summary_resource_names:
                summary_resource_names.append(resource_names)
                ws2[2].range('A' + str(summary_row)).value = resource_names
                ws2[2].range('B' + str(summary_row)).value = focus_duration
                summary_total += focus_duration
                summary_row += 1
            else:
                summary_index = summary_resource_names.index(resource_names) + 2
                current_value = int(ws2[2].range('B' + str(summary_index)).value)
                updated_value = current_value + int(focus_duration)
                ws2[2].range('B' + str(summary_index)).value = updated_value
                summary_total += updated_value

        if feature_key and (issue_type, resource_names) not in feature_children[feature_key]:
            feature_children[feature_key].add((issue_type, resource_names))
            if issue_type == 'User Story':
                continue
        else:
            continue

        data.append({
            'selector': selector,
            'unique_id': unique_id,
            'succ_list': succ_list,
            'pred_list': pred_list,
            'index': index,
            'parent_feature': parent_feature,
            'key': key,
            'task_name': task_name,
            'issue_type': issue_type,
            'status': status,
            'art': art,
            'resource_names': resource_names,
            'task_type': task_type,
            'sum_total_sp': sum_total_sp,
            'sum_done_sp': sum_done_sp,
            'focus_duration': focus_duration,
            'low_risk_duration': low_risk_duration,
            'remaining_sp': remaining_sp,
            'perc_complete': perc_complete,
            'realistic_est': realistic_est,
            'worst_case_est': worst_case_est,
            'refinement_lvl': refinement_lvl
        })

        key_to_unique_id[row] = unique_id_counter  # Map the original row number to the new unique ID
        unique_id_counter += 1  # Increment the counter for the next unique ID

    print('Data gathered.')
    processing_time = (time.time() - start_time)/60

    key_to_unique_id = {item['key']: item['unique_id'] for item in data}
    log_row = 1

    for item in data:
        unique_id_succ = ''
        for id in item['succ_list']:
            if id in key_to_unique_id:
                id = key_to_unique_id[id]
                if unique_id_succ == '':
                    unique_id_succ = id
                else:
                    unique_id_succ = f'{unique_id_succ},{id}'
            else:
                unique_id_succ = ''
                log_row += 1
                ws2[1].range('D' + str(log_row)).value = selector
                ws2[1].range('E' + str(log_row)).value = id

        item['unique_id_succ'] = unique_id_succ

        unique_id_pred = ''
        for id in item['pred_list']:
            if id in key_to_unique_id:
                id = key_to_unique_id[id]
                if unique_id_pred == '':
                    unique_id_pred = id
                else:
                    unique_id_pred = f'{unique_id_pred},{id}'
            else:
                unique_id_pred = ''
                log_row += 1
                ws2[1].range('D' + str(log_row)).value = selector
                ws2[1].range('F' + str(log_row)).value = id
        item['unique_id_pred'] = unique_id_pred
    
    # SUMMARY SHEET
    ws2[2].range('A' + str(summary_row)).value = 'Total FDR:'
    ws2[2].range('B' + str(summary_row)).value = summary_total
    ws2[2].range('A' + str(summary_row) +':B' + str(summary_row)).api.Font.Bold = True  # Apply bold font to the total row
    
    # PROJECT SHEET
    proj_row = 2
    for item in data:
        ws2[0].range('A' + str(proj_row)).value = item['selector']
        ws2[0].range('B' + str(proj_row)).value = item['unique_id']
        ws2[0].range('C' + str(proj_row)).value = item['unique_id_succ']
        ws2[0].range('D' + str(proj_row)).value = item['unique_id_pred']
        ws2[0].range('E' + str(proj_row)).value = item['index']
        ws2[0].range('F' + str(proj_row)).value = item['parent_feature']
        ws2[0].range('G' + str(proj_row)).value = item['key']
        ws2[0].range('H' + str(proj_row)).value = item['task_name']
        ws2[0].range('I' + str(proj_row)).value = item['issue_type']
        ws2[0].range('J' + str(proj_row)).value = item['status']
        ws2[0].range('K' + str(proj_row)).value = item['art']
        ws2[0].range('L' + str(proj_row)).value = item['resource_names']
        ws2[0].range('M' + str(proj_row)).value = item['task_type']
        ws2[0].range('N' + str(proj_row)).value = 'User Stories'
        ws2[0].range('O' + str(proj_row)).value = item['sum_total_sp']
        ws2[0].range('P' + str(proj_row)).value = item['sum_done_sp']
        ws2[0].range('Q' + str(proj_row)).value = item['focus_duration']
        ws2[0].range('R' + str(proj_row)).value = item['low_risk_duration']
        ws2[0].range('S' + str(proj_row)).value = item['remaining_sp']
        ws2[0].range('T' + str(proj_row)).value = item['perc_complete']
        proj_row += 1

    # LOG SHEET
    ws2[1].range('A2').value = time.time()
    ws2[1].range('A3').value = time.time()
    ws2[1].range('A4').value = time.time()
    ws2[1].range('A5').value = time.time()
    ws2[1].range('A6').value = time.time()
    ws2[1].range('A7').value = time.time()
    ws2[1].range('A8').value = time.time()
    ws2[1].range('B2').value = 'Start:'
    ws2[1].range('B3').value = 'Input File:'
    ws2[1].range('B4').value = 'Output File:'
    ws2[1].range('B5').value = 'Info'
    ws2[1].range('B6').value = 'Processing:'
    ws2[1].range('B7').value = 'Result:'
    ws2[1].range('B8').value = 'Processing Complete:'
    ws2[1].range('C2').value = 'DAN SAP Formatter v1.6'
    ws2[1].range('C3').value = data_file
    ws2[1].range('C4').value = new_file_name
    ws2[1].range('C5').value = mapping_file
    ws2[1].range('C6').value = f'Jira Data: Rows: {data_last_row}, Cols: 16'
    ws2[1].range('C7').value = f'Project Data: Rows: {proj_row - 2}, Cols: 20'
    ws2[1].range('C8').value = f'{processing_time} seconds'

    ws2[0].range('A1').value = ['Text10', 'Unique ID', 'Unique ID Successors', 'Unique ID Predecessors', 'Text11', 'Text12', 
                                'Text13', 'Name', 'Text14', 'Text15', 'Text16', 'Resource Names', 'Text17', 'Text18', 
                                'Text19', 'Text20', 'Duration', 'Duration1', 'Text21', '% Complete']

    print('v/ 3')

    ws2[0].autofit('c')
    ws2[0].autofit('r')
    ws2[1].autofit('c')
    ws2[1].autofit('r')
    ws2[2].autofit('c')
    ws2[2].autofit('r')

    wb2.save(new_file_name + '.xlsx')

    wb2.app.visible = True

    print('v/ 4')

    end_time = time.time()
    duration_sec = end_time - start_time
    duration_min = duration_sec / 60

    print(f"Script execution time: {duration_sec:.2f} seconds or {duration_min:.2f} minutes")

if __name__ == "__main__":

    # input DATA excel from user
    data_file = input('Enter the name of the data file you would like to input into this program (without .xls): ') + '.xls' 
    # Sentinel_Hugo_Produc_240802_1103 (Adam + with RefinementLevel)
    
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
        mapping_file = input('Enter the name of the mapping team converter file you would like to input into this program (without .xlsx): ') + '.xlsx' 
        # Team Mapping for Jira Project converter - Copy
    else:
        mapping_file = ''

    process_excel(data_file, new_file_name, perc_est, mapping_file)
