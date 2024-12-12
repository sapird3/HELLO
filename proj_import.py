# Project #1 for Adam Calvert
# PROJECT IMPORT FROM JIRA
# BY DANA SAPIR
# 09/20/2024

# Input an excel file and output an excel file in a different specific format (Jira --> Microsoft Project)
# Transfer to Git

# Version 18) HOPEFUL FINAL FINAL

# old input: IRR_Medium_Term_NEW_240617_1621 - full list2.xls
# new input: IRR_Medium_Term_NEW_240710_1645.xls
# new #2 input: IRR_Medium_Term_NEW_240711_1524.xlsx
# new #3 input: IRR_Medium_Term_NEW_240718_1045.xls
# new #4 input: IRR_Medium_Term_NEW_240730_1640.xls
# NEWEST input: IRR_Medium_Term_NEW_240731_1150.xls
# test #1 input: Sentinel_Hugo_Produc_240802_1103 (Adam + with RefinementLevel).xls
# MAPPING #1 input: Team Mapping for Jira Project converter - Copy2.xlsx
# test #2 input: PnP_Categorized_240904_1601 (discipline+refinement level updated).xls
# MAPPING #2 input: Team Mapping for Jira Project converter (1).xlsx
# old team's output: 2024-06-18_19-58-51_Project_Import.xlsx

# import program to grab data from excel
import xlwings as xw
import time  # to time how long it takes script to run

def process_excel(data_file, new_file_name, perc_est, mapping_file):
    """
    Processes the input Excel file and outputs it in a specific format for Microsoft Project.

    Args:
        data_file (str): The path to the input data Excel file.
        new_file_name (str): The desired name for the output Excel file.
        perc_est (float): The percentage estimation for worst-case scenario.
        mapping_file (str): The path to the mapping team converter file (optional).
    """
    start_time = time.time()

    # Initialize variables with default values
    sum_total_sp = 0
    sum_done_sp = 0
    perc_complete = '0%'
    feature_key = ''
    realistic_est = 0
    worst_case_est = 0
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

    print('v/ Checkpoint #1')

    # new naming
    ws2[0].name = 'Project'
    ws2.add(name='Log', after=ws2[0])
    ws2.add(name='Summary', after=ws2[1])

    # find last row of input excels
    data_last_row = ws1.range('A' + str(ws1.cells.last_cell.row)).end('up').row
    if mapping_file != '':
        map_last_row = ws3.range('A' + str(ws3.cells.last_cell.row)).end('up').row

    print('v/ Checkpoint #2')

    # setting up titles in new excel
    ws2[0].range('A1').value = ['Selector', 'Unique ID', 'Unique ID Successors', 'Unique ID Predecessors', 'Index', 'Parent Feature', 
                                'Key', 'Task Name', 'Issue Type', 'Status', 'ART', 'Resource Names', 'Task Type', 'Source', 
                                'Σ Total SP', 'Σ Done SP', 'Focus Duration', 'Low-Risk Duration', 'Remaining SP', 
                                '% Complete']
    ws2[1].range('A1').value = ['DataTimeStamp', 'Event', 'Data', 'Selector', 'Unique ID Successors NOT FOUND', 'Unique ID Predecessors NOT FOUND']
    ws2[2].range('A1').value = ['Resource Names', 'Sum FDR']

    print('v/ Checkpoint #3')

    # First pass: collect necessary info
    data = []
    feature_teams = set()
    feature_key_list = []
    children_key_list = []
    summary_resource_names = []

    current_value = 0
    updated_value = 0
    summary_row = 2
    summary_total = 0
    unique_id_counter = 1
    key_to_unique_id = {}

    for row in range(2, data_last_row + 1):  # Start from 2 to skip the header row
        issue_type = ws1.range('F' + str(row)).value
        if issue_type != 'Feature' and issue_type != 'Feature Estimation' and issue_type != 'Issue' and issue_type != 'User Story':  
            continue  # Skip rows that are not Features, Feature Estimations, Issues, or User Stories (or those that are empty)

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
        art_scrum_team = ws1.range('K' + str(row)).value
        refinement_lvl = ws1.range('H' + str(row)).value
        resource_names = ''
        focus_duration = 0
        low_risk_duration = 0

        # Check for repeated feature & children keys
        if key in feature_key_list: 
            print(f'This feature key was found again: {key}')
            continue 
        elif key in children_key_list: 
            print(f'This child key was found again: {key}')
            continue 
        else:
            if issue_type == 'Feature':
                feature_key_list.append(key)
            else:
                children_key_list.append(key)

        if issue_type == 'Feature':
            feature_teams = set()
            feature_key = key
            feature_ref_level = refinement_lvl

        if issue_type == 'Feature Estimation' or issue_type == 'User Story' or issue_type == 'Issue':
            if mapping_file != '':
                for map_row in range(2, map_last_row + 1):
                    original_team = ws3.range('A' + str(map_row)).value
                    mapped_team = ws3.range('B' + str(map_row)).value
                    if art_scrum_team == original_team:
                        art_scrum_team = mapped_team

            part1 = art_scrum_team.upper()
            if '-' in part1:
                part2 = part1.split('-')
                if len(part2) >= 2:
                    part3 = part2[1].strip()
                    resource_names = part3.replace(' ', '')
            else:
                resource_names = part1.replace(' ', '')

            # If resource_names is not already in feature_teams, add it and collect estimates
            if resource_names not in feature_teams:
                feature_teams.add(resource_names)
                sum_total_sp = ws1.range('O' + str(row)).value
                sum_done_sp = ws1.range('P' + str(row)).value
                realistic_est = ws1.range('M' + str(row)).value
                worst_case_est = ws1.range('N' + str(row)).value
                # Ensure numeric values
                sum_total_sp = float(sum_total_sp) if sum_total_sp is not None else 0.0
                sum_done_sp = float(sum_done_sp) if sum_done_sp is not None else 0.0
                realistic_est = float(realistic_est) if realistic_est is not None else 0.0
                worst_case_est = float(worst_case_est) if worst_case_est is not None else 0.0
            else:
                # Find the relevant entry in the data list
                child_index = -1
                for i, item in enumerate(data):
                    if item['selector'] == f'{feature_key}_{resource_names}':
                        child_index = i
                        break
                
                if child_index == -1:
                    print(f"Error: resource_names '{resource_names}' not found in data.")
                    continue

                # Function to calculate and update the values
                def calculate(column_name, column_letter):
                    old_value = float(data[child_index][column_name])
                    add_value = ws1.range(column_letter + str(row)).value
                    add_value = float(add_value) if add_value is not None else 0.0
                    new_value = add_value + old_value
                    data[child_index][column_name] = new_value

                calculate('sum_total_sp', 'O')
                calculate('sum_done_sp', 'P')
                calculate('realistic_est', 'M')
                calculate('worst_case_est', 'N')

                continue

        # Calculate the remaining story points
        if sum_total_sp is not None:
            remaining_sp = sum_total_sp - sum_done_sp

        if issue_type == 'Feature':
            # For 'Feature' type issues, set the selector and task name
            selector = key
            task_name = f'[{selector}] | {summary}'

            # Default focus and low-risk duration for features
            focus_duration = '1 hr'
            low_risk_duration = '1 hr'

            # Determine the percentage completion based on the status and story points
            if status == 'Done':
                perc_complete = '100%'
            elif sum_total_sp == 0:
                perc_complete = '0%'
            else:
                perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
            
            task_type = 'Parent Task'  # Set task type as 'Parent Task' for features
        
        else:
            # For 'User Story' or 'Issue' type issues
            if issue_type == 'User Story' or issue_type == "Issue":
                selector = f'{feature_key}_{resource_names}'
                task_name = f'[{selector}] | {summary}'
                feature_teams.add(resource_names)

            # Use the feature refinement level if none is provided
            refinement_lvl = feature_ref_level
            if refinement_lvl == 'Feature Estimation':
                focus_duration = realistic_est
                low_risk_duration = worst_case_est
            elif refinement_lvl == 'Story Points':
                focus_duration = sum_total_sp
                low_risk_duration = sum_total_sp * perc_est  # Calculate low-risk duration based on percentage estimation
            elif refinement_lvl == '':
                print("No refinement level found.")
            
            # Determine the percentage completion based on the status and story points
            if status == 'Done':
                perc_complete = '100%'
            elif sum_done_sp == 0:
                perc_complete = '0%'
            else:
                perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
            
            # Ensure the percentage completion does not exceed 100%
            if int(float(perc_complete.rstrip('%'))) > 100: 
                perc_complete = '100%'
            
            # Handle unique ID successors
            if unique_id_succ is None:
                unique_id_succ = feature_key
                succ_list.append(feature_key)
            else:
                if feature_key not in unique_id_succ:
                    unique_id_succ = f'{unique_id_succ},{feature_key}'
                    succ_list.append(feature_key)
            
            task_type = 'Child Task'  # Set task type as 'Child Task' for user stories and issues
        
            # Update the summary sheet with resource names and focus duration
            if resource_names != '':
                if resource_names not in summary_resource_names:
                    summary_resource_names.append(resource_names)
                    ws2[2].range('A' + str(summary_row)).value = resource_names
                    ws2[2].range('B' + str(summary_row)).value = focus_duration
                    summary_row += 1
                else:
                    # Update the focus duration for existing resource names
                    summary_index = summary_resource_names.index(resource_names) + 2
                    current_value = ws2[2].range('B' + str(summary_index)).value
                    if isinstance(focus_duration, str): 
                        print(f'Focus Duration: {focus_duration}')
                    else:
                        updated_value = current_value + focus_duration
                        ws2[2].range('B' + str(summary_index)).value = updated_value

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
            'art_scrum_team': art_scrum_team,
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

    print('Data appended.')

    # Create a dictionary mapping keys to unique IDs
    key_to_unique_id = {item['key']: item['unique_id'] for item in data}
    log_row = 1  # Initialize log row counter

    # Iterate through each item in the data list to update unique ID successors and predecessors
    for item in data:
        unique_id_succ = ''  # Initialize unique ID successors as an empty string
        for id in item['succ_list']:
            if id in key_to_unique_id:  # Check if the successor ID exists in the key_to_unique_id dictionary
                id = key_to_unique_id[id]
                if unique_id_succ == '':
                    unique_id_succ = f'\'{id}'  # Add the first successor ID
                else:
                    unique_id_succ = f'{unique_id_succ},{id}'  # Append subsequent successor IDs
            else:
                unique_id_succ = ''
                log_row += 1  # Increment log row counter
                ws2[1].range('D' + str(log_row)).value = selector
                ws2[1].range('E' + str(log_row)).value = id  # Log missing successor IDs
        item['unique_id_succ'] = unique_id_succ  # Update the item with the formatted unique ID successors

        unique_id_pred = ''  # Initialize unique ID predecessors as an empty string
        for id in item['pred_list']:
            if id in key_to_unique_id:  # Check if the predecessor ID exists in the key_to_unique_id dictionary
                id = key_to_unique_id[id]
                if unique_id_pred == '':
                    unique_id_pred = f'\'{id}'  # Add the first predecessor ID
                else:
                    unique_id_pred = f'{unique_id_pred},{id}'  # Append subsequent predecessor IDs
            else:
                unique_id_pred = ''
                log_row += 1  # Increment log row counter
                ws2[1].range('D' + str(log_row)).value = selector
                ws2[1].range('F' + str(log_row)).value = id  # Log missing predecessor IDs
        item['unique_id_pred'] = unique_id_pred  # Update the item with the formatted unique ID predecessors
    
    print('v/ Checkpoint #4')
    
    # LOG SHEET PART 1
    from datetime import datetime

    # Get the current date and time with correct formatting
    now = datetime.now()
    current_date = now.strftime("%m/%d/%Y")
    current_time = now.strftime("%I:%M:%S %p")

    # Populate the log sheet with timestamps
    ws2[1].range('A2').value = f'{current_date} {current_time}'
    ws2[1].range('A3').value = f'{current_date} {current_time}'
    ws2[1].range('A4').value = f'{current_date} {current_time}'
    ws2[1].range('A5').value = f'{current_date} {current_time}'
    ws2[1].range('A6').value = f'{current_date} {current_time}'
    processing_time = time.time()  # Record the processing time

    print('Log sheet initialized.')

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
        ws2[0].range('K' + str(proj_row)).value = item['art_scrum_team']
        ws2[0].range('L' + str(proj_row)).value = item['resource_names']
        ws2[0].range('M' + str(proj_row)).value = item['task_type']
        ws2[0].range('N' + str(proj_row)).value = 'User Stories'
        ws2[0].range('O' + str(proj_row)).value = item['sum_total_sp']
        ws2[0].range('P' + str(proj_row)).value = item['sum_done_sp']
        ws2[0].range('Q' + str(proj_row)).value = item['focus_duration']
        if item['focus_duration'] != '1 hr':
            summary_total += item['focus_duration']
        ws2[0].range('R' + str(proj_row)).value = item['low_risk_duration']
        ws2[0].range('S' + str(proj_row)).value = item['remaining_sp']
        ws2[0].range('T' + str(proj_row)).value = item['perc_complete']
        proj_row += 1

    print('Data printed. Project sheet finalized.')

    # SUMMARY SHEET
    ws2[2].range('A' + str(summary_row)).value = 'Total FDR:'
    ws2[2].range('B' + str(summary_row)).value = summary_total

    print('Summary sheet finalized.')

    # LOG SHEET PART 2
    # Get the current date and time
    now = datetime.now()
    # Format date and time correctly
    current_date = now.strftime("%m/%d/%Y")
    current_time = now.strftime("%I:%M:%S %p")
    complete_time = (time.time() - processing_time)
    ws2[1].range('A7').value = f'{current_date} {current_time}'
    ws2[1].range('A8').value = f'{current_date} {current_time}'
    ws2[1].range('B2').value = 'Start:'
    ws2[1].range('B3').value = 'Input File:'
    ws2[1].range('B4').value = 'Output File:'
    ws2[1].range('B5').value = 'Info'
    ws2[1].range('B6').value = 'Processing:'
    ws2[1].range('B7').value = 'Result:'
    ws2[1].range('B8').value = 'Processing Complete:'
    ws2[1].range('C2').value = 'DANSAP Formatter, v1.6'
    ws2[1].range('C3').value = data_file
    ws2[1].range('C4').value = new_file_name + '.xlsx'
    ws2[1].range('C5').value = mapping_file
    ws2[1].range('C6').value = f'Jira Data: Rows: {data_last_row}, Cols: 16'
    ws2[1].range('C7').value = f'Project Data: Rows: {proj_row - 2}, Cols: 20'
    ws2[1].range('C8').value = f'{complete_time} seconds'

    print('Log sheet finalized.')

        # Set column headers for the 'Project' sheet
    ws2[0].range('A1').value = [
        'Text10', 'Unique ID', 'UniqueIDSuccessors', 'UniqueIDPredecessors', 'Text11', 'Text12', 
        'Text13', 'Name', 'Text14', 'Text15', 'Text16', 'Resource Names', 'Text17', 'Text18', 
        'Text19', 'Text20', 'Duration', 'Duration1', 'Text21', '% Complete'
    ]

    # Auto-fit columns and rows for better readability
    ws2[0].autofit('c')
    ws2[0].autofit('r')
    ws2[1].autofit('c')
    ws2[1].autofit('r')
    ws2[2].autofit('c')
    ws2[2].autofit('r')

    # Save the workbook with the new file name
    wb2.save(new_file_name + '.xlsx')

    # Make the workbook visible to the user
    wb2.app.visible = True

    print('Done. File saved.')

    # Calculate and print the script execution time
    end_time = time.time()  # Record the end time
    duration_sec = end_time - start_time  # Calculate the duration in seconds
    duration_min = duration_sec / 60  # Convert the duration to minutes

    print(f"Script execution time: {duration_sec:.2f} seconds or {duration_min:.2f} minutes")

if __name__ == "__main__":
    """
    Main entry point for the script.
    Prompts the user for input and calls the process_excel function to process the data.
    """

    # Prompt the user to enter the name of the data file (without the .xls extension)
    data_file = input('Enter the name of the data file you would like to input into this program (without .xls): ') + '.xls' 
    # Example: Sentinel_Hugo_Produc_240802_1103 (Adam + with RefinementLevel)
    # Example: PnP_Categorized_240904_1601 (discipline+refinement level updated)
    # Example: IRR_Medium_Term_NEW_240913_1413
    # "C:\Users\sapird3\Downloads\IRR_Medium_Term_NEW_241002_1155.xls"
    
    # Prompt the user to enter the desired name for the new Excel file (without the .xlsx extension)
    new_file_name = input('Enter the name you would like for the new excel file (without .xlsx): ') 
    # Example: 2024-06-18_19-58-51_Project_Import_DS_
    # Example: Project_Import_DS_

    # Prompt the user to enter the percentage estimation for the worst-case scenario
    perc_est = input('Enter the percentage of worst-case estimation from realistic (without %): ')
    perc_est = int(perc_est) / 100 + 1  # Convert to decimal and add one to multiply duration time
    # Example: 20

    # Prompt the user to enter the name of the mapping team converter file (optional)
    ask = input('Do you have a mapping team converter file to input? (Y/N) ')
    if ask.lower() in ['y', 'yes']:
        mapping_file = input('Enter the name of the mapping team converter file you would like to input into this program (without .xlsx): ') + '.xlsx' 
        # Example: Team Mapping for Jira Project converter - Copy2
        # Example: Team Mapping for Jira Project converter - Copy
        # Example: Team Mapping for Jira Project converter (1)
    else:
        mapping_file = ''  # No mapping file provided

    # Call the process_excel function with the provided inputs
    process_excel(data_file, new_file_name, perc_est, mapping_file)
