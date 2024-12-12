import os
import xlwings as xw
from datetime import datetime

def read_excel_file(file_path):
    """
    Read an Excel file and return the workbook and first worksheet.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        tuple: A tuple containing the workbook and the first worksheet.
    """
    if os.path.isfile(file_path):
        wb = xw.Book(file_path)
        ws = wb.sheets[0]
    else:
        wb, ws = None, None
        print(f"The file '{file_path}' does not exist.")
    return wb, ws

def find_last_row_of_excel(ws, column='A'):
    """
    Find the last row of data in a specified column of an Excel worksheet.

    Args:
        ws (xlwings.Sheet): The Excel worksheet object.
        column (str): The column to search for the last row (default is 'A').

    Returns:
        int: The last row number with data in the specified column.
    """
    last_row = ws.range(f'{column}1').end('down').row
    return last_row

def read_column_data_from_excel(ws1, column, start_row, data_last_row):
    """
    Read data from a specified column in an Excel worksheet.

    Args:
        ws1 (xlwings.Sheet): The Excel worksheet object.
        column (str): The column to read data from.
        start_row (int): The starting row for reading data.
        data_last_row (int): The last row of data to read.

    Returns:
        list: A list of values read from the specified column.
    """
    data_range = ws1.range(f'{column}{start_row}').expand('down').resize(data_last_row, 1)
    data = data_range.value
    return data

def read_art_scrum_teams(ws1, row, ws2, art_scrum_team, map_file_last_row, issue_type):  #AC: Moved Feature team logic to row 159. Removed Feature team from list.
    """
    Read and map ART Scrum teams from an Excel worksheet.

    Args:
        ws1 (xlwings.Sheet): The input worksheet object.
        row (int): The row number to process.
        ws2 (xlwings.Sheet): The mapping worksheet object.
        art_scrum_team (str): The ART Scrum team name.
        map_file_last_row (int): The last row of the mapping file.
        issue_type
        feature_teams
        

    Returns:
        str: The mapped resource names.
    """
    if issue_type == 'Feature Estimation' or issue_type == 'User Story' or issue_type == 'Issue':
        if ws2 != '':
            for map_row in range(2, map_file_last_row + 1):
                original_team = ws2.range('A' + str(map_row)).value
                mapped_team = ws2.range('B' + str(map_row)).value
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

        """                                                         #AC: Moved Feature team logic to row 159
        if resource_names not in feature_teams:
            feature_teams.add(resource_names)
        else:
            add_total_sp = ws1.range(f'O{row + 1}').value
            if sum_total_sp is not None and add_total_sp is not None:
                sum_total_sp += add_total_sp
            excel_value3 = ws1.range('M' + str(row)).value
            if realistic_est is not None and excel_value3 is not None:
                realistic_est += excel_value3
            excel_value4 = ws1.range('N' + str(row)).value
            if worst_case_est is not None and excel_value4 is not None:
                worst_case_est += excel_value4"""

        return resource_names

def process_data(ws1, ws2, issue_type_data, status_data, map_file_last_row, perc_est): #AC: Removed ws3 after ws2
    """
    Process data from the input Excel file and create the output data structure.

    Args:
        ws1 (xlwings.Sheet): The input data worksheet object.
        ws2 (xlwings.Sheet): The mapping worksheet object.
        ws3 (xlwings.Sheet): The output data worksheet object. AC: inconsistent naming here ws2 is sometimes the mapping file and other times the output file. Corrected here
        issue_type_data (list): A list of issue types.
        status_data (list): A list of issue statuses.
        map_file_last_row (int): The last row of the mapping file.
        perc_est (float): The percentage estimation for worst-case scenario.

    Returns:
        list: A list of processed data.
    """
    data = []
    feature_key_list = []
    summary_resource_names = []
    key_to_unique_id = {}
    unique_id_counter = 1
    feature_status = 'Open'

    for row in range(2, len(issue_type_data) + 1):
        print(row)
        issue_type = issue_type_data[row - 1]
        if issue_type not in ['Feature', 'Feature Estimation', 'Issue', 'User Story']:
            continue

        status = status_data[row - 1]
        if status == 'Cancelled':
            if issue_type == 'Feature':
                feature_status = status
            continue

        unique_id = unique_id_counter
        unique_id_succ = ws1.range(f'C{row}').value
        unique_id_pred = ws1.range(f'D{row}').value
        succ_list = [value.strip() for value in unique_id_succ.split(',')] if unique_id_succ else []
        pred_list = [value.strip() for value in unique_id_pred.split(',')] if unique_id_pred else []
        index = ws1.range(f'A{row}').value
        summary = ws1.range(f'E{row}').value
        parent_feature = ws1.range(f'L{row}').value
        key = ws1.range(f'B{row}').value
        art_scrum_team = ws1.range(f'K{row}').value
        refinement_lvl = ws1.range(f'H{row}').value
        resource_names = read_art_scrum_teams(ws1, row, ws2, art_scrum_team, map_file_last_row, issue_type)

        if key in feature_key_list:
            print(f'This key was found again: {key}')
            continue
        else:
            feature_key_list.append(key)

        if issue_type == 'Feature':
            feature_teams = set()                       #AC: We were missing the logic to skip lines starting at 161 if there is a repeat resource_name in feature_teams
            feature_key = key                           #AC: We could also use Feature Key to confirm children are linked properly.
            feature_ref_level = refinement_lvl
            feature_summary = summary
            feature_row = unique_id
            feature_status = status

        if feature_status == 'Cancelled':
            continue

        if resource_names not in feature_teams:         #AC: Added based on comments in 156
            feature_teams.add(resource_names)
            sum_total_sp = float(ws1.range(f'O{row}').value) #or 0         #AC: converted to float because 166 was throwing a bug
            sum_done_sp = float(ws1.range(f'P{row}').value) #or 0          #AC: converted to float because 166 was throwing a bug
            realistic_est = ws1.range(f'M{row}').value #or 0
            worst_case_est = ws1.range(f'N{row}').value #or 0
            if feature_ref_level == "Story Points":
                remaining_sp = sum_total_sp - sum_done_sp if sum_total_sp is not None else 0 
            else:
                remaining_sp = realistic_est if realistic_est is not None else 0

            selector, task_name, focus_duration, low_risk_duration, perc_complete, task_type = get_task_details(
                issue_type, feature_key, resource_names, feature_summary, status, feature_ref_level, sum_total_sp, sum_done_sp, realistic_est, worst_case_est, perc_est, feature_key) #AC: compared row201, this was missing refinement level I think this should be Feature Refinement Level. Updated.
        
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
            
            print ("created new UID: ", unique_id_counter)  #AC: Debugging

            key_to_unique_id[row] = unique_id_counter
            unique_id_counter += 1

        else:                                            #AC: Added based on comments in 157. proj_import file sent on Sep 18
            # Find the relevant entry in the data list
            child_index = -1                               #AC change -1 to row to speed up the processing time? Define a new "feature_row" variable?
            summary_index = 2
            for i, item in enumerate(range(feature_row, unique_id + 1)):
                item = data[i + feature_row]
                print('Variable i is: ', i)
                if item['resource_names'] == resource_names and item['parent_feature'] == parent_feature:           
                    child_index = i + feature_row
                    print("Found matching team to add to totals")   #AC Debugging
                    print("Resource name: ",item['resource_names'])
                    print("Parent_ID: ", item['parent_feature'])
                    print('Variable i is: ', i)
                    print('Child index is: ', child_index)
                    break
            
            if child_index == -1:
                print(f"Error: resource_names '{resource_names}' not found in data.")
                continue

            if feature_ref_level == "Story Points":
                # sum total sp
                old_total_sp = float(data[child_index]['sum_total_sp'])
                print ("Old total Sp for row: ", old_total_sp)
                add_total_sp = ws1.range('O' + str(row)).value
                add_total_sp = float(add_total_sp) if add_total_sp is not None else 0.0
                print ("Add total sp for row", add_total_sp)
                new_total_sp = add_total_sp + old_total_sp
                print ("new total story points for row", new_total_sp)
                data[child_index]['sum_total_sp'] = new_total_sp
                print("Loaded new total SP: ", data[child_index]['sum_total_sp'])
                print("Child index: ", child_index)

                # sum done sp
                old_done_sp = float(data[child_index]['sum_done_sp'])
                add_done_sp = ws1.range('P' + str(row)).value
                add_done_sp = float(add_done_sp) if add_done_sp is not None else 0.0
                new_done_sp = add_done_sp + old_done_sp
                data[child_index]['sum_done_sp'] = new_done_sp

                # update FD and LRD and perc_complete
                data[child_index]['focus_duration'] = new_total_sp
                data[child_index]['low_risk_duration'] = new_total_sp * perc_est
                data[child_index]['perc_complete'] = f'{(new_done_sp / new_total_sp) * 100}%' if new_total_sp > 0 else 0
                data[child_index]['remaining_sp'] = new_total_sp - new_done_sp if new_total_sp is not None else 0 

            if feature_ref_level == "Feature Estimation":
                # realistic est
                old_real_est = float(data[child_index]['realistic_est']) if data[child_index]['realistic_est'] is not None else 0.0
                add_real_est = ws1.range('M' + str(row)).value
                add_real_est = float(add_real_est) if add_real_est is not None else 0.0
                new_real_est = add_real_est + old_real_est
                data[child_index]['realistic_est'] = new_real_est

                # worst case est
                old_wc_est = float(data[child_index]['worst_case_est']) if data[child_index]['worst_case_est'] is not None else 0.0
                add_wc_est = ws1.range('N' + str(row)).value
                add_wc_est = float(add_wc_est) if add_wc_est is not None else 0.0
                new_wc_est = add_wc_est + old_wc_est
                data[child_index]['worst_case_est'] = new_wc_est

                # update FD and LRD. Perc_complete will be 0% for all Feature Estimation tickets
                data[child_index]['focus_duration'] = new_real_est
                data[child_index]['low_risk_duration'] = new_wc_est
                data[child_index]['remaining_sp'] = new_real_est 


            continue

        """                             #AC: Moved up to sit inside a If statment for non-repeating teams
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
        

        print ("created new UID: ", unique_id_counter)  #AC: Debugging

        key_to_unique_id[row] = unique_id_counter
        unique_id_counter += 1
        """

    print ("Collecting data complete")


    
    """ AC: This needs to be moved out of this def function because ws3 (output excel) is not used here
    update_unique_ids(data, ws3, key_to_unique_id)
    """
    return data
    

def get_task_details(issue_type, key, resource_names, summary, status, refinement_lvl, sum_total_sp, sum_done_sp, realistic_est, worst_case_est, perc_est, feature_key):
    """
    Get the task details for an issue.

    Args:
        issue_type (str): The type of issue.
        key (str): The key of the issue.
        resource_names (str): The resource names.
        summary (str): The summary of the issue.
        status (str): The status of the issue.
        refinement_lvl: Determines if Story Points or Feature Estimations should be used.
        sum_total_sp (float): The total story points.
        sum_done_sp (float): The done story points.
        realistic_est (float): The realistic estimate.
        worst_case_est (float): The worst-case estimate.
        perc_est (float): The percentage estimation for worst-case scenario.
        refinement_lvl

    Returns:
        tuple: The task details (selector, task_name, focus_duration, low_risk_duration, perc_complete, task_type).
    """
    if issue_type == 'Feature':
        selector = key
        task_name = f'[{selector}] | {summary}'
        focus_duration = '1 hr'
        low_risk_duration = '1 hr'
        if status == 'Done':
            perc_complete = '100%'
        elif sum_total_sp == 0:
            perc_complete = '0%'
        elif refinement_lvl == 'Feature Estimation':  #AC: Added this condition because in this case sum_total_sp will likely be 0. Not sure if row 234 can handle dividing by 0.
            perc_complete = '0%'
        else:
            perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
        task_type = 'Parent Task'
    else:
        selector = f'{key}_{resource_names}'
        task_name = f'[{selector}] | {summary}'
        if refinement_lvl == 'Feature Estimation':
            focus_duration = realistic_est
            low_risk_duration = worst_case_est
        elif refinement_lvl == 'Story Points':
            focus_duration = sum_total_sp
            low_risk_duration = sum_total_sp * perc_est
        elif refinement_lvl == '':
            print("No refinement level found on: ", feature_key)  #AC Add to log no refinement level for for feature key
        if refinement_lvl == 'Feature Estimation':  #AC: Added this condition because in this case sum_total_sp will likely be 0. Not sure if row 254 can handle dividing by 0.
            perc_complete = '0%'
        elif status == 'Done':
            perc_complete = '100%'
        elif sum_done_sp == 0:
            perc_complete = '0%'
        else:
            perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
            if int(float(perc_complete.rstrip('%'))) > 100:
                perc_complete = '100%'
        task_type = 'Child Task'

    return selector, task_name, focus_duration, low_risk_duration, perc_complete, task_type

def update_unique_ids(data, ws3, key_to_unique_id):
    """
    Update the unique ID successors and predecessors in the data list.

    Args:
        data (list): The list of processed data.
        ws3 (xlwings.Sheet): The output data worksheet object.
        key_to_unique_id (dict): The dictionary mapping keys to unique IDs.
    """
    log_row = 1
    for item in data:
        unique_id_succ = get_unique_ids(item['succ_list'], key_to_unique_id, log_row, 'E', ws3)
        item['unique_id_succ'] = unique_id_succ

        unique_id_pred = get_unique_ids(item['pred_list'], key_to_unique_id, log_row, 'F', ws3)
        item['unique_id_pred'] = unique_id_pred

def get_unique_ids(id_list, key_to_unique_id, log_row, column_letter, ws2, feature_key):
    """
    Get the unique IDs for a list of IDs.

    Args:
        id_list (list): The list of IDs.
        key_to_unique_id (dict): The dictionary mapping keys to unique IDs.
        log_row (int): The current log row.
        column_letter (str): The column letter for logging.
        ws2 (xlwings.Sheet): The worksheet object for logging.
        feature_key

    Returns:
        str: The formatted unique IDs.
    """
    unique_id_str = ''
    for id in id_list:
        if id in key_to_unique_id:
            id = key_to_unique_id[id]
            if unique_id_str == '':
                unique_id_str = f'\'{id}'
            else:
                unique_id_str = f'{unique_id_str},{id}'
        else:
            unique_id_str = ''
            log_row += 1
            ws2.range(f'D{log_row}').value = feature_key
            ws2.range(f'{column_letter}{log_row}').value = id
    return unique_id_str

def create_excel_file(file_path):          
    """
    Create a new Excel file at the specified path.

    Args:
        file_path (str): The full path where the Excel file will be created.

    Returns:
        xlwings.Book: The created workbook object.
    """
    try:
        wb = xw.Book()
        print(file_path)
        wb.save(file_path)
        print(f"Workbook created and saved at '{file_path}'.")
        return wb
    except Exception as e:
        print(f"An error occurred while creating the Excel file: {e}")
    finally:
        print("did not use wb.close()")             #AC change back to wb.close() once file location is working.

def create_sheets_in_excel(wb3, sheet_data):    #AC: Changed to call wb3 instead of file_path
    """
    Create sheets in the output Excel file.

    Args:
        file_path (str): The path to the Excel file.
        sheet_data (dict): A dictionary with sheet names as keys and headers as values.
    """
    try:
        #wb = xw.Book(file_path)
        for sheet_name, sheet_headers in sheet_data.items():
            wb3.sheets.add(sheet_name)
            set_headers(wb3.sheets[sheet_name], sheet_headers)
            print(f"Sheet '{sheet_name}' created.")
        if 'Sheet1' in [sheet.name for sheet in wb3.sheets]:
            wb3.sheets['Sheet1'].delete()
        wb3.save()
        #wb3.close()
    except Exception as e:
        print(f"An error occurred while creating sheets: {e}")

def load_data (data, wb3): #AC Added 
    """
    Load data that has been populated from input file. Use the data list to populate the output file
    
    Args:
        data (list): The list of processed data.
        ws3 (xlwings.Sheet): The output data worksheet object.    

    Return:
        summary total (int): A running total for the summary tab on the Sum FDR.
    """

    ws3 = wb3.sheets[2]
    proj_row = 2
    summary_total = 0
    for item in data:
        ws3.range('A' + str(proj_row)).value = item['selector']  
        ws3.range('B' + str(proj_row)).value = item['unique_id']
        ws3.range('C' + str(proj_row)).value = item['succ_list']
        ws3.range('D' + str(proj_row)).value = item['pred_list']
        ws3.range('E' + str(proj_row)).value = item['index']
        ws3.range('F' + str(proj_row)).value = item['parent_feature']
        ws3.range('G' + str(proj_row)).value = item['key']
        ws3.range('H' + str(proj_row)).value = item['task_name']
        ws3.range('I' + str(proj_row)).value = item['issue_type']
        ws3.range('J' + str(proj_row)).value = item['status']
        ws3.range('K' + str(proj_row)).value = item['art_scrum_team']
        ws3.range('L' + str(proj_row)).value = item['resource_names']
        ws3.range('M' + str(proj_row)).value = item['task_type']
        #ws3.range('N' + str(proj_row)).value = 'User Stories'              #AC: Not sure what data should come in here? (this is not listed in the data list format)
        ws3.range('O' + str(proj_row)).value = item['sum_total_sp']
        ws3.range('P' + str(proj_row)).value = item['sum_done_sp']
        ws3.range('Q' + str(proj_row)).value = item['focus_duration']
        if item['focus_duration'] != '1 hr':                               #AC: Not sure what this was for?
            summary_total += item['focus_duration']                        # for the summary tab, a running total of the sum fdr
        ws3.range('R' + str(proj_row)).value = item['low_risk_duration']
        ws3.range('S' + str(proj_row)).value = item['remaining_sp']
        ws3.range('T' + str(proj_row)).value = item['perc_complete']
        
        proj_row += 1

    print('Data printed. Project sheet finalized.')

    return summary_total


def map_uid (data, wb3):
    """
    Update the Successors and Predecessors to replace TAP-xxx ticket numbers with UIDs
    Map UID of parents to child as successors
    
    Args:
        data (list): The list of processed data.
        ws3 (xlwings.Sheet): The output data worksheet object.    
    """


    key_to_unique_id = {item['key']: item['unique_id'] for item in data}
    log_row = 1
    missing_dependency = []

    print ("key_to_unique_id = ", key_to_unique_id)

    ws3 = wb3.sheets[2]
    
    for item in data:
        unique_id_succ = ''
        for id in item['succ_list']:
            if id in key_to_unique_id:
                id = key_to_unique_id[id]
                if unique_id_succ == '':
                    unique_id_succ = f'\'{id}'
                else:
                    unique_id_succ = f'{unique_id_succ},{id}'
            else:
                print ("Dependency ID not found in input file: ",id)
                # PUT INTO LOG
        

        unique_id_pred = ''
        for id in item['pred_list']:
            if id in key_to_unique_id:
                id = key_to_unique_id[id]
                if unique_id_pred == '':
                    unique_id_pred = f'\'{id}'
                else:
                    unique_id_pred = f'{unique_id_pred},{id}'
            else:
                print ("Dependency ID not found in input file: ",id)

                if id not in missing_dependency:
                    missing_dependency.append(id)

        log_row += 1
        ws3.range('C' + str(log_row)).value = unique_id_succ 
        ws3.range('D' + str(log_row)).value = unique_id_pred 

    log_row = 1

    for item in data:
        if item['task_type'] == "Child Task" and item["parent_feature"] is not None:
            print ("child task item: ", item)
            print ("child task log_row: ", log_row)
            print ("parent ID", item['parent_feature'])
            id = f'\'{key_to_unique_id[item['parent_feature']]}'
            print ("mapped parent UID:", id)
            ws3.range('C' + str(log_row+1)).value = id
        log_row += 1

    print ("Missing dependencies: ", missing_dependency)
        

def set_headers(ws, headers):
    """
    Set headers in the first row of an Excel worksheet.

    Args:
        ws (xlwings.Sheet): The Excel worksheet object.
        headers (list): A list of headers to set in the first row.
    """
    ws.range('A1').expand('right').value = headers

def derive_time_stamp_file_name():
    """
    Derive a timestamped file name for the output file.

    Returns:
        str: The derived file name.
    """
    now = datetime.now()
    formatted_timestamp = now.strftime('%Y-%m-%d_%H-%M-%S')
    file_name = str(formatted_timestamp) + '_' + 'Jira Project Converter Output File AC' + '.xlsx'
    return file_name



def resize_columns(wb):
    '''
    Adjust column width for user readability.
    
    Args:
        ws (xlwings): The Excel workbook object.
    '''
    for ii in range (3):
        wb[ii].autofit('c')
        wb[ii].autofit('r')

def populate_summary(resource_names, summary_total):
    '''
    
    '''