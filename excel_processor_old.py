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

def read_art_scrum_teams(ws1, row, ws3, art_scrum_team, map_file_last_row, issue_type, feature_teams):
    """
    Read and map ART Scrum teams from an Excel worksheet.

    Args:
        ws1 (xlwings.Sheet): The Excel worksheet object.
        row (int): The row number to process.
        ws3 (xlwings.Sheet): The mapping worksheet object.
        art_scrum_team (str): The ART Scrum team name.
        map_file_last_row (int): The last row of the mapping file.
        issue_type
        feature_teams
        sum_total_sp
        realistic_est
        worst_case_est

    Returns:
        str: The mapped resource names.
    """
    if issue_type == 'Feature Estimation' or issue_type == 'User Story' or issue_type == 'Issue':
        if ws3 != '':
            for map_row in range(2, map_file_last_row + 1):
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
                worst_case_est += excel_value4

        return resource_names, feature_teams

def process_data(ws1, ws3, ws2, issue_type_data, status_data, map_file_last_row, perc_est):
    """
    Process data from the input Excel file and create the output data structure.

    Args:
        ws1 (xlwings.Sheet): The input data worksheet object.
        ws3 (xlwings.Sheet): The mapping worksheet object.
        ws2 (xlwings.Sheet): The output data worksheet object.
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

    for row in range(2, len(issue_type_data) + 1):
        issue_type = issue_type_data[row - 1]
        if issue_type not in ['Feature', 'Feature Estimation', 'Issue', 'User Story']:
            continue

        status = status_data[row - 1]
        if status == 'Cancelled':
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
        resource_names = read_art_scrum_teams(ws1, row, ws3, art_scrum_team, map_file_last_row, issue_type, None)

        if key in feature_key_list:
            print(f'This feature key was found again: {key}')
            continue
        else:
            feature_key_list.append(key)

        if issue_type == 'Feature':
            feature_teams = set()
            feature_key = key
            feature_ref_level = refinement_lvl

        sum_total_sp = ws1.range(f'O{row}').value or 0
        sum_done_sp = ws1.range(f'P{row}').value or 0
        realistic_est = ws1.range(f'M{row}').value or 0
        worst_case_est = ws1.range(f'N{row}').value or 0
        remaining_sp = sum_total_sp - sum_done_sp if sum_total_sp is not None else 0

        selector, task_name, focus_duration, low_risk_duration, perc_complete, task_type = get_task_details(
            issue_type, key, resource_names, summary, status, sum_total_sp, sum_done_sp, realistic_est, worst_case_est, perc_est)

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

        key_to_unique_id[row] = unique_id_counter
        unique_id_counter += 1

    update_unique_ids(data, ws2, key_to_unique_id)
    return data

def get_task_details(issue_type, key, resource_names, summary, status, refinement_lvl, sum_total_sp, sum_done_sp, realistic_est, worst_case_est, perc_est):
    """
    Get the task details for an issue.

    Args:
        issue_type (str): The type of issue.
        key (str): The key of the issue.
        resource_names (str): The resource names.
        summary (str): The summary of the issue.
        status (str): The status of the issue.
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
            print("No refinement level found.")
        if status == 'Done':
            perc_complete = '100%'
        elif sum_done_sp == 0:
            perc_complete = '0%'
        else:
            perc_complete = f'{(sum_done_sp / sum_total_sp) * 100}%'
        if int(float(perc_complete.rstrip('%'))) > 100:
            perc_complete = '100%'
        task_type = 'Child Task'

    return selector, task_name, focus_duration, low_risk_duration, perc_complete, task_type

def update_unique_ids(data, ws2, key_to_unique_id):
    """
    Update the unique ID successors and predecessors in the data list.

    Args:
        data (list): The list of processed data.
        ws2 (xlwings.Sheet): The output data worksheet object.
        key_to_unique_id (dict): The dictionary mapping keys to unique IDs.
    """
    log_row = 1
    for item in data:
        unique_id_succ = get_unique_ids(item['succ_list'], key_to_unique_id, log_row, 'E', ws2)
        item['unique_id_succ'] = unique_id_succ

        unique_id_pred = get_unique_ids(item['pred_list'], key_to_unique_id, log_row, 'F', ws2)
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
        wb.save(file_path)
        print(f"Workbook created and saved at '{file_path}'.")
        return wb
    except Exception as e:
        print(f"An error occurred while creating the Excel file: {e}")
    finally:
        wb.close()

def create_sheets_in_excel(file_path, sheet_data):
    """
    Create sheets in the output Excel file.

    Args:
        file_path (str): The path to the Excel file.
        sheet_data (dict): A dictionary with sheet names as keys and headers as values.
    """
    try:
        wb = xw.Book(file_path)
        for sheet_name, sheet_headers in sheet_data.items():
            wb.sheets.add(sheet_name)
            set_headers(wb.sheets[sheet_name], sheet_headers)
            print(f"Sheet '{sheet_name}' created.")
        if 'Sheet1' in [sheet.name for sheet in wb.sheets]:
            wb.sheets['Sheet1'].delete()
        wb.save()
        wb.close()
    except Exception as e:
        print(f"An error occurred while creating sheets: {e}")

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
    file_name = str(formatted_timestamp) + '_' + 'Project_Import_DS_' + '.xlsx'
    return file_name
