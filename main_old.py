# Jira to Microsft Project
# Written by Dana Sapir with Adam Calvert
# Last Updated: 9/20/2024

import os
from excel_processor_old import read_excel_file, find_last_row_of_excel, read_column_data_from_excel, process_data, create_excel_file, create_sheets_in_excel, derive_time_stamp_file_name

data_file_path = r'C:\Users\sapird3\Downloads\Jira_Project\Jira_Project\new\input\IRR_Medium_Term_NEW_240913_1413.xls'
mapping_file_path = r'C:\Users\sapird3\Downloads\Jira_Project\Jira_Project\new\input\Team Mapping for Jira Project converter - Copy.xlsx'
output_path = r'C:\Users\sapird3\Downloads\Jira_Project\Jira_Project\new\output'

output_file_sheet_data = {'Project': ['Text10', 'Unique ID', 'UniqueIDSuccessors', 'UniqueIDPredecessors', 'Text11', 'Text12', 
                                      'Text13', 'Name', 'Text14', 'Text15', 'Text16', 'Resource Names', 'Text17', 'Text18', 
                                      'Text19', 'Text20', 'Duration', 'Duration1', 'Text21', '% Complete'],
                          'Log': ['DataTimeStamp', 'Event', 'Data', 'Selector', 'Unique ID Successors NOT FOUND', 'Unique ID Predecessors NOT FOUND'],
                          'Summary': ['Resource Names', 'Sum FDR']}

def main():
    """
    Main entry point for the script.
    Prompts the user for input and processes the data.
    """
    # Open data file from input folder
    wb1, ws1 = read_excel_file(data_file_path)
    data_file_last_row = find_last_row_of_excel(ws1, column='A')

    # Open mapping file from input folder
    wb2, ws2 = read_excel_file(mapping_file_path)
    map_file_last_row = find_last_row_of_excel(ws2, column='A')

    # Read the issue type & status data from data file
    issue_type_data = read_column_data_from_excel(ws1, 'F', 1, data_file_last_row)
    print(issue_type_data)
    status_data = read_column_data_from_excel(ws1, 'G', 1, data_file_last_row)
    print(status_data)

    # Process the data
    perc_est = 1.2  # Example percentage estimation
    data = process_data(ws1, ws2, ws2, issue_type_data, status_data, map_file_last_row, perc_est)

    # Create a file in output folder
    output_file_name = derive_time_stamp_file_name()
    output_file_path = os.path.join(output_path, output_file_name)
    create_excel_file(output_file_path)

    # Create sheets and headers for sheets
    create_sheets_in_excel(output_file_path, output_file_sheet_data)

    wb1.close()
    wb2.close()

if __name__ == '__main__':
    main()
