# Jira to Microsft Project
# Written by Dana Sapir with Adam Calvert
# Last Updated: 10/18/2024

import os
import time
import xlwings as xw #AC: Added to save file in this .py
from excel_processor import read_excel_file, find_last_row_of_excel, read_column_data_from_excel, process_data, create_excel_file, create_sheets_in_excel, derive_time_stamp_file_name, load_data, set_headers, map_uid, resize_columns, populate_summary, populate_log, finalize_log

# Dec 3
data_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 2\IRR_Medium_Term_241203_1011.xls'
output_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 5\Jira to MS Project Converter Outputs\2024-12-03_1011_Jira Project Converter Output DS.xlsx'
# Nov 18
#data_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 2\IRR_Medium_Term_NEW_241118_0948 (4).xls'
#output_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 5\Jira to MS Project Converter Outputs\2024-11-18_0948_Jira Project Converter Output DS.xlsx'
# Nov 7
#data_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 2\IRR_Medium_Term_NEW_241107_1359.xls'
#output_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 5\Jira to MS Project Converter Outputs\2024-11-07_1359_Jira Project Converter Output DS.xlsx'
# Nov 1
#data_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 2\IRR_Medium_Term_NEW_241101_1345 (1).xls'
#output_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\JIRA PROJ CONVERTER\Step 5\Jira to MS Project Converter Outputs\2024-11-01_1345_Jira Project Converter Output DS.xlsx'

#data_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\IRR_Medium_Term_subset for testing3.xls'
#'C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\IRR_Medium_Term_subset for testing3.xls'
#C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\IRR_Medium_Term_NEW_241002_1155.xls
mapping_file_path = r'C:\Users\sapird3\Downloads\Team Mapping for Jira Project converter.xlsx'
#'C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\Team Mapping for Jira Project converter.xlsx'

output_file_sheet_data = {'Project': ['Text10', 'Unique ID', 'UniqueIDSuccessors', 'UniqueIDPredecessors', 'Text11', 'Text12', 
                                    'Text13', 'Name', 'Text14', 'Text15', 'Text16', 'Resource Names', 'Text17', 'Text18', 
                                    'Text19', 'Text20', 'Duration', 'Duration1', 'Text21', '% Complete', 'Milestone'],
                        'Log': ['DataTimeStamp', 'Selector', 'Event', 'Data'],
                        'Summary': ['Resource Names', 'Sum FDR']}

def main():
    """
    Main entry point for the script.
    Prompts the user for input and processes the data.
    """
    start_time = time.time()

    # Open data file from input folder
    wb1, ws1 = read_excel_file(data_file_path, 0)
    data_file_last_row = find_last_row_of_excel(ws1)

    # Open mapping file from input folder
    wb2, ws2 = read_excel_file(mapping_file_path, 0)
    map_file_last_row = find_last_row_of_excel(ws2)

    wb2, milestone_map = read_excel_file(mapping_file_path, 1)
    ms_map_file_last_row = find_last_row_of_excel(milestone_map)

    # Create a file in output folder
    #output_file_name = derive_time_stamp_file_name()                    
    #output_file_path = os.path.join(output_path, output_file_name)
    #output_file_path = output_file_name                                     #AC Debugging test
    #wb3 = create_excel_file(output_file_path)
    wb3 = create_excel_file(output_path)

    # Create sheets and headers for sheets
    create_sheets_in_excel(wb3, output_file_sheet_data) 

    # Read the issue type & status data from data file
    issue_type_data = read_column_data_from_excel(ws1, 'F', 1, data_file_last_row)
    print("issue_type_data complete")
    status_data = read_column_data_from_excel(ws1, 'G', 1, data_file_last_row)
    print("status_data complete")
    
    # Process the data
    perc_est = 1.2  # Example percentage estimation
    data, key_to_unique_id = process_data(ws1, ws2, wb3, milestone_map, issue_type_data, status_data, map_file_last_row, ms_map_file_last_row, perc_est)

    load_data(data, wb3, key_to_unique_id, milestone_map, ms_map_file_last_row)
    map_uid(data, wb3)

    #wb1.close()
    #wb2.close()

    populate_summary(wb3, data)

    end_time = time.time()  # Record the end time
    duration_sec = end_time - start_time  # Calculate the duration in seconds
    duration_min = duration_sec / 60  # Convert the duration to minutes

    print(f"Script execution time: {duration_sec:.2f} seconds or {duration_min:.2f} minutes")
    finalize_log(wb3, data_file_path, mapping_file_path, duration_sec, duration_min)
    resize_columns(wb3)

    wb3.save(output_path)
    #wb3.close()

if __name__ == '__main__':
    main()
