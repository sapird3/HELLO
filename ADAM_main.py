# Jira to Microsft Project
# Written by Dana Sapir with Adam Calvert
# Last Updated: 10/18/2024

import os
import time
import xlwings as xw #AC: Added to save file in this .py
from ADAM_excel_processor import read_excel_file, find_last_row_of_excel, read_column_data_from_excel, process_data, create_excel_file, create_sheets_in_excel, derive_time_stamp_file_name, load_data, set_headers, map_uid, resize_columns, populate_summary

data_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\IRR_Medium_Term_subset for testing3.xls'
#data_file_path = r'C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\IRR_Medium_Term_subset for testing2.xls'
#data_file_path = r'C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\IRR_Medium_Term_NEW_241101_1345.xls'
mapping_file_path = r'C:\Users\sapird3\OneDrive - Medtronic PLC\Team Mapping for Jira Project converter.xlsx'
#mapping_file_path = r'C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\Team Mapping for Jira Project converter.xlsx'
output_path = r'CC:\Users\sapird3\OneDrive - Medtronic PLC\test_output1'
#output_path = r'C:\Users\calvea2\OneDrive - Medtronic PLC\My OneDrive Documents\19 Co-op\Dana Sapir\Jira MS Project Converter\Oct 2 2024 work\test_output1'

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
    start_time = time.time()

    # Open data file from input folder
    wb1, ws1 = read_excel_file(data_file_path)
    data_file_last_row = find_last_row_of_excel(ws1)

    # Open mapping file from input folder
    wb2, ws2 = read_excel_file(mapping_file_path)
    map_file_last_row = find_last_row_of_excel(ws2)

    # Read the issue type & status data from data file
    issue_type_data = read_column_data_from_excel(ws1, 'F', 1, data_file_last_row)
    print("issue_type_data complete")
    status_data = read_column_data_from_excel(ws1, 'G', 1, data_file_last_row)
    print("status_data complete")
    
    # Process the data
    perc_est = 1.2  # Example percentage estimation
    data = process_data(ws1, ws2, issue_type_data, status_data, map_file_last_row, perc_est) #AC: changed second variable from ws2 to ws3. However ws3 isn't defined yet. Remove ws3 after ws2


    # Create a file in output folder
    output_file_name = derive_time_stamp_file_name()                    
    output_file_path = os.path.join(output_path, output_file_name)
    output_file_path = output_file_name                                     #AC Debugging test
    #------------     create_excel_file(output_file_path)            #AC: I couldn't get it to save at the file path, so I save it to the default location only. Added output_file_name to this function.

    #-----------------------AC: Trying to create excel file with out create_excel_file function in excel
    wb3 = xw.Book()
    print(output_file_path)
    wb3.save(output_file_path)


    # Create sheets and headers for sheets
    create_sheets_in_excel(wb3, output_file_sheet_data) 


    summary_total = load_data (data, wb3)

    map_uid (data, wb3)
        
    #wb1.close()
    #wb2.close()

    #resize_columns(wb3)

   # populate_summary(resource_names, summary_total)
   
    end_time = time.time()  # Record the end time
    duration_sec = end_time - start_time  # Calculate the duration in seconds
    duration_min = duration_sec / 60  # Convert the duration to minutes

    print(f"Script execution time: {duration_sec:.2f} seconds or {duration_min:.2f} minutes")
    
# if __name__ == '__main__':
main()
