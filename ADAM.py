# Project for Adam Calvert
# BY DANA SAPIR
# 07/08/2024

# Input an excel file and output an excel file in a different specific format

# input excel: IRR_Medium_Term_NEW_240617_1621 - full list2.xls
# old output: 2024-06-18_19-58-51_Project_Import.xlsx
# new output: IRR_Medium_Term_NEW_240710_1645.xls

# import program to grab data from excel
import xlwings as xw

# input excel from user
file = input('Enter the name of the file you would like to input into this program: ') + '.xlsx' # IRR_Medium_Term_NEW_240617_1621 - full list2
sheet = input('Enter the name of the specific sheet: ') # Sheet1

# specify excel file & sheet
wb1 = xw.Book(file)
ws1 = wb1.sheets[sheet]
wb2 = xw.Book()
ws2 = wb2.sheets[0]

# new naming
new_file_name = input('Enter the name you would like for the new excel file (without .xlsx): ') # IRR_Medium_Term_NEW_240710_1645_DS
new_sheet_name = input('Enter the name you would like for the new excel sheet: ') # IRR Medium Term - NEW! DS
ws2.name = new_sheet_name

# find last row
last_row = ws1.range('A' + str(ws1.cells.last_cell.row)).end('up').row

# collect necessary info
for value in range(1, last_row + 1):
    info1_index = ws1.range('A' + str(value)).value
    info2_key = ws1.range('B' + str(value)).value
    info3_summary = ws1.range('E' + str(value)).value
    info4_ref_lvl = ws1.range('M' + str(value)).value
    info5_art_scrum_team = ws1.range('K' + str(value)).value
    info6_status = ws1.range('G' + str(value)).value
    info7_real_est = ws1.range('M' + str(value)).value
    info8_worst_case_est = ws1.range('M' + str(value)).value
    info9_story_pts = ws1.range('M' + str(value)).value
    info10_par_feature = ws1.range('L' + str(value)).value

    # organize into new excel
    ws2.range('A' + str(value)).value = info1_index
    ws2.range('B' + str(value)).value = info2_key
    ws2.range('C' + str(value)).value = info3_summary
    ws2.range('D' + str(value)).value = info4_ref_lvl
    ws2.range('E' + str(value)).value = info5_art_scrum_team
    ws2.range('F' + str(value)).value = info6_status
    ws2.range('G' + str(value)).value = info7_real_est
    ws2.range('H' + str(value)).value = info8_worst_case_est
    ws2.range('I' + str(value)).value = info9_story_pts
    ws2.range('J' + str(value)).value = info10_par_feature

# setting up titles in new excel
ws2.range('A1').value = ['Key', 'Isuue Type', 'Summary', 'Refinement Level', 'Art Scrum Team', 'Status', 
                         'Σ Realistic Estimate', 'Σ Worst-Case Estimate', 'Σ Story Points', 'Parent Feature']
ws2.range('A1:J1').api.Font.Bold = True  # Apply bold font to the header row
ws2.autofit('c')

# save to file
wb2.save(new_file_name + '.xlsx')

# open new excel
wb2.app.visible = True

# output new excel to pdf
# ws2.api.PrintOut()