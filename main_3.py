import os
import xlwings as xw

from view_parsing import gather_code_import, gather_teams_list, gather_section_code, team_code_template, PI_renumbering, output_code

# User interface instructions
code, style = gather_code_import()
#C:\Users\sapird3\OneDrive - Medtronic PLC\VIEW STRUCTURE PROJ\text_2.txt
#C:\Users\sapird3\Downloads\Foundation ART mapping (1).xlsx

while True:
    if style == 'TEAMS':
        file_path = input("Enter the file path of the Excel that includes your project's header names, ART Scrum Teams, and capacity numbers: ")
        wb = xw.Book(file_path)
        teams_file = wb.sheets[1]
        #C:\Users\sapird3\OneDrive - Medtronic PLC\VIEW STRUCTURE PROJ\Team Capacity Input.xlsx

        # Creates list of teams for new project
        header_list, team_list, capacity_list = gather_teams_list(teams_file)
        placeholder_header = 'Astra Developers'
        placeholder_team = 'Astra NPS Developers'
        placeholder_capacity = '138'
        base_code, column_section = gather_section_code(code, 7)
        teams_code = team_code_template(column_section, header_list, team_list, capacity_list, placeholder_header, placeholder_team, placeholder_capacity)
        end_code = '{"key":"actions","csid":"actions","params":{}}],"columnDisplayMode":2,"rowDisplayMode":11,"pins":["1","main"]}'
        code = base_code + teams_code + end_code
        break
    elif style == 'SPRINTS':
        # Creates list of sprints for new project
        sprints_num = int(input('Enter the number of sprints in your project: '))
        base_code, column_section = gather_section_code(code, sprints_num + 6)
        end_code = '{"key":"formula","name":"Σ Story Points","csid":"54","params":{"formula":"with any_sp_undefined = SUM{\n    (issuetype = \"Issue\" OR issuetype = \"User Story\")\n        AND ((NOT DEFINED(story_points)) OR story_points = 0) \n        AND (status != \"Cancelled\" AND status != \"Done\")\n}:\n\nwith total_sp = SUM{\n    IF((story_points > 0) AND (issuetype = \"Issue\" OR issuetype = \"User Story\"), \n        story_points, 0)\n}:\n\nwith done_sp = SUM{\n    IF(DEFINED(resolution) AND story_points > 0 AND (issuetype = \"Issue\" OR issuetype = \"User Story\"), \n        story_points, 0)\n}:\n\nCONCAT(\n    done_sp, \" / \", total_sp,\n    IF(any_sp_undefined, \"*\")\n)","variables":{"issuetype":{"id":"issuetype","format":"text"},"story_points":{"id":"customfield","params":{"fieldId":10002},"format":"number"},"status":{"id":"status","format":"text"},"resolution":{"id":"resolution","format":"text"}}}},{"key":"actions","csid":"actions","params":{}}],"columnDisplayMode":2,"rowDisplayMode":11,"pins":["1","main"]}'
        code = base_code + end_code
        break
    else:
        print('Please enter either "TEAMS" or "SPRINTS".')

# Renumbering PI
oldPI = input('Please enter the old PI this view uses: ')
newPI = input('Please enter the new PI you would like this view to use: ')
code = PI_renumbering(code, oldPI, newPI)

# Finialized output ready to paste
output_code(code)
