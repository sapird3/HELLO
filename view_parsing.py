import os
import re

def gather_code_import():
    '''
    Provide user interface instructions. Prompt for the view input.

    Returns:
        A (very long) string of code that contains the version the user would like to duplicate with new ART Scrum Teams.
        style (string): Used to decide whether encoding needs to be used between a TEAMS or SPRINTS input.
    '''
    print('INSTRUCTIONS: When you are on a structure, click on the view drop down, then select "Manage Views".')
    print('Then click "Details" that appears on the right as you hover over the current view name. Then click the "Advanced" header.')
    print('Now select the entire text in the View Specification and copy. You may delete the contents in the box, as we will be supplying it with new code.')
    
    print('\nOpen a new text document file and paste in the text.')
    text_file = input('Save this file and enter its path here: ')

    style = input('Enter whether this view is from a structure of TEAMS or SPRINTS: ')
    
    if style == 'TEAMS':
        with open(text_file, 'r') as file:
            code = file.read()
    else:
        with open(text_file, 'r', encoding='utf-8', errors='ignore') as file:
            code = file.read()

    return code, style

def gather_teams_list(file):
    '''
    Gather the list of teams that the new version of code aims to contain.

    Parameters:
        teams (string): Inputted by user.
    
    Returns:
        teams_list (list): 
    '''
    header_list = []
    team_list = []
    capacity_list = []

    last_row = file.range('A' + str(file.cells.last_cell.row)).end('up').row
    for row in range(2, last_row + 1):
        header_list.append(file.range('A' + str(row)).value)
        team_list.append(file.range('B' + str(row)).value)
        capacity_string = str((file.range('C' + str(row)).value))
        capacity_list.append(capacity_string)

    return header_list, team_list, capacity_list

def gather_section_code(code, sections):
    '''
    Find the beginning section of the View Specification code to use as a common base, since it should stay the same.
    Also, extract the template for team-specific columns and identify the placeholder team name.

    Parameters:
        code (string):

    Return:
        base_code (string): 
        teams_code_template (string):
        placeholder_team (string):
    '''
    # Find the base code before the column's section
    base_section = -1
    for _ in range(sections):
        base_section = code.find('{"key":"formula","name":', base_section + 1)

    column_start = base_section
    base_code = code[0:column_start]

    # Find the column's section
    column_end = code.find('{"key":"formula","name":', column_start + 1)
    column_section = code[column_start:column_end]

    return base_code, column_section

def team_code_template(template, header_list, team_list, capacity_list, placeholder_header, placeholder_team, placeholder_capacity):
    '''
    Builds a template from the old view based off the first ART Scrum Team. 
    Then, repopulates it for each new team this project contains (as inputted by the user).

    Parameters:
        template (string):
        new_teams_list (list): 
        placeholder_team (string):
        new_scrum_teams_list (list): 
        placeholder_scrum (string):

    Returns:
        teams_code (string):
    '''
    teams_code = ''
    
    for header in header_list:
        index = header_list.index(header)
        new_team_code = template.replace(placeholder_header, header)
        new_team_code = new_team_code.replace(placeholder_team, team_list[index])
        new_team_code = new_team_code.replace(placeholder_capacity, capacity_list[index])
        teams_code += new_team_code
    
    return teams_code

def PI_renumbering(code, oldPI, newPI):
    '''
    Replaces the old PI number with the new number.

    Parameters:
        code (string): Continuously inputted code.
        oldPI (string): The PI this input version is using.
        newPI (string): The current PI being used within the company.
    
    Return: 
        code (string): Updated code.
    '''
    code = code.replace(f'PI_number = {oldPI}', f'PI_number = {newPI}')
    oldPI = f'PI{oldPI}'
    newPI = f'PI{newPI}'
    code = code.replace(oldPI, newPI)
    
    return code

def output_code(code):
    '''
    Outputs the new and finalized version of the view specification for the user to paste into Jira for an updated Structure IN A TEXT DOCUMENT FILE.

    Parameters:
        code (string)
    '''
    with open('New_View_Code.txt', 'w', encoding='utf-8') as file:
        file.write(code)

    #with open('New_View_Code.txt', 'w') as file:
        #file.write(code)

    print('Please copy and paste the text from the new text document file named "New_View_Code.txt" from VS Code''s files into the node specialization box, and then press save. Enjoy your new Strucutre View!')
