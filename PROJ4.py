# PROJECT INTRO WORK PT 4
# DS 6/11/2024

# read from a log file you would find on a system: check the Logs_To_Parse folder with 3 example scenarios
# good place to start brainstorming/experimenting with how you can read and create hashes similar to how you did with the Excel

# FILE 1: timeout_error.log

# data from sharepoint
import requests

# Make a GET request to the SharePoint file URL
file = 'https://medtronic.sharepoint.com/sites/SCSW/Shared%20Documents/Forms/AllItems.aspx?csf=1&web=1&e=zTts2t&cid=a517ab26%2Dcd20%2D42b9%2D806e%2D7b9f34a97659&FolderCTID=0x0120008944CD929AE71E449D37E9C02CFCFC94&OR=Teams%2DHL&CT=1718194879003&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNDA1MDMwNzYxNyIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D&id=%2Fsites%2FSCSW%2FShared%20Documents%2FSummer%20Intern%20Project%2007%2D24%2FLogs%5FTo%5FParse%2Ftimeout%5Ferror%2Elog&viewid=11229f96%2D3fe9%2D41db%2Dba6b%2D99caf8644c3c&parent=%2Fsites%2FSCSW%2FShared%20Documents%2FSummer%20Intern%20Project%2007%2D24%2FLogs%5FTo%5FParse'
response = requests.get(file)

# Check if the request was successful
if response.status_code >= 20:
    # Split the content into lines and determine the last row
    lines = response.text.splitlines()
    last_row = lines[-1]  # Get the last row
    print("Last row:", last_row)
else:
    print(f"Failed to retrieve the file. Status code: {response.status_code}")

