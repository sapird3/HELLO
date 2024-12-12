from UpdateSentinelSchedule import SentinelScheduleUpdater
#from FormatterToolPython import FormatterToolPython
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

class RunThatUpdate:
    def __init__(self):
        print("Get ready to rock a schedule update!!!!")
    
    def main(self):
        
        #Not presently operational.  Maybe sometime soon, but I don't have time to continue working to port this.
        #FormatterToolPython = FormatterToolPython("C:\Python Code\ScheduleToolPython\LatestStructExport.xlsx", "C:\Python Code\ScheduleToolPython\ProjectImportPythonGenerated.xlsx")
        #FormatterToolPython.formatterMain()
        
        # Sentinel ScheduleUpdater takes the following arguments
        #1: Latest export of the JIRA structure in xlsx format
        #2: Path to the formatter tool excel macro
        #3: Path to output the new schedule
        #4: Path to baseline schedule
        #5: Path to export schedule update summary
        #6: debug flag to export JSON and be more verbose for debugging purposes
        
        # Create the root window
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        
        # Baseline Schedule
        oldSchedulePath = filedialog.askopenfilename(title="Select the MPP schedule you would like to update", filetypes=(("MS Project Files", "*.mpp"),  ("all files", "*.*")))
        oldSchedulePath = oldSchedulePath.replace('/', '\\')
        
        # Latest Structure Export File
        latestJiraStructExport = filedialog.askopenfilename(title="Select the xlsx of the latest JIRA structure export (MUST BE XLSX)", filetypes=(("Excel files", "*.xlsx"),  ("all files", "*.*")))
        latestJiraStructExport = latestJiraStructExport.replace('/', '\\')
        
        # Location to save Schedule Update
        now = datetime.now()
        current_date = now.strftime("%m.%d.%Y")
        default_filename = f"UpdatedSchedule_{current_date}.mpp"
        newSchedulePath = filedialog.asksaveasfilename(title="Select the location to save the updated MPP schedule", initialfile=default_filename, filetypes=(("Project Files", "*.mpp"),  ("all files", "*.*")))
        newSchedulePath = newSchedulePath.replace('/', '\\')
                      
        # Location to save the Schedule Summary Report
        default_filename = f"ScheduleUpdateSummary_{current_date}.xlsx"
        scheduleUpdateSummaryPath= filedialog.asksaveasfilename(title="Select the location to save the XLSX summary report for the schedule update", initialfile=default_filename, filetypes=(("Excel files", "*.xlsx"),  ("all files", "*.*")))
        scheduleUpdateSummaryPath = scheduleUpdateSummaryPath.replace('/', '\\')
        
        sentinelScheduleUpdater = SentinelScheduleUpdater(latestJiraStructExport,"Formatter.xlsm", newSchedulePath, oldSchedulePath, scheduleUpdateSummaryPath, True)
        sentinelScheduleUpdater.RunTheRoutine()

if __name__ == '__main__':    
   RunThatUpdate = RunThatUpdate()
   RunThatUpdate.main()
  