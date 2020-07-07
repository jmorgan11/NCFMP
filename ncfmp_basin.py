"""Updates the STATUS_CODE field in the RAS2D feature class for the FLoodMigitatinoStudies.gdb
   This is for the NCFMP project."""

# ncfmp_basin_update.py
# Author: Jesse Morgan
# Contact: jesse.morgan@atkinsglobal.com
# Date: 7/2/2020
# Version: 3
# Update: changed default value from 99 to 00

# Import needed modules
import sys
import openpyxl
import arcpy
from arcpy.da import UpdateCursor

# Variables
workspace = sys.argv[1] # Workspace from user
excel_file = sys.argv[2] # Excel file from user
date = sys.argv[3] # Date from user
huc_status_dict = {} # Dictionary of huc codes and status codes
arcpy.env.workspace = workspace  # Set the workspace
ROW = 4 # Start on row 4 in the Excel file

# Constant variables
MAX_ROW = 83 # Maximum row number to process
FEATURE_CLASS = "RAS2D" # Feature class to update

# Check if the feature class exists
if not arcpy.Exists(workspace + '\\' + FEATURE_CLASS):
    arcpy.AddError(FEATURE_CLASS + " does not exist.  Exiting...")
    sys.exit(1)

# Convert the ESRI Date string to a format of YearMonthDay
date_only = date.split()[0]
month, day, year = date_only.split('/')

# Add a zero if month or day is a single character
if len(day) == 1:
    day = "0" + day
if len(month) == 1:
    month = "0" + month
new_field_name = "Status_" + year + month + day

# Open the Excel File
workbook = openpyxl.load_workbook(excel_file, data_only=True)

# Check for the worksheet and set it as the current sheet
sheet_names = workbook.get_sheet_names() # Worksheets in the workbook
if 'Dashboard Tracking' not in sheet_names:
    arcpy.AddError("Could not find the Dashboard Tracking worksheet. Exiting...")
    sys.exit(1)

sheet = workbook.get_sheet_by_name('Dashboard Tracking')

# Iterate through the rows in the worksheet
while ROW <= MAX_ROW:
    # Get the HUC10 value (B4) and Overall status (Z4)
    huc_code = sheet['B' + str(ROW)].value
    STATUS_CODE = sheet['Z' + str(ROW)].value

    # Putting these in a dictionary
    huc_status_dict[huc_code] = STATUS_CODE
    ROW += 1

# Check for field already existing
field_list = arcpy.ListFields(FEATURE_CLASS)

# Add the new field
arcpy.AddField_management(FEATURE_CLASS, new_field_name, "SHORT")

# Iterate through feature class and update the columns
fields = ["HUC10", "Milestone", "Task_Num", new_field_name]
with UpdateCursor(FEATURE_CLASS, fields) as cursor:
    for row in cursor:
        if row[0] in huc_status_dict.keys():
            # Update Milestones code
            STATUS_CODE = huc_status_dict[row[0]]
            STATUS_CODE_STR = str(STATUS_CODE)
            if STATUS_CODE is None:
                STATUS_CODE = 0
            elif len(str(STATUS_CODE_STR)) == 1:
                STATUS_CODE_STR = "0" + STATUS_CODE_STR
            row[1] = STATUS_CODE_STR

            # Calc with status code
            row[3] = STATUS_CODE_STR

            # Update Task field milestone
            if 1 <= STATUS_CODE <= 5:
                row[2] = "01"
            elif 6 <= STATUS_CODE <= 9:
                row[2] = "02"
            elif 10 <= STATUS_CODE <= 13:
                row[2] = "03"
            elif 14 <= STATUS_CODE <= 17:
                row[2] = "04"
            elif 18 <= STATUS_CODE <= 20:
                row[2] = "05a"
            elif STATUS_CODE == 21:
                row[2] = "06"
            else:
                row[2] = "00"

            # Display Values
            data_found = "HUC: " + row[0] + "\tTask Num: " + row[2] + "\tStatus Code: " + row[3]
            arcpy.AddMessage(data_found)

            # Update the row
            cursor.updateRow(row)
