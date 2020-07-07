"""Updates the Task Number and Milestone fields in the BasinStudies feature class
   for the FLoodMigitatinoStudies.gdb.
   This is for the NCFMP project."""

# basin_update.py
# Author: Jesse Morgan
# Contact: jesse.morgan@atkinsglobal.com
# Date: 7/6/2020
# Version: 2
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
arcpy.env.workspace = workspace  # Set the workspace
CAPE_FEAR_COUNT = 0 # Count of values found for Cape Fear
CASHIE_COUNT = 0 # Count of values found for Cashie Basin
NE_CAPE_FEAR_COUNT = 0 # Count of values found for North East Cape Fear

# Constant variables
FEATURE_CLASS = "BasinStudies" # Feature class to update
basin_cell_list = ['V48', 'W48', 'X48', 'Y48', 'AN48', 'AO48', 'AP48', 'AQ48', 'AR48', 'AS48',
                   'AT48', 'AU48', 'AV48', 'AW48', 'Z49', 'AA49', 'AB49', 'AC49', 'AD49',
                   'AE49', 'AN49', 'AO49', 'AP49', 'AQ49', 'AR49', 'AS49', 'AT49', 'AU49',
                   'AV49', 'AW49', 'AF50', 'AG50', 'AH50', 'AI50', 'AJ50', 'AK50', 'AL50',
                   'AM50', 'AN50', 'AO50', 'AP50', 'AQ50', 'AR50', 'AS50', 'AT50', 'AU50',
                   'AV50', 'AW50']

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
if 'ESP_2D_Actual' not in sheet_names:
    arcpy.AddError("Could not find the ESP_2D_Actual worksheet. Exiting...")
    sys.exit(1)

sheet = workbook.get_sheet_by_name('ESP_2D_Actual')

# Iterate through the rows in the worksheet
for basin_cell in basin_cell_list:
    # Get the cell value for basin_cell cell
    cell_value = sheet[basin_cell].value

    if cell_value:
        if basin_cell.endswith('48'):
            CAPE_FEAR_COUNT += 1
        elif basin_cell.endswith('49'):
            CASHIE_COUNT += 1
        elif basin_cell.endswith('50'):
            NE_CAPE_FEAR_COUNT += 1

# Check for field already existing
field_list = arcpy.ListFields(FEATURE_CLASS)

# Add the new field
arcpy.AddField_management(FEATURE_CLASS, new_field_name, "SHORT")

# Iterate through feature class and update the columns
fields = ["Name", "Milestone", "Task_Num", new_field_name]
with UpdateCursor(FEATURE_CLASS, fields) as cursor:
    for row in cursor:
        MILESTONE = '00'
        TASK_NUM = '00'

        # Cape Fear Basin Update
        if row[0] == 'Cape Fear Basin':
            if 1 <= CAPE_FEAR_COUNT <= 4:
                MILESTONE = '0' + str(CAPE_FEAR_COUNT)
                TASK_NUM = '05'
            elif CAPE_FEAR_COUNT in (5, 6):
                MILESTONE = str(CAPE_FEAR_COUNT + 14)
                TASK_NUM = '06'
            elif CAPE_FEAR_COUNT in (7, 8):
                MILESTONE = str(CAPE_FEAR_COUNT + 14)
                TASK_NUM = '07'
            elif 9 <= CAPE_FEAR_COUNT <= 14:
                MILESTONE = str(CAPE_FEAR_COUNT + 14)
                TASK_NUM = '08'

        # Cashie Basin Update
        elif row[0] == 'Cashie Basin':
            if 1 <= CASHIE_COUNT <= 6:
                if CASHIE_COUNT != 6:
                    MILESTONE = '0' + str(CASHIE_COUNT + 4)
                else:
                    MILESTONE = str(CASHIE_COUNT + 4)
                TASK_NUM = '05'
            elif CASHIE_COUNT in (7, 8):
                MILESTONE = str(CASHIE_COUNT + 12)
                TASK_NUM = '06'
            elif CASHIE_COUNT in (9, 10):
                MILESTONE = str(CASHIE_COUNT + 12)
                TASK_NUM = '07'
            elif 11 <= CASHIE_COUNT <= 16:
                MILESTONE = str(CASHIE_COUNT + 12)
                TASK_NUM = '08'

        # Northeast Cape Fear Basin Update
        elif row[0] == 'Northeast Cape Fear Basin':
            if 1 <= NE_CAPE_FEAR_COUNT <= 8:
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '05'
            elif NE_CAPE_FEAR_COUNT in (9, 10):
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '06'
            elif NE_CAPE_FEAR_COUNT in (11, 12):
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '07'
            elif 13 <= NE_CAPE_FEAR_COUNT <= 18:
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '08'

        # Update the Milestone field
        row[1] = MILESTONE

        # Update the Task_Num field
        row[2] = TASK_NUM

        # Update the New Date field
        row[3] = TASK_NUM

        # Display Values
        data_found = "Name: " + row[0] + "\tTask Num: " + str(row[2]) +\
                     "\tStatus Code: " + str(row[3])
        arcpy.AddMessage(data_found)

        # Update the row
        cursor.updateRow(row)

=======
"""Updates the Task Number and Milestone fields in the BasinStudies feature class
   for the FLoodMigitatinoStudies.gdb.
   This is for the NCFMP project."""

# basin_update.py
# Author: Jesse Morgan
# Contact: jesse.morgan@atkinsglobal.com
# Date: 7/6/2020
# Version: 2
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
arcpy.env.workspace = workspace  # Set the workspace
CAPE_FEAR_COUNT = 0 # Count of values found for Cape Fear
CASHIE_COUNT = 0 # Count of values found for Cashie Basin
NE_CAPE_FEAR_COUNT = 0 # Count of values found for North East Cape Fear

# Constant variables
FEATURE_CLASS = "BasinStudies" # Feature class to update
basin_cell_list = ['V48', 'W48', 'X48', 'Y48', 'AN48', 'AO48', 'AP48', 'AQ48', 'AR48', 'AS48',
                   'AT48', 'AU48', 'AV48', 'AW48', 'Z49', 'AA49', 'AB49', 'AC49', 'AD49',
                   'AE49', 'AN49', 'AO49', 'AP49', 'AQ49', 'AR49', 'AS49', 'AT49', 'AU49',
                   'AV49', 'AW49', 'AF50', 'AG50', 'AH50', 'AI50', 'AJ50', 'AK50', 'AL50',
                   'AM50', 'AN50', 'AO50', 'AP50', 'AQ50', 'AR50', 'AS50', 'AT50', 'AU50',
                   'AV50', 'AW50']

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
if 'ESP_2D_Actual' not in sheet_names:
    arcpy.AddError("Could not find the ESP_2D_Actual worksheet. Exiting...")
    sys.exit(1)

sheet = workbook.get_sheet_by_name('ESP_2D_Actual')

# Iterate through the rows in the worksheet
for basin_cell in basin_cell_list:
    # Get the cell value for basin_cell cell
    cell_value = sheet[basin_cell].value

    if cell_value:
        if basin_cell.endswith('48'):
            CAPE_FEAR_COUNT += 1
        elif basin_cell.endswith('49'):
            CASHIE_COUNT += 1
        elif basin_cell.endswith('50'):
            NE_CAPE_FEAR_COUNT += 1

# Check for field already existing
field_list = arcpy.ListFields(FEATURE_CLASS)

# Add the new field
arcpy.AddField_management(FEATURE_CLASS, new_field_name, "SHORT")

# Iterate through feature class and update the columns
fields = ["Name", "Milestone", "Task_Num", new_field_name]
with UpdateCursor(FEATURE_CLASS, fields) as cursor:
    for row in cursor:
        MILESTONE = '00'
        TASK_NUM = '00'

        # Cape Fear Basin Update
        if row[0] == 'Cape Fear Basin':
            if 1 <= CAPE_FEAR_COUNT <= 4:
                MILESTONE = '0' + str(CAPE_FEAR_COUNT)
                TASK_NUM = '05'
            elif CAPE_FEAR_COUNT in (5, 6):
                MILESTONE = str(CAPE_FEAR_COUNT + 14)
                TASK_NUM = '06'
            elif CAPE_FEAR_COUNT in (7, 8):
                MILESTONE = str(CAPE_FEAR_COUNT + 14)
                TASK_NUM = '07'
            elif 9 <= CAPE_FEAR_COUNT <= 14:
                MILESTONE = str(CAPE_FEAR_COUNT + 14)
                TASK_NUM = '08'

        # Cashie Basin Update
        elif row[0] == 'Cashie Basin':
            if 1 <= CASHIE_COUNT <= 6:
                if CASHIE_COUNT != 6:
                    MILESTONE = '0' + str(CASHIE_COUNT + 4)
                else:
                    MILESTONE = str(CASHIE_COUNT + 4)
                TASK_NUM = '05'
            elif CASHIE_COUNT in (7, 8):
                MILESTONE = str(CASHIE_COUNT + 12)
                TASK_NUM = '06'
            elif CASHIE_COUNT in (9, 10):
                MILESTONE = str(CASHIE_COUNT + 12)
                TASK_NUM = '07'
            elif 11 <= CASHIE_COUNT <= 16:
                MILESTONE = str(CASHIE_COUNT + 12)
                TASK_NUM = '08'

        # Northeast Cape Fear Basin Update
        elif row[0] == 'Northeast Cape Fear Basin':
            if 1 <= NE_CAPE_FEAR_COUNT <= 8:
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '05'
            elif NE_CAPE_FEAR_COUNT in (9, 10):
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '06'
            elif NE_CAPE_FEAR_COUNT in (11, 12):
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '07'
            elif 13 <= NE_CAPE_FEAR_COUNT <= 18:
                MILESTONE = str(NE_CAPE_FEAR_COUNT + 10)
                TASK_NUM = '08'

        # Update the Milestone field
        row[1] = MILESTONE

        # Update the Task_Num field
        row[2] = TASK_NUM

        # Update the New Date field
        row[3] = TASK_NUM

        # Display Values
        data_found = "Name: " + row[0] + "\tTask Num: " + str(row[2]) +\
                     "\tStatus Code: " + str(row[3])
        arcpy.AddMessage(data_found)

        # Update the row
        cursor.updateRow(row)
