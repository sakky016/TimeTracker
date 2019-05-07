import datetime
from datetime import timedelta
from openpyxl.styles import Font
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.styles import Color, Fill
from openpyxl.worksheet.dimensions import ColumnDimension
import os

#################################################################################################
# Configurations
#################################################################################################
ROOT_PATH = r"C:\Users\m0pxnn\Documents\TimeTracking"
HEADING_FIELDS = ["     Date    ", "   Day   ", "  InTime  ", "  OutTime  ", "   Hours   "]
NUM_FIELDS = len(HEADING_FIELDS)
DATA_START_ROW = 1
MAX_MONTH_DAYS = 31
MAX_ROWS = MAX_MONTH_DAYS + 2
MAX_COLS = NUM_FIELDS

# Index of different columns in excel sheet
DATE_INDEX    = 0
WEEKDAY_INDEX = 1
INTIME_INDEX  = 2
OUTTIME_INDEX = 3
HOURS_INDEX   = 4

#################################################################################################
# Functions
#################################################################################################

#################################################################################################
# @name         : CreateNewWorkbook
# @description  : Creates a new excel document corresponding to this month. Adds heading row and formatting for
#                 it and saves it.
#################################################################################################
def CreateNewWorkbook(fileNameWithPath):
    book = Workbook()
    sheet = book.active    
    
    row_heading = HEADING_FIELDS
    sheet.append(row_heading)
    
    # Adjust width of each column
    i = 0
    while (i < len(HEADING_FIELDS)):
        sheet.column_dimensions[get_column_letter(i+1)].width = len(HEADING_FIELDS[i])
        #sheet[get_column_letter(i+1)].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        i = i + 1
    
    # Add font to heading row
    headingRowFont = Font(color='00000000', bold=True)
    for cell in sheet["1:1"]:
        cell.font = headingRowFont
    
    # Save the workbook
    book.save(fileNameWithPath)
    

#################################################################################################
# @name         : PrepareDataForToday
# @description  : Opens the excel sheet for this month and looks for presence of today date. If
#                 available, it updates its outTime and Hours fields. If not available, it will
#                 add a new entry for current day.
#################################################################################################    
def PrepareDataForToday(fileNameWithPath, dateTimeObj):
    book = openpyxl.load_workbook(fileNameWithPath)
    sheet = book.active 
    
    date = dateTimeObj.strftime("%d-%b-%Y")
    time = dateTimeObj.strftime("%H:%M:%S")   
    weekDay = dateTimeObj.strftime("%a")
    
    inTime = time
    outTime = ""
    hours = ""
    
    recordFound = False
    maxRows = sheet.max_row
    print ("Records present: ", maxRows - 1)
    
    for row in sheet.rows:
    
        # Date value in this row of the sheet
        sheet_date = row[DATE_INDEX].value
        if (sheet_date == "Date"):
            # This is the heading row, skip it
            continue
            
        # Week day in this row of the sheet    
        sheet_weekDay = row[WEEKDAY_INDEX].value
        
        # InTime value in this row of the sheet. This field should NOT BE 'None', because
        # if this entry is present, it must have an inTime.
        sheet_inTime = row[INTIME_INDEX].value
        
        # OutTime value in this row of the sheet. This field CAN be 'None'.       
        sheet_outTime = row[OUTTIME_INDEX].value
        
        # Hours value in this row of the sheet. This CAN be 'None'
        sheet_hours = row[HOURS_INDEX].value
        
        # If this record has the same date as the current date, then we need to update
        # the outTime of this entry and re-calculate the Hours.
        if (date == sheet_date):
            recordFound = True
            inTime = sheet_inTime
            outTime = time                             
                
            #hours = outTime - inTime
            hours = datetime.datetime.strptime(outTime, "%H:%M:%S") - datetime.datetime.strptime(inTime, "%H:%M:%S")
            
            # Update in excel sheet
            print("\nUpdating entry...")          
            row[OUTTIME_INDEX].value = outTime
            row[HOURS_INDEX].value = hours            
            break

    # If entry for this date is not present, only then we need to add this entry, else
    # we need to just update the current record.
    if (not recordFound):
        print("\nAdding entry...")          
        newRow = [date, weekDay, inTime, outTime, hours]
        sheet.append(newRow)
    
    # Display this info on console
    print("Date     : ", date)
    print("Day      : ", weekDay)
    print("In Time  : ", inTime)
    print("Out Time : ", outTime)
    print("Hours    : ", hours)
    
    # Save the excel sheet
    print("\nSaving [%s]\n" % fileNameWithPath)
    book.save(fileNameWithPath)

    
#################################################################################################
# Main
#################################################################################################

print()
print("+---------------------------------------------------------------------+")
print("|                      T I M E     T R A C K E R                      |")
print("+---------------------------------------------------------------------+")

# Current date-time. This will determine the entries in the time tracker.
dateTimeObj = datetime.datetime.now()

# Find filename of Excel that should hold this record
fileName = dateTimeObj.strftime("%b-%Y")
fileName = fileName + ".xlsx"
fileNameWithPath = os.path.join(ROOT_PATH, fileName)

# Open and Read excel for this month
print("Checking [%s]..." % fileNameWithPath)
if (not os.path.isfile(fileNameWithPath)):
    print("Creating workbook for this month...")
    CreateNewWorkbook(fileNameWithPath)

row_data = PrepareDataForToday(fileNameWithPath, dateTimeObj)

print("+---------------------------------------------------------------------+")
