import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
import Incident

#function to get letters of specific columns
def getLetters():
    global snowNumberCol
    global shortDescCol
    global descCol
    global resolvedCol
    global resolutionInfoCol
    for i in range(1, sheet.max_column + 1):
        columnValue = sheet[get_column_letter(i) + '1'].value
        columnLetter = get_column_letter(i) 
        if columnValue=='Number':
           snowNumberCol = columnLetter
           print(snowNumberCol)
        elif columnValue=='Short Description':
            shortDescCol = columnLetter
            print(shortDescCol)
        elif columnValue=='Description':
            descCol = columnLetter
            print(descCol)
        elif columnValue=='Resolved':
            resolvedCol = columnLetter
            print(resolvedCol)
        elif columnValue=='Resolution Information':
            resolutionInfoCol = columnLetter
            print(resolutionInfoCol)

            
    return

# This function will filter data by date and create array of INC elements
def dataFilter(lastReportDate):
    global Incidents
    # loop through rows
    for row in range(2, sheet.max_row + 1):
        #HERE

#opening the workbook
wb=openpyxl.load_workbook('incident.xlsx')
sheet=wb['Page 1']

#creating variables for columns' letters
snowNumberCol=''
shortDescCol=''
descCol=''
resolvedCol=''
resolutionInfoCol=''

getLetters()

Incidents = []

#print(sheet[snowNumberCol+'2'])

print('snowNumberCol=', snowNumberCol)
print('shortDescCol=', shortDescCol)
print('descCol=', descCol)
print('resolvedCol=', resolvedCol)
print('resolutionInfoCol=', resolutionInfoCol)
