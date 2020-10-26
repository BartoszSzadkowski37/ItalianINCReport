import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
import Incident
import pyinputplus as pyip

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
    global incidents
    # loop through rows
    for row in range(2, sheet.max_row + 1):
        currentRowResolvedDate = sheet[resolvedCol+str(row)].value
        if currentRowResolvedDate > lastReportDate:
            currentSnowNumber = sheet[snowNumberCol+str(row)].value
            currentShortDesc = sheet[shortDescCol+str(row)].value
            currentDesc = sheet[descCol+str(row)].value
            currentResultInfo = sheet[resolvedCol+str(row)].value

            currentIncident = Incident.Incident(currentSnowNumber, currentShortDesc, currentDesc, currentRowResolvedDate, currentResultInfo)

            incidents.append(currentIncident)

            f = open('incidents.txt', 'a')
            f.write(str(currentIncident.snowNumber) + '\n')
            f.close()
        

#opening the workbook
wb=openpyxl.load_workbook('incident.xlsx')
sheet=wb['Page 1']

#creating variables for columns' letters
snowNumberCol=''
shortDescCol=''
descCol=''
resolvedCol=''
resolutionInfoCol=''

#getting columns' letters 
getLetters()

#array of incident objects
incidents = []


lastReportDate = pyip.inputDatetime('Date of last report: ', formats=['%Y/%m/%d %H:%M'])

#filter incidents by date provided by user and appending incidents array by filtered objects
dataFilter(lastReportDate)