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
    global locationCol
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
        elif columnValue=='Location':
            locationCol = columnLetter
            print(locationCol)

            
    return

# This function will filter data by date and create array of INC elements
def dataFilter(lastReportDate):
    global incidents
    # loop through rows
    for row in range(2, sheet.max_row + 1):
        currentRowResolvedDate = sheet[resolvedCol+str(row)].value
        currentLocation = sheet[locationCol+str(row)].value
        if currentRowResolvedDate > lastReportDate and  currentLocation == 'HRAIL.IT.Napoli.Hitachi Rail Spa':
            currentSnowNumber = sheet[snowNumberCol+str(row)].value
            currentShortDesc = sheet[shortDescCol+str(row)].value
            currentDesc = sheet[descCol+str(row)].value
            currentResultInfo = sheet[resolvedCol+str(row)].value
            currentIncident = Incident.Incident(currentSnowNumber, currentShortDesc, currentDesc, currentRowResolvedDate, currentResultInfo, currentLocation)

            incidents.append(currentIncident)

            currentIncident.getHRIandREF()

            if currentIncident.REF == '' or currentIncident.hriNumber == '':
                ### HERE ####

                #IF INCIDENT DOES NOT HAVE REF OR HRI NUMBER IT SEEMS THAT IS TICKET NOT FROM HRI AND 
                #SHOULD NOT BE INCLUDE IN THE REPORT IN THAT CASE DELETE THIS INC OBJECT

            #ELSE ADD THIS INFO TO THE TEMPLATE
            f = open('incidents.txt', 'a')
            f.write(str(currentIncident.snowNumber) + ' ' + currentIncident.hriNumber + ' ' + currentIncident.location + ' ' + currentIncident.REF + '\n')
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
locationCol=''

#getting columns' letters 
getLetters()

#array of incident objects
incidents = []


lastReportDate = pyip.inputDatetime('Date of last report: ', formats=['%Y/%m/%d %H:%M'])

#filter incidents by date provided by user and appending incidents array by filtered objects
dataFilter(lastReportDate)