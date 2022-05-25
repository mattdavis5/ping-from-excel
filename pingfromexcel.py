#This program will ping each row of an Excel spreadsheet that contains a value

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import colors, Font, Color
from pythonping import ping

while True:
    #Input the file path of the Excel file
    fileLocation = input("Enter the Excel file's path: \n")
    print("\nConfirmed path is " + fileLocation)

    #Load the excel workbook
    try:
        excelWorkbook = openpyxl.load_workbook(fileLocation)
    except InvalidFileException as ife:
        print("\nPlease enter a valid file path. \nAvailable formats are .xlsx, .xlsm, .xltx, .xltm\n")
    except FileNotFoundError as fnf:
        print("\nPlease enter an existing file path.\n")
    else:
        break

#Sets the first/only Excel sheet in the workbook as the active sheet for further processing in this program
excelSheet = excelWorkbook.active

#Input the column letter to ping
column = input("\nEnter the column to ping: ")
print("\nColumn " + column + " will be tested")

#Convert the column letter to a column index (example: A = 1, B = 2)
column = openpyxl.utils.column_index_from_string(column)

#Input which row of Excel sheet to begin testing cells
while True:
    try:
        minRow = int(input("\nEnter the first row with pingable values: "))
        print("\nConfirmed starting row is " + str(minRow))
    except:
        print("\nPlease enter an integer for the starting row")
    else:
        break


#Find the last row which contains cell values
maxRow = excelSheet.max_row

#Find the number of hosts that will be pinged
hostCount = maxRow - minRow + 1
print("\nThe number of hosts to run the ping test are: " + str(hostCount))

#Input file location of log file
while True:
    logLocation = input("\nEnter the log .txt file's path: \n")
    print("\nConfirmed log file path is " + logLocation)

    #Write to log file, create one if does not already exist
    try:
        with open(logLocation, 'w') as f:
            f.write("Log file for ping test from Excel file - " + fileLocation)
    except:
        print("\nPlease enter a valid path for the .txt log file")
    else:
        break

#Append ping test statement to log file
with open(logLocation, 'a') as f:
    f.write("\nPinging " + str(hostCount) + " hosts\n")

#Counter variable of pingable hosts
pingSuccessCount = 0

#Ratio variable of pingable hosts to all hosts
pingSuccessRatio = 0

print("\nRunning the ping test...")

#For each row in the Excel sheet, and each cell in the row, ping the host
for rows in excelSheet.iter_rows(min_row=minRow, max_col=column, max_row=maxRow):
     for cell in rows:
         host = cell.value

        #pingResult is an object created from the ping() method
         pingResult = ping(host)

        #Append ping result to log file
         with open(logLocation, 'a') as f:
            f.write("\n\n\n\nPinging " + host + " at " + str(cell))
            f.write("\n\n" + str(pingResult) + "\n")
            f.write("\n-------------------------------------------------------------")
         
         #If the ping result is successful, change cell font to green and update success counter variable
         if pingResult.success():
            cell.font = Font(color="0000FF00", italic=False)
            pingSuccessCount += 1
         #If ping result is not successful, change cell font to red, italicize font
         else:
            cell.font = Font(color="00FF0000", italic=True)


#Save the Excel workbook to the file location entered previously
excelWorkbook.save(fileLocation)

print("\n" + str(pingSuccessCount) + " hosts are reachable")

#Calculate the percentage of tested hosts that are reachable
pingSuccessRatio = round((pingSuccessCount / hostCount)*100,2)
print("\n" + str(pingSuccessRatio) + "% of hosts are reachable")

#Append ping summary to log file
with open(logLocation, 'a') as f:
    f.write("\n\n\nSummary:\n" + str(pingSuccessCount) + " hosts are reachable")
    f.write("\n" + str(pingSuccessRatio) + "% of hosts are reachable")

print("\nPing test is complete. Please review the log file for detailed ping results, and Excel file for text modifications. \nHosts that are pingable are in standard green text, unreachable hosts are in italic red text in the Excel fil.")