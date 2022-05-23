#This program will ping each row of an Excel spreadsheet that contains a value

import openpyxl
from openpyxl.styles import colors, Font, Color
from pythonping import ping


#Input the file path of the Excel file
fileLocation = input("Enter the Excel file's path: \n Note - double-enter \ ")
print("Confirmed path is " + fileLocation)


#Input the column letter to ping
column = input("\nEnter the column to ping: ")
print("Column " + column + " will be tested")


#Load the excel workbook
excelWorkbook = openpyxl.load_workbook(fileLocation)

#Sets the first/only Excel sheet in the workbook as the active sheet for further processing in this program
excelSheet = excelWorkbook.active


#Convert the column letter to a column index (example: A = 1, B = 2)
column = openpyxl.utils.column_index_from_string(column)


#Find the last row which contains cell values
maxRow = excelSheet.max_row
print("\nThe number of hosts to run the ping test are: " + str(maxRow))

#Counter variable of pingable hosts
pingSuccessCount = 0

#Ratio variable of pingable hosts to all hosts
pingSuccessRatio = 0

#For each row in the Excel sheet, and each cell in the row, ping the host
for rows in excelSheet.iter_rows(min_row=1, max_col=column, max_row=maxRow):
     for cell in rows:
         host = cell.value
         print("\nPinging " + host + " at " + str(cell))

        #pingResult is an object created from the ping() method
         pingResult = ping(host)
         print(pingResult)
         
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
pingSuccessRatio = round((pingSuccessCount / maxRow)*100,2)
print("\n" + str(pingSuccessRatio) + "% of hosts are reachable")

print("\nPing test is complete. Hosts that are pingable are in standard green text , unreachable hosts are in italic red text. Re-open the Excel file to find the results.")




