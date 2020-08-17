import openpyxl

phaseFreqs = ['0.5 1', '3 4', '5 6' , '7 8']
ampRange = ['80 200', '200 300']
edfFiles = ['ABC', 'DEF', 'GHI']

#The following block should go inside the for loop at "Compute pac here" section
f = open("ComputedData.txt")
textContent = f.readlines()
print(textContent)

fileName = 'PutData.xlsx'
lastColumn = 1 #To start with the first column
firstIteration = True #The computed PAC for the first edf File (all parameters)
count = 0 #Guide counter for Row Record indexing 
RowRec = [0] * len(phaseFreqs) #Tracking the Row Entries
rowIndex = 0 #For storing RowRec[count] in the Iterations excluding the first iteration 

for edfFile in edfFiles:
    for phaseFreq in phaseFreqs:
        #Change the file in the .m file 
        #Compute pac here
        print('Computing pac for: ' + phaseFreq + " ")
        if (phaseFreqs.index(phaseFreq)+1)%2 == 0:
            print("ampRange: " + ampRange[1])
        else:
            print("ampRange: " + ampRange[0])
        print('Using edf file:' + edfFile )
        wb = openpyxl.load_workbook(fileName) #PutData.xlsx should exist
        ##wb = openpyxl.Workbook()
        ws1 = wb.active
        if firstIteration == True:
            lastRow = ws1.max_row
        else:
            lastRow = RowRec[count]
            
        if lastRow == 1:
            ws1.cell(row=lastRow , column=lastColumn, value=edfFile)
            RowRec[count] = lastRow
        else:
            if firstIteration == True:
                ws1.cell(row=lastRow + 3, column=lastColumn, value=edfFile)
                RowRec[count] = lastRow + 3
            else:
                ws1.cell(row=lastRow, column=lastColumn, value=edfFile)
                RowRec[count] = lastRow

        if firstIteration == True:            
            for text in textContent:
                ws1.append({lastColumn: text})
        else:
            rowIndex = RowRec[count]
            for text in textContent:
                ws1.cell(row=rowIndex + 1, column = lastColumn, value = text)
                rowIndex = rowIndex + 1
        count = count + 1
        print (lastRow)

        wb.save(fileName)
    lastColumn = lastColumn + 1
    count = 0
    firstIteration = False 



### Open excel and your workbook
##col = 2 # column B
##excel=win32com.client.Dispatch("Excel.Application")
##excel.Visible=True # Note: set to false when scripting, only True for this example
##wb=excel.Workbooks.Open('PutData.xlsx')
##ws = wb.Worksheets('Sheet1')
##
###Write text contents to column range
##ws.Range(ws.Cells(col ,1),ws.Cells(col,len(text_contents))).Value = text_contents
##
###Save the workbook and quit
##wb.Close(True)
##excel.Application.Quit() 
