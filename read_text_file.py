from pywinauto import application
from pywinauto.keyboard import send_keys
import time
import pyautogui as pg
import openpyxl
from datetime import datetime
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

phaseFreqs = ['0.5 1', '0.5 1', '3 4' , '3 4']
ampRange = ['80 200', '200 300']
edfFiles = ['No 1 for MI.edf', 'No 2 for MI.edf', 'No 3 for MI.edf']

#Create an Excel file to save the results
now = datetime.now()
dt_string = now.strftime("%d-%m-%Y_%H-%M-%S")
fileName = "Results_" + dt_string + ".xlsx"
wb_new = openpyxl.Workbook()
wb_new.save(fileName)

#fileName = 'PutData.xlsx'
lastColumn = 1 #To start with the first column
firstIteration = True #The computed PAC for the first edf File (all parameters)
count = 0 #Guide counter for Row Record indexing 
RowRec = [0] * len(phaseFreqs) #Tracking the Row Entries
rowIndex = 0 #For storing RowRec[count] in the Iterations excluding the first iteration 
ampRangeIndex = 0
IsFirstEdfFile = True

#Windows and Dialogs definition
app = application.Application().connect(path = r"C:\Program Files\MATLAB\R2019b\bin\win64\MATLAB.exe")
dlg = app.window(title_re = ".*MATLAB*")
dlg2 = app.window(title_re = ".*pac_pop_main.*")
dlg3 = app.window(title_re = ".*pac_pop_statsSetUp.*")

#Main Part
for edfFile in edfFiles:
    #Change the edf file name in the .m file
    logging.debug('Entered the edfFile loop' + edfFile)
    with open(r"C:\Users\Samantha\Desktop\Sammy - Ope ECoG\Motohashi Yuunosuke 24.8.2020 R temporal ganglioglioma or DNET\Post\ScriptTest.m", 'r') as file:
        filedata = file.read()

    if IsFirstEdfFile == True:
        IsFirstEdfFile = False
    else:
        filedata = filedata.replace(edfFiles[edfFiles.index(edfFile)-1], edfFile)
        print("File changed")

    with open(r"C:\Users\Samantha\Desktop\Sammy - Ope ECoG\Motohashi Yuunosuke 24.8.2020 R temporal ganglioglioma or DNET\Post\ScriptTest.m", 'w') as file:
        file.write(filedata)
        
    for phaseFreq in phaseFreqs:        
        #Activate Matlab window and run ScriptTest.m
        dlg.set_focus()
        time.sleep(0.3)
        pg.keyDown('ctrl')
        pg.press('0')
        pg.keyUp('ctrl')
        pg.write('ScriptTest')
        pg.press('enter')
        

        #Wait for pac_pop_main window to be opened from the ScriptTest.m
        app.wait_cpu_usage_lower(threshold=2, timeout=None, usage_interval=None)
        time.sleep(2.5)
        waitVar = dlg2.wait("exists enabled visible ready active", timeout=30, retry_interval=20)
        print("PAC pop main Window Opened")
        

        #Enter the Parameters in the pac_pop_man window and press OK button
        pg.press('tab')
        pg.write(phaseFreq)
        pg.press('tab')
        
        if ampRangeIndex > 1:
            ampRangeIndex = 0        
        pg.write(ampRange[ampRangeIndex])
        ampRangeIndex = ampRangeIndex + 1

        pg.press('tab')
        pg.write('5')
        pg.press( 'tab', presses=4, interval=0.3)
        pg.write('0.05')
        pg.press('tab')
        pg.write('500')
        pg.press('tab')
        pg.write('18')
        pg.press( 'tab', presses=2, interval=0.3)
        time.sleep
        pg.press('space')
        print('Parameter Entered')            


        #Press OK button when pac_pop_statsSetUp window appears
        app.wait_cpu_usage_lower(threshold=2, timeout=None, usage_interval=None)
        time.sleep(2.5)
        waitVar = dlg3.wait("exists enabled visible ready active", timeout=30, retry_interval=20)
        pg.press( 'tab', presses=7, interval=0.3)
        pg.press('space')
        
        #open the txt file containing EEG.pac.mi value
        f = open(r"C:\Users\Samantha\Desktop\Sammy - Ope ECoG\Motohashi Yuunosuke 24.8.2020 R temporal ganglioglioma or DNET\Post\saveTest.txt")
        textContent = f.readlines()
        print(textContent)

        #Write data to Excel
        wb = openpyxl.load_workbook(fileName) #xlsx should exist
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
            print("Computed Value:")
            for text in textContent:
                ws1.append({lastColumn: float(text)})                
                print(float(text))
        else:
            rowIndex = RowRec[count]
            print("Computed Value:")
            for text in textContent:
                ws1.cell(row=rowIndex + 1, column = lastColumn, value = float(text))
                rowIndex = rowIndex + 1                
                print(float(text))
        count = count + 1
        print ("LastRow:")
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
