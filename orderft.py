import os
import sys
import openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors

def get_application_path():
    filePath = ' '
    filePath = sys.executable #get path of executable
    last_dir = filePath.rfind("/") #find the last directory of executable
    filePath = filePath[:last_dir] #index to this last directory
    return filePath


def combined_to_order(filePath):
    fileCombined = os.listdir(filePath + "/Combined/") #get all files in Input folder
    fileOrder = os.listdir(filePath + "/Order/")
    for file in fileCombined:   
        wbCombined = openpyxl.load_workbook(filePath + "/Combined/" + file) #load excel file
        wsCombined = wbCombined.worksheets[2]   # AVERAGES ACROSS CODERS worksheet
    for file in fileOrder: 
        wbOrder = openpyxl.load_workbook(filePath + "/Order/" + file)
        wsOrder = wbOrder.active

        combinedRows = wsCombined.max_row

        # copying left look
        for i in range(2, combinedRows + 1):
            c = wsCombined.cell(row = i, column = 5)
            wsOrder.cell(row = i + 3, column = 3).value = c.value

        # copying right look
        for i in range(2, combinedRows + 1):
            c = wsCombined.cell(row = i, column = 7)
            wsOrder.cell(row = i + 3, column = 4).value = c.value
        
        # copying center look
        for i in range(2, combinedRows +1):
            c = wsCombined.cell(row = i, column = 9)
            wsOrder.cell(row = i + 3, column = 2).value = c.value

        wbOrder.save(str(filePath + "/Order/" + file))
    

def main():
    filePath = get_application_path()
    combined_to_order(filePath)


if __name__ == "__main__":
    main()