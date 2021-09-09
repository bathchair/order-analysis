import os
import sys
import openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.workbook.workbook import Workbook

def get_application_path():
    filePath = ' '
    filePath = sys.executable #get path of executable
    last_dir = filePath.rfind("/") #find the last directory of executable
    filePath = filePath[:last_dir] #index to this last directory
    return filePath


def prop_calculations(wbOrder: Workbook):
    redFill = PatternFill(start_color = 'FF0000',
                            end_color = 'FF0000',
                            fill_type = 'solid')
    wsOrder = wbOrder.worksheets[0]
    rCount = 0
    for row in wsOrder.iter_rows():
        rCount = rCount + 1
        if rCount < 5:
            continue
        else:
            center = wsOrder.cell(row = rCount, column = 2).value
            left = wsOrder.cell(row = rCount, column = 3).value
            right = wsOrder.cell(row = rCount, column = 4).value

            if center != 0 and center > 15:
                wsOrder.cell(row = rCount, column = 7).value = center
            elif left > 15 or right > 15:
                sum = left + right
                wsOrder.cell(row = rCount, column = 5).value = left/sum       
                wsOrder.cell(row = rCount, column = 6).value = right/sum 
            else:
                wsOrder.cell(row = rCount, column = 1).fill = redFill
                wsOrder.cell(row = rCount, column = 8).fill = redFill


def combined_to_order(filePath):
    with open(filePath + '/IO/Participants.txt') as p:
        rawLines = p.readlines()
        pcps = [] # To store column names

        i = 1
        temp = []
        for line in rawLines:
            line = line.strip() # remove leading/trailing white spaces
            if line != '':
                temp.append(line)
            else:
                pcps.append(temp)
                temp = []
        pcps.append(temp)

    for pcp in pcps:
        ID = pcp[0]
        wbCombined = openpyxl.load_workbook(filePath + "/Combined/" + ID + "_FaceTalk_Combined.xlsx")
        wsCombined = wbCombined.worksheets[2]
        if len(pcp) == 3:
            ordnum = pcp[1]
        else:
            ordnum = pcp[4] 
        lang = pcp[len(pcp) - 1]

        orderFile = "Order " + ordnum + "_FaceTalk_" + lang + ".xlsx"
        wbOrder = openpyxl.load_workbook(filePath + "/Order Template/" + orderFile)
        wsOrder = wbOrder.active

        combinedRows = wsCombined.max_row

        # insert participant ID into file
        wsOrder.cell(row = 1, column= 2).value = ID

        # copying left look
        for i in range(2, combinedRows + 1):
            c = wsCombined.cell(row = i, column = 5)
            wsOrder.cell(row = i + 3, column = 4).value = c.value

        # copying right look
        for i in range(2, combinedRows + 1):
            c = wsCombined.cell(row = i, column = 7)
            wsOrder.cell(row = i + 3, column = 3).value = c.value
        
        # copying center look
        for i in range(2, combinedRows +1):
            c = wsCombined.cell(row = i, column = 9)
            wsOrder.cell(row = i + 3, column = 2).value = c.value

        wbOrder.create_sheet("Values Only")
        wbOrder.save(str(filePath + "/Order/" + ID + "_" + orderFile))

        prop_calculations(wbOrder)

        wbOrder.save(str(filePath + "/Order/" + ID + "_" + orderFile))


def create_data_file(wsData):
    secHeaders = ['Demo Info', 'Familiarization', 'PreNaming', 'Naming (Center Looks)', 'PostNaming', 'Post - Pre naming increase in looking', 'Noise vs NoNoise', 'AV vs AO']
    wsData.cell(row = 1, column = 1).value = secHeaders[0]
    wsData.cell(row = 1, column = 7).value = secHeaders[1]
    wsData.cell(row = 1, column = 9).value = secHeaders[2]
    wsData.cell(row = 1, column = 13).value = secHeaders[3]
    wsData.cell(row = 1, column = 17).value = secHeaders[4]
    wsData.cell(row = 1, column = 21).value = secHeaders[5]
    wsData.cell(row = 1, column = 25).value = secHeaders[6]
    wsData.cell(row = 1, column = 27).value = secHeaders[7]

    dataHeaders = ['Participant #', 'Gender', 'Age at Test', 'Mono/Bilingual?', 'Order', 'O/IP', 
        'Left Prop', 'Right Prop', 
        'AO Noise', 'AO NoNoise', 'AV Noise', 'AV NoNoise', 
        'AO Noise', 'AO NoNoise', 'AV Noise', 'AV NoNoise',
        'AO Noise', 'AO NoNoise', 'AV Noise', 'AV NoNoise',
        'AO Noise', 'AO NoNoise', 'AV Noise', 'AV NoNoise',
        'AO', 'AV',
        'Noise', 'NoNoise']

    for i in range(1, len(dataHeaders) + 1):
        wsData.cell(row = 2, column = i).value = dataHeaders[i - 1]


def extract_order_data(filePath):
    with open(filePath + '/IO/Participants.txt') as p:
        rawLines = p.readlines()
        pcps = [] # To store column names

        i = 1
        temp = []
        for line in rawLines:
            line = line.strip() # remove leading/trailing white spaces
            if line != '':
                temp.append(line)
            else:
                pcps.append(temp)
                temp = []
        pcps.append(temp)

    fileOrder = os.listdir(filePath + "/Order/")

    wbData = openpyxl.Workbook()
    wsData = wbData.active

    create_data_file(wsData)

    dataSpots1 = [32, 16, 40, 24]
    dataSpots2 = [55, 62, 53, 60]

    rowCount = 3

    for pcp in pcps:
        ID = pcp[0]
        if len(pcp) == 3:
            gender = None
            ordnum = pcp[1]
        else:
            gender = pcp[1]
            age = pcp[2]
            monobili = pcp[3]
            ordnum = pcp[4] 
            medium = pcp[5]

        lang = pcp[len(pcp) - 1]

        orderFile = "Order " + ordnum + "_FaceTalk_" + lang + ".xlsx"

        wbOrder = openpyxl.load_workbook(filePath + "/Order/" + ID + "_" + orderFile)
        wsOrder = wbOrder.worksheets[1]

        data = []

        # familiarization data
        data.append(wsOrder.cell(row = 8, column = 12).value)
        data.append(wsOrder.cell(row = 8, column = 13).value)

        # prenaming data
        for r in dataSpots1:
            data.append(wsOrder.cell(row = r, column = 14).value)

        # naming (center looks)
        for r in dataSpots2:
            data.append(wsOrder.cell(row = r, column = 14).value)

        # postnaming data
        for r in dataSpots1:
            data.append(wsOrder.cell(row = r, column = 17).value)

        # post - prenaming data
        for r in dataSpots1:
            data.append(wsOrder.cell(row = r, column = 18).value)

        colInd = 7
        for d in data:
            wsData.cell(row = rowCount, column = 1).value = ID
            if gender:
                wsData.cell(row = rowCount, column = 2).value = gender
                wsData.cell(row = rowCount, column = 3).value = age
                wsData.cell(row = rowCount, column = 4).value = monobili
                wsData.cell(row = rowCount, column = 6).value = medium
            wsData.cell(row = rowCount, column = 5).value = ordnum
            wsData.cell(row = rowCount,  column = colInd).value = d
            colInd = colInd + 1
        rowCount = rowCount + 1

    wbData.save(str(filePath + "/files/" + "results.xlsx"))


def main():
    print('Welcome to Order Analysis! Please wait while we load the filepath.\n')
    filePath = get_application_path()
    ans = input('Do you wish to (1) copy over combined data to order files or (2) collect all data from order files?\n')
    if int(ans) == 1:
        combined_to_order(filePath)
    else:
        extract_order_data(filePath)
        


if __name__ == "__main__":
    main()