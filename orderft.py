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

    # fileCombined = os.listdir(filePath + "/Combined/") #get all files in Input folder
    # fileOrder = os.listdir(filePath + "/Order/")

    for pcp in pcps:
        ID = pcp[0]
        wbCombined = openpyxl.load_workbook(filePath + "/Combined/" + ID + "_FaceTalk_Combined.xlsx")
        if len(pcp) == 2:
            ordnum = pcp[1]
        else:
            ordnum = pcp[4] 

    # PAUSE HERE: find a way to find order file with number ONLY (may need to add identifier for language)
    # have not redirected file returns/outputs yet

    for file in fileCombined:   
        wbCombined = openpyxl.load_workbook(filePath + "/Combined/" + file) #load excel file
        wsCombined = wbCombined.worksheets[2]
    for file in fileOrder: 
        if(file == ".DS_Store"):
            continue
        wbOrder = openpyxl.load_workbook(filePath + "/Order Template/" + file)
        wsOrder = wbOrder.active

        combinedRows = wsCombined.max_row

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
        wbOrder.save(str(filePath + "/Order/" + file))

def extract_order_data(filePath):
    fileOrder = os.listdir(filePath + "/Order/")

    wbData = openpyxl.Workbook()
    wsData = wbData.active

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

    dataSpots1 = [32, 16, 40, 24]
    dataSpots2 = [55, 62, 53, 60]

    for file in fileOrder:
        if(file == ".DS_Store"):
            continue
        wbOrder = openpyxl.load_workbook(filePath + "/Order/" + file)
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
            wsData.cell(row = 3,  column = colInd).value = d
            colInd = colInd + 1

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