import openpyxl

#One Sample Z Interval
def oneSampleZInterval():
    #path of excel file
    src = input('Enter path of excel file\n')
    try:
        wb = openpyxl.load_workbook(src)
    except(FileNotFoundError):
        print('Path is incorrect')
    #Set sheet
    sheet = input('Enter sheet name')
    wb.se
    #Confidence Level
    confidence = input('Enter confidence level in percent\n')
    column = input('Enter column name \n')
