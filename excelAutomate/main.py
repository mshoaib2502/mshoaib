import os
from write_excel import *

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #folderName = "C:\\Users\\Shoaib\\Desktop\\Z6Copy"
    folderName = "D:\\OneDrives\\Shoaib - OneDrive\\OneDrive\\Saved Files\\Excel Files"
    filename = "Finance - Debt.xlsx"
    dataFile = "data.json"
    fullFilename = os.path.join(folderName, filename)
    fullDataFile = os.path.join(folderName, dataFile)
    #update_json_by_month(fullDataFile)
    write_excel(fullFilename, fullDataFile)
