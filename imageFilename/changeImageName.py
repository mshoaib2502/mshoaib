import os
from pathlib import Path

monthDesc = {
    "01": "Jan",
    "02": "Feb",
    "03": "Mar",
    "04": "Apr",
    "05": "May",
    "06": "Jun",
    "07": "Jul",
    "08": "Aug",
    "09": "Sep",
    "10": "Oct",
    "11": "Nov",
    "12": "Dec"
}

def getDateFromFilename(filename):
    if "-" in filename:
        dateTime = filename[0:4] + filename[5:7] + filename[11:13]
    else:
        dateTime = filename[0:8]
    return dateTime

def getSeqChangeFlag(filename):
    if "-" in filename:
        onlySeqChange = 'Y'
    else:
        onlySeqChange = 'N'
    return onlySeqChange



def populateFileDateDict(files):
    fileDateDict = {}
    prevDateTime = ""
    for filename in files:
        dateTime = getDateFromFilename(filename)

        if prevDateTime != dateTime:
            fileDateDict[dateTime] = 1
        else:
            fileDateDict[dateTime] = fileDateDict[dateTime] + 1
        prevDateTime = dateTime
    return fileDateDict


def rename_file(folderName, old_filename, new_filename):
    if new_filename != "":
        old_filenameWP = os.path.join(folderName, old_filename)
        new_filenameWP = os.path.join(folderName, new_filename)
        os.rename(old_filenameWP, new_filenameWP)

def change_filename_dummy(folderName):
    # Appending with _ to avoid error in renaming
    files = os.listdir(folderName)
    files.sort(key=lambda x: os.path.join(folderName, x))

    splitByDate = ''

    for filename in files:
        if (os.path.isdir(os.path.join(folderName, filename))):
            continue

        dateTime = getDateFromFilename(filename)

        if (dateTime.isdigit())==False:
            continue

        filename_wo_ext = os.path.splitext(filename)[0]
        if splitByDate == '':
            folder_desc = get_folder_desc(filename)
            if folder_desc != "":
                splitByDate = 'N'
            else:
                splitByDate = 'Y'

        file_ext = os.path.splitext(filename)[1]
        new_filename = filename_wo_ext + "_" + file_ext
        rename_file(folderName, filename, new_filename)

    return splitByDate

def get_filename_wo_seq(filename):
    filename_wo_ext = os.path.splitext(filename)[0]
    filename_wo_seq_list = filename_wo_ext.split(" ")
    filename_wo_seq = ""
    for i in range(len(filename_wo_seq_list)-1):
        if i == 0:
            filename_wo_seq = filename_wo_seq_list[i]
        else:
            filename_wo_seq = filename_wo_seq + " " + filename_wo_seq_list[i]

    return filename_wo_seq


def file_full_format(dateTime, folderName):
    folderDesc = ""
    if len(os.path.basename(folderName).split(" ", 1)) > 1:
        folderDesc = os.path.basename(folderName).split(" ", 1)[1]

    year = dateTime[0:4]
    month = dateTime[4:6]
    day = dateTime[6:8]

    filename_wo_seq = year + "-" + month + monthDesc[month] + "-" + day
    if folderDesc != "":
        filename_wo_seq = filename_wo_seq + " " + folderDesc
    return filename_wo_seq

def get_folder_desc(filename):
    folderDesc = ""
    filename_wo_seq = get_filename_wo_seq(filename)
    if len(filename_wo_seq.split(" ", 1)) > 1:
        folderDesc = filename_wo_seq.split(" ", 1)[1]
    return folderDesc

def folder_format(folderName, splitByDate):
    prevDateTime = ""
    fileCount = 0
    dirs = os.listdir(folderName)
    dirs.sort(key=lambda x: os.path.join(folderName, x))

    if splitByDate == "Y":
        fileDateDict = populateFileDateDict(dirs)

    for filename in dirs:

        if (os.path.isdir(os.path.join(folderName, filename))):
            continue

        dateTime = getDateFromFilename(filename)

        if (dateTime.isdigit())==False:
            continue
 
        noOfZeroes = str(len(dirs)).__len__()
        fileCount = fileCount + 1

        if splitByDate == "Y":
            if prevDateTime != dateTime:
                fileCount = 1
            prevDateTime = dateTime
            noOfZeroes = str(fileDateDict[dateTime]).__len__()

        if (noOfZeroes < 2):
            noOfZeroes = 2

        onlySeqChange = getSeqChangeFlag(filename)

        if (onlySeqChange == "Y"):
            filename_wo_seq = get_filename_wo_seq(filename)
        else:
            filename_wo_seq = file_full_format(dateTime, folderName)

        file_ext = os.path.splitext(filename)[1]
        new_filename = filename_wo_seq + " " + str(fileCount).rjust(noOfZeroes, '0') + file_ext
        rename_file(folderName, filename, new_filename)
