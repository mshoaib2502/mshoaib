import os
from pathlib import Path
from changeImageName import *

#root = "C:\\Users\\Shoaib\\Desktop\\Z6Copy"
#root = "D:\\Onedrives\\OneDrive - Saima\\OneDrive\\Temp\\SelectedUnorganized"
root = "C:\\Users\\shoaib\\Downloads\\QuickShare"

if os.path.exists(root):
    dir_to_process = root

    for folderName, subs, files in os.walk(dir_to_process):
        splitByDate = change_filename_dummy(folderName)
        folder_format(folderName, splitByDate)
