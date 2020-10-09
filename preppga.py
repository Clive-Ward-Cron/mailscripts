#! python3
#
# Clive Ward-Cron
# November 13, 2019
# Recreated February 17, 2020
# Updated October 8, 2020

import openpyxl
import os
import sys
import shutil
import pprint
from datetime import date

cwd = os.getcwd()
isAltFile = "-F"
fileName = "Presley General Agency AZ NV_File"

# print(cwd)

# The first arg is the script name
args = sys.argv[1:]

# OPTIONAL FILE NAME FLAG TEST
# The file name that contains `${job#} ${fileName}.csv`
# can change the default fileName variable by using the
# -F flag followed by the new file name

for i in range(0, len(args)):
    if args[i] == isAltFile:
        fileName = ' '.join(args[i+1:])
        args = args[:i]
        break


"""

Type in the name of the AGA excel file in quotes and the 
program will systematically create the folders, find 
the files, and move the files into the created folders.

It will do this based off of the column headers in the 
AGA document that describe the job number and mail date.

"""

AGAfile = ' '.join(args) + ".xlsx"

path = os.path.join(cwd, AGAfile)

# print(path)

if os.path.exists(path):
    print('Path Exists In Current Folder')
    # print(path)
elif os.path.exists(path=os.path.join(cwd, "AGA DATABASE", AGAfile)):
    print('Path Exists in AGA DATABASE Folder')
    path = os.path.join(cwd, "AGA DATABASE", AGAfile)
    # print(path)
else:
    print('File Not Found In Current Folder Or "AGA DATABASE" Folder')
    sys.exit()

wb = openpyxl.load_workbook(path)
sheet = wb[wb.sheetnames[0]]

"""
Create an Array for the job #'s and an Array for the mail dates
that will use the job # to associate the date.

"""
jobNums = []
mailDates = []
jobCol = ""
dateCol = ""
missingFiles = []
dupeFolders = []
jobs = 0

# Find the job # field and Mail Date fields
for col in tuple(sheet.columns):
    fieldName = col[0].value
    print(fieldName)
    if fieldName == None:
        continue
    if fieldName.capitalize().strip() == "Job #".capitalize() or fieldName.capitalize().strip() == "Job #s".capitalize() or fieldName.capitalize().strip() == "Job Number".capitalize():
        jobCol = col
    elif fieldName.capitalize().strip() == "Mail Date".capitalize():
        dateCol = col

# Ensure that the fields were found
if jobCol == "" or dateCol == "":
    print("proper fields for job # and mail dates were not found.")
    sys.exit()

for row in range(1, len(jobCol)):

    # Ensure there are no None values
    if dateCol[row].value == None or jobCol[row].value == None:
        continue

    mailDate = str(dateCol[row].value.month) + \
        "-" + str(dateCol[row].value.day)
    job = str(jobCol[row].value)
    folderName = mailDate + ' ' + job
    csvName = None
    filePath = None

    # Ensure that the folder doesn't already exist before making one
    if not os.path.exists(folderName):
        os.makedirs(folderName)
    else:
        dupeFolders.append(folderName)

    # Retrieve the file name and add extension
    csvName = job + " " + fileName + ".csv"

    # Determine the file path based off working directory
    filePath = os.path.join(cwd, csvName)

    # Ensure file exists and is a file
    if os.path.isfile(filePath):

        # Check that the folder exists and the file isn't already there
        if os.path.isdir(folderName) and not os.path.isfile(os.path.join(folderName, os.path.basename(filePath))):
            shutil.move(filePath, folderName)
            jobs += 1
        else:
            print()
            print("The file already exists in ", folderName)
    else:
        missingFiles.append(job + "\t" + csvName)
        continue

# Display errors
if len(dupeFolders) > 0:
    print("\nThe following folder(s) already exist in the current directory: ")
    for folder in dupeFolders:
        print(folder)
if len(missingFiles) > 0:
    print("\nThe following file(s) could not be found in the working directory: ")
    for job in missingFiles:
        print(job)
if jobs > 0:
    print("\nJob(s) successfully moved: " + str(jobs))

sys.exit()
