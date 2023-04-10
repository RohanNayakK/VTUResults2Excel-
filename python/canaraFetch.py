# import modules
import json
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import sys

# Opening JSON file
f = open('data.json')

# Returns JSON object as a dictionary
data = json.load(f)

# Closing JSON file
f.close()


#get path to write generated file
filePath = sys.argv[1]

# Open Template
app = xw.App(visible=False)
wb = app.books.open('./canaraTemplateFormat.xlsx')
sht = wb.sheets['Sheet1']

# Constants
subCellAddress = ["D:I", "J:O", "P:U", "V:AA",
                  "AB:AG", "AH:AM", "AN:AS", "AT:AY"]
defaultNumOfSubjects = 8

cellAddressJson = [["D", "E", "F"],  ["J", "K", "L"], ["P", "Q", "R"], ["V", "W", "X"], [
    "AB", "AC", "AD"], ["AH", "AI", "AJ"], ["AN", "AO", "AP"], ["AT", "AU", "AV"]]


# set number of subjects in template

numOfSubjects = len(data[0]['subjects'])


def delExtraSub(num):
    for i in range(len(subCellAddress)-1, (len(subCellAddress)-num)-1, -1):
        sht.range(subCellAddress[i]).api.Delete(
            DeleteShiftDirection.xlShiftToLeft)


if(numOfSubjects != defaultNumOfSubjects):
    numToDel = defaultNumOfSubjects - (numOfSubjects % defaultNumOfSubjects)
    delExtraSub(numToDel)


# Sort Subjects
sortedSubjectsList = []
subjectCodeList = []

for subject in data[0]["subjects"]:
    sortedSubjectsList.append(
        subject["subjectCode"] + " - " + subject["subjectName"])
    subjectCodeList.append(subject["subjectCode"])
sortedSubjectsList.sort()


for subject in sortedSubjectsList:
    # Subject name cell address
    cellAddressIdx = subCellAddress[sortedSubjectsList.index(subject)]
    cellAddress = cellAddressIdx.split(":")
    cellAddress = cellAddress[0]+"9"

    # Write subject name
    sht.range(cellAddress).value = subject


studentDataStartingRowIndex = 11

count = 0

for student in data:
    # Add sl no
    sht.range("A" + str((studentDataStartingRowIndex+count))
              ).value = count+1

    # Add USN
    sht.range("B" + str((studentDataStartingRowIndex+count))
              ).value = student["studentUSN"]

    # Add name
    sht.range("C" + str((studentDataStartingRowIndex+count))
              ).value = student["studentName"]

    for(subject) in student["subjects"]:
        subCode = subject["subjectCode"]
        subIndex = subjectCodeList.index(subCode)
        cellAddress = subCellAddress[subIndex].split(":")[0]
        cellAddress = cellAddress + str((studentDataStartingRowIndex+count))

        # CIE Marks
        sht.range(cellAddressJson[subIndex][0] +
                  str(studentDataStartingRowIndex+count)).value = subject["iaMarks"]

        # SEE Marks
        sht.range(cellAddressJson[subIndex][1] +
                  str(studentDataStartingRowIndex+count)).value = subject["eaMarks"]

        # Total Marks
        sht.range(cellAddressJson[subIndex][2]+str(
            studentDataStartingRowIndex+count)).value = subject["totalMarks"]

    count += 1



wb.save(filePath+"/canaraFormat.xlsx")

wb.close()
