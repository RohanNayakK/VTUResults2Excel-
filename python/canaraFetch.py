# import modules
import json
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import sys
import chompjs


# Get sys args
filePath = sys.argv[1]
dept = sys.argv[2]
sem = sys.argv[3]
examType = sys.argv[4]
credit = chompjs.parse_js_object(sys.argv[5])
acdYear = sys.argv[6]


# test data for debugging (Success)
# filePath = "E:"
# dept = "INFORMATION SCIENCE"
# sem = "8"
# examType = "JUNE-JULY"
# js_obj = '{18CS71: 4,18CS72: 4,18CS734: 4,18CSL76: 3,18ME751: 3, 18CS745: 1}'
# credit = chompjs.parse_js_object(js_obj)
# acdYear = "2018-2019"

# # test data for debugging (Fail)
# filePath = "E:"
# dept = "INFORMATION SCIENCE"
# sem = "8"
# examType = "JUNE-JULY"
#added non existing subject code to test error handling
# js_obj = '{18CS71: 4,18CS72: 4,18CS73: 4,18CSL76: 3,18ME751: 3, 18CS745: 1}'
# credit = chompjs.parse_js_object(js_obj)
# acdYear = "2018-2019"


# Constants
extractedJSONPath = "data.json"

templatePath = "./canaraTemplateFormat.xlsx"

subCellAddress = ["D:I", "J:O", "P:U", "V:AA",
                  "AB:AG", "AH:AM", "AN:AS", "AT:AY"]

defaultNumOfSubjects = 8

cellAddressJson = [["D", "E", "F", ],  ["J", "K", "L", ], ["P", "Q", "R", ], ["V", "W", "X"], [
    "AB", "AC", "AD"], ["AH", "AI", "AJ"], ["AN", "AO", "AP"], ["AT", "AU", "AV"]]


creditCellAddress = ["I", "O", "U", "AA", "AG", "AM", "AS", "AY"]

gradePointPairForCredit = ["H", "N", "T", "Z", "AF", "AL", "AR", "AX"]

studentDataStartingRowIndex = 11

def delExtraSub(num):
    for i in range(len(subCellAddress)-1, (len(subCellAddress)-num)-1, -1):
        sht.range(subCellAddress[i]).api.Delete(
            DeleteShiftDirection.xlShiftToLeft)


try:
    # Opening JSON file
    f = open(extractedJSONPath)

    # Returns JSON object as a dictionary
    data = json.load(f)

    # Closing JSON file
    f.close()

    # Open Template
    app = xw.App(visible=False)
    wb = app.books.open(templatePath)
    sht = wb.sheets['Sheet1']

    # set number of subjects in template

    numOfSubjects = len(data[0]['subjects'])

    if(numOfSubjects != defaultNumOfSubjects):
        numToDel = defaultNumOfSubjects - \
            (numOfSubjects % defaultNumOfSubjects)
        delExtraSub(numToDel)

    # Sort Subjects
    sortedSubjectsList = []
    subjectCodeList = []

    for subject in data[0]["subjects"]:
        sortedSubjectsList.append(subject["subjectCode"] + " - " + subject["subjectName"])
        subjectCodeList.append(subject["subjectCode"])
        sortedSubjectsList.sort()

    acdYearFirst = acdYear.split("-")[0]

    sht.range("A5").value = "DEPARTMENT OF "+dept
    sht.range("A7").value = "Results - " + examType +" "+ acdYearFirst  +" - "+ sem + "- SEMESTER - AY- " + acdYear

    for subject in sortedSubjectsList:
        # Subject name cell address
        cellAddressIdx = subCellAddress[sortedSubjectsList.index(subject)]
        cellAddress = cellAddressIdx.split(":")
        cellAddress = cellAddress[0]+"9"
        # Write subject name
        sht.range(cellAddress).value = subject



    count = 0

    lastCountRow = 0

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
            cellAddress = cellAddress+str((studentDataStartingRowIndex+count))

            # CIE Marks
            sht.range(cellAddressJson[subIndex][0]+str(
                studentDataStartingRowIndex+count)).value = subject["iaMarks"]

            # SEE Marks
            sht.range(cellAddressJson[subIndex][1]+str(
                studentDataStartingRowIndex+count)).value = subject["eaMarks"]

            # Total Marks
            sht.range(cellAddressJson[subIndex][2]+str(
                studentDataStartingRowIndex+count)).value = subject["totalMarks"]

            # Credit

            # find subject code in credit list
            creditPoint = credit.get(subCode)
            gradePointValue = (sht.range(
                gradePointPairForCredit[subIndex]+str(studentDataStartingRowIndex+count)).value)
            sht.range(creditCellAddress[subIndex] + str(
                studentDataStartingRowIndex+count)).value = (gradePointValue * creditPoint)

        lastCountRow = studentDataStartingRowIndex+count
        count += 1

    # delete rows from lastCountRow to 150 (remove additional formula rows Grade Point and Grade)
    sht.range("A"+str(lastCountRow+1) +
              ":AY150").api.Delete(DeleteShiftDirection.xlShiftUp)

    # Save file
    wb.save(filePath+"/canaraFormat.xlsx")

    print("Success")

except Exception as e:
    print(e)
    raise e

finally:
    wb.close()

    app.quit()
