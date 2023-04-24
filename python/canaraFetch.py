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
totalCreditsSum = sys.argv[7]


#########START: This is for testing purpose only##########

# test data for debugging (Success)
# filePath = "E:"
# dept = "INFORMATION SCIENCE"
# sem = "7"
# examType = "JUNE-JULY"
# js_obj = '{18CS71: 4,18CS72: 4,18CS734: 3,18CSL76: 2,18ME751: 3, 18CS745: 3}'
# credit = chompjs.parse_js_object(js_obj)
# acdYear = "2018-2019"
# totalCreditsSum = 19

# # test data for debugging (Fail)
# filePath = "E:"
# dept = "INFORMATION SCIENCE"
# sem = "7"
# examType = "JUNE-JULY"
# added non existing subject code to test error handling
# js_obj = '{18CS71: 4,18CS72: 4,18CS73: 4,18CSL76: 3,18ME751: 3, 18CS745: 1}'
# credit = chompjs.parse_js_object(js_obj)
# acdYear = "2018-2019"
# totalCreditsSum = 19


#########END :This is for testing purpose only##########

# Constants
extractedJSONPath = "data.json"

templatePath = "./canaraTemplateFormat.xlsx"

subCellAddress = ["D:I", "J:O", "P:U", "V:AA",
                  "AB:AG", "AH:AM", "AN:AS", "AT:AY"]

lastFourCellsAddress = [
    "E,F,G",
    "K,L,M",
    "Q,R,S",
    "W,X,Y",
    "AC,AD,AE",
    "AI,AJ,AK",
    "AO,AP,AQ",
    "AU,AV,AW",
]

defaultNumOfSubjects = 8

cellAddressJson = [["D", "E", "F", "G"],  ["J", "K", "L",  "N"], ["P", "Q", "R", "S"], ["V", "W", "X", "Y"], [
    "AB", "AC", "AD", "AE"], ["AH", "AI", "AJ", "AK"], ["AN", "AO", "AP", "AQ"], ["AT", "AU", "AV", "AW"]]


creditCellAddress = ["I", "O", "U", "AA", "AG", "AM", "AS", "AY"]

gradePointPairForCredit = ["H", "N", "T", "Z", "AF", "AL", "AR", "AX"]

studentDataStartingRowIndex = 11

lastDeletedSubIndex = 7

# Default Cell Addresses

totalMarksCellAddress = "AZ"

sgpaCellAddress = "BA"

percentageCellAddress = "BB"

classCellAddress = "BC"

# Functions

# Function to delete extra subjects cells from template


def delExtraSub(num):
    for i in range(len(subCellAddress)-1, (len(subCellAddress)-num)-1, -1):
        sht.range(subCellAddress[i]).api.Delete(
            DeleteShiftDirection.xlShiftToLeft)
        lastDeletedSubIndex = i

    global totalMarksCellAddress
    global sgpaCellAddress
    global percentageCellAddress
    global classCellAddress

    totalMarksCellAddress = subCellAddress[lastDeletedSubIndex].split(":")[0]

    # sgpaCellAddress (next cell to right of totalMarksCellAddress)
    sgpaCellAddress = lastFourCellsAddress[lastDeletedSubIndex].split(",")[0]

    # percentageCellAddress (next cell to right of sgpaCellAddress)
    percentageCellAddress = lastFourCellsAddress[lastDeletedSubIndex].split(",")[
        1]

    # classCellAddress (next cell to right of percentageCellAddress)
    classCellAddress = lastFourCellsAddress[lastDeletedSubIndex].split(",")[2]


# Function to get grade class ( FCD, FC, SC, Fail )
def getGradeClass(percent):
    if(percent >= 70 and percent <= 100):
        return "FCD"
    elif(percent >= 60 and percent < 70):
        return "FC"
    elif(percent >= 35 and percent < 60):
        return "SC"
    else:
        return "Fail"


# Function to get acronym of string

def getAcronym(s):
    tokens = s.split()
    string = ""
    for word in tokens:
        if word != "AND":
            string += str(word[0])
    return string.upper()

# Try & Except to handle errors


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
        sortedSubjectsList.append(
            subject["subjectCode"] + " - " + subject["subjectName"])
        subjectCodeList.append(subject["subjectCode"])
        sortedSubjectsList.sort()

    acdYearFirst = acdYear.split("-")[0]

    sht.range("A5").value = "DEPARTMENT OF "+dept
    sht.range("A7").value = "Results - " + examType + " " + \
        acdYearFirst + " - " + sem + "- SEMESTER - AY- " + acdYear

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
        failedInAnySubject = False

        # Add sl no
        sht.range("A" + str((studentDataStartingRowIndex+count))
                  ).value = count+1

        # Add USN
        sht.range("B" + str((studentDataStartingRowIndex+count))
                  ).value = student["studentUSN"]

        # Add name
        sht.range("C" + str((studentDataStartingRowIndex+count))
                  ).value = student["studentName"]

        # Initialize totalCreditsPoints to 0
        totalCreditsPoints = 0

        for(subject) in student["subjects"]:
            # Find subject code in subject code list

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

            # Check if SEE marks is not greater than or equal to 21 then set fail manually
            if int(subject["eaMarks"]) < 21:
                failedInAnySubject = True
                # Go to next cell to right ( Grade Cell ) and set value to F
                sht.range(cellAddressJson[subIndex][3]+str(
                    studentDataStartingRowIndex+count)).value = "F"

            # Credit

            # Find subject code in credit list

            creditPoint = credit.get(subCode)
            gradePointValue = (sht.range(
                gradePointPairForCredit[subIndex]+str(studentDataStartingRowIndex+count)).value)
            sht.range(creditCellAddress[subIndex] + str(
                studentDataStartingRowIndex+count)).value = (gradePointValue * creditPoint)
            totalCreditsPoints += (int(gradePointValue) * int(creditPoint))

        # Total Grade Points ( Sum of all subject credits)
        sht.range(totalMarksCellAddress +
                  str((studentDataStartingRowIndex+count))).value = totalCreditsPoints

        # SGPA
        sht.range(sgpaCellAddress + str((studentDataStartingRowIndex+count))
                  ).value = round(int(totalCreditsPoints) / int(totalCreditsSum), 2)

        # Percentage
        percentageValue = round(
            (int(totalCreditsPoints) / int(totalCreditsSum)), 2) * 10

        sht.range(percentageCellAddress +
                  str((studentDataStartingRowIndex+count))).value = percentageValue

        # Class
        sht.range(classCellAddress + str((studentDataStartingRowIndex+count))
                  ).value = getGradeClass(percentageValue)

        # Check if failed in any subject then set class to fail manually
        if failedInAnySubject:
            sht.range(classCellAddress + str((studentDataStartingRowIndex+count))
                      ).value = "Fail"

        lastCountRow = studentDataStartingRowIndex+count
        count += 1

    # delete rows from lastCountRow to 150 (remove additional formula rows Grade Point and Grade)
    sht.range("A"+str(lastCountRow+1) +
              ":BC150").api.Delete(DeleteShiftDirection.xlShiftUp)

    # get dept accronym
    deptAccro = getAcronym(dept)

    # rename sheet
    sht.name = str(sem)+" Sem " + deptAccro + " Results"

    # Save file
    wb.save(filePath+"/canaraFormat.xlsx")

    print("Success")

except Exception as e:
    print("Error!")
    print(e)
    raise e

finally:
    # Reset focus to A1
    sht.range("A2").select()

    # Close file and quit app (will quit even if error occurs)

    wb.close()

    app.quit()
