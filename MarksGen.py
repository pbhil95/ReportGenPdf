import pandas as pd
from CoScholasticAreasGradeGen import fetchCoScholasticAreasGrade
from DisciplineGradeGen import fetchDisciplneGrade
from BuildResult import fetchResultSubject
from StudentsDeatailsGen import fetchAllStudentDetails,fetchStudentDetails
from SchoolDetailsGen import fetchSchoolDetails
from IniRead import ConfigFile
def fetchResult(roll):
    file = ConfigFile.Default['ResultFile']
    resultData = pd.read_excel(file, sheet_name=['BestUT1', 'HY', 'HYP',
                                                 'T1UT',
                                                 'BestUT2', 'Y', 'YP',
                                                 'T2UT','GT','GTT','GTP',
                                                 'Final', 'Grade',
                                                 'Rank'])
    sRecord = pd.DataFrame()
    for examName in resultData:
        rollColumn = resultData[examName].columns[0]
        x = resultData[examName].loc[resultData[examName][rollColumn] == roll]
        sRecord = sRecord.append(x)
    sRecord=sRecord.fillna('')
    return sRecord

def fetchOverAll(roll):
    file = ConfigFile.Default['ResultFile']
    resultData = pd.read_excel(file, sheet_name=['Final', 'Grade',
                                                 'Rank'])
    sRecord = pd.DataFrame()
    for examName in resultData:
        rollColumn = resultData[examName].columns[0]
        x = resultData[examName].loc[resultData[examName][rollColumn] == roll]
        sRecord = sRecord.append(x)
    sRecord=sRecord.fillna('')
    return sRecord

def fetchAttendance(roll):
    file = ConfigFile.Default['ResultFile']
    resultData = pd.read_excel(file, sheet_name=['Attendance'])
    sRecord = pd.DataFrame()
    for examName in resultData:
        rollColumn = resultData[examName].columns[0]
        x = resultData[examName].loc[resultData[examName][rollColumn] == roll]
        sRecord = sRecord.append(x)
    sRecord=sRecord.fillna('')
    return sRecord

def MarksGenAll():
    sData = fetchAllStudentDetails()
    rolColumnName = sData['StudentDetails'].columns[0]
    rollColData = sData['StudentDetails'][rolColumnName]
    for roll in rollColData:
        schoolDetails=fetchSchoolDetails()
        sDetails=fetchStudentDetails(roll)
        result = fetchResult(roll)
        overAll=fetchOverAll(roll)
        attendance=fetchAttendance(roll)
        Cograde= fetchCoScholasticAreasGrade(roll)
        DiscGrade= fetchDisciplneGrade(roll)
        fetchResultSubject(roll, schoolDetails, sDetails, 
                           result, overAll, attendance, Cograde, DiscGrade)
