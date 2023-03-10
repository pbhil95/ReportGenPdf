import pandas as pd 
from IniRead import ConfigFile

pd.set_option('display.max_colwidth', None)

def fetchAllStudentDetails():
    file = ConfigFile.Default['ResultFile']
    sData = pd.read_excel(file, sheet_name=['StudentDetails'])
    return sData


def fetchStudentDetails(roll):
    file = ConfigFile.Default['ResultFile']
    resultData = pd.read_excel(file,
                               sheet_name=['StudentDetails'])
    sRecord = pd.DataFrame()
    for examName in resultData:
        rollColumn = resultData[examName].columns[0]
        x = resultData[examName].loc[resultData[examName][rollColumn] == roll]
        sRecord = sRecord.append(x)
    return sRecord