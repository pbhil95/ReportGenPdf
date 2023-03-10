import pandas as pd 
from IniRead import ConfigFile

def fetchSchoolDetails():
    file = ConfigFile.Default['ResultFile']
    sData = pd.read_excel(file, sheet_name=['SchoolDetails'])
    for sd in sData:
        sdata = sData[sd]
    return sdata