from pprint import pprint
import pandas as pd
import xlsxwriter
from pandas import ExcelWriter
file = pd.read_excel('Expatriation Photographs 13.07.21_ZG.xlsx')
def getListofVals(colname):
    toReturn = {'in Data' : [], 'Correction': []}
    for s in file[colname]:
        if not s in toReturn['in Data'] and not pd.isnull(s):
            toReturn['in Data'].append(s)
            toReturn['Correction'].append('')
    return toReturn


addressor = pd.DataFrame(getListofVals('Addressor (who is writing document)'))
addressee = pd.DataFrame(getListofVals('Addressee of document'))

with ExcelWriter('output.xlsx') as writer:
    addressor.to_excel(writer, sheet_name='Addressor', index=False)
    addressee.to_excel(writer, sheet_name='Addresse', index=False)