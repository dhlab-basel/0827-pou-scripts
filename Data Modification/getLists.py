from pprint import pprint
import pandas as pd
import xlsxwriter
from pandas import ExcelWriter
file = pd.read_excel('Expatriation Photographs 22.3.2021.xlsx', 'Photographs attached')
def getListofVals(colname):
    toReturn = {'in Data' : [], 'Correction': []}
    for s in file[colname]:
        if not s in toReturn['in Data'] and not pd.isnull(s):
            toReturn['in Data'].append(s)
            toReturn['Correction'].append('')
    return toReturn


photographer = pd.DataFrame(getListofVals('Photographer Name'))
destinationCountry = pd.DataFrame(getListofVals('Destination'))
destinationCity = pd.DataFrame(getListofVals('Destination - city'))

with ExcelWriter('output.xlsx') as writer:
    photographer.to_excel(writer, sheet_name='Photographer', index=False)
    destinationCountry.to_excel(writer, sheet_name='Destination Country', index=False)
    destinationCity.to_excel(writer, sheet_name='Destination City', index=False)
# photographer.to_excel('output.xlsx', sheet_name='Photographer', index=False)
# destinationCountry.to_excel('output.xlsx', sheet_name='Destination Country', index=False)
# destinationCity.to_excel('output.xlsx', sheet_name='Destination City', index=False)