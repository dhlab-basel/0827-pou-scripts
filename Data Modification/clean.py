import pandas as pd
from pprint import pprint
file = pd.read_excel('helper files/Places names POU_vkedits_2SEP (1).xlsx', 'Vilayet names')

def removeDuplicates():
    store = []
    delStack = []
    for i in range(0, len(file.index)):
        data = {'sanjak': '', 'vilayet': '', 'correction': '', 'sanarm': '', 'vilarm' : ''}
        if not pd.isnull(file['Sanjak'][i]):
            data["sanjak"] = file['Sanjak'][i]
        if not pd.isnull(file['Vilayet'][i]):
            data["vilayet"] = file['Vilayet'][i]
        if not pd.isnull(file['Correction if necessary'][i]):
            data["correction"] = file['Correction if necessary'][i]
        if not pd.isnull(file['Sanjak Armenian Name'][i]):
            data["sanarm"] = file['Sanjak Armenian Name'][i]
        if not pd.isnull(file['Sanjak Armenian Name'][i]):
            data["vilarm"] = file['Vilayet Armenian Name'][i]
        if data in store:
            delStack.append(i)
        else:
            store.append(data)

    while len(delStack) > 0:
        j = delStack.pop()
        file.drop([file.index[j]], inplace=True)

def checkBackSpaces():
    for col in file.columns:
        for i in range(0, len(file.index)):
            spacesHelper(col, i)


def spacesHelper(col, index):
    s = file[col][index]
    if pd.isnull(s):
        return
    try:
        if not s.strip() == s:
            file[col][index] = s.strip()
    except AttributeError:
        print(s)
checkBackSpaces()
removeDuplicates()
file.to_excel("test.xlsx", index=False, engine='openpyxl')