import pandas as pd
from openpyxl import load_workbook
from pprint import pprint


file = pd.read_excel('All renunciation of nationality cases 23.2.2021 (2) (1).xlsx', 'Photographs attached')

def getLeffenFira():
    arr = file['Folder name']
    for i in range(0, len(file.index)):
        if not pd.isnull(arr[i]):
            if not arr[i].lower().find('leffen') == -1:
                file['Folder name'][i] = ''
                file['Photograph Attached (leffen) '][i] = 'L'
            if not arr[i].lower().find('rar') == -1:  # chose 'rar' as this is in every spelling of firar.
                file['Folder name'][i] = ''
                file['Firar-I iade'][i] = 'FI'


def firstNames():
    first_names = pd.read_excel('first_names - Vahakn suggestions (1).xlsx')
    colnames = ['First Name', "Husband's Name", "Father's Name", "Mother's name", "Grandfather's name", "Anchoring Individual"]
    for i in range(0, len(first_names.index)):
        if (not pd.isnull(first_names['first_names'][i])) and (not pd.isnull(first_names['Vahakn suggestions'][i])) and (not first_names['first_names'][i] == first_names['Vahakn suggestions'][i]) and (not first_names['Vahakn suggestions'][i] == 'X'):
            val = first_names['first_names'][i]
            print('Searching ' + val)
            for col in colnames:
                res = list(file[file[col].str.contains(val, regex=False, na=False)].index)
                for j in res:
                    file[col][j] = file[col][j] + '/' + first_names['Vahakn suggestions'][i]


def genderData():
    gender_data = pd.read_excel('gender_data (1).xlsx')
    colnames = ['First Name', "Husband's Name", "Father's Name", "Mother's name", "Grandfather's name",
                "Anchoring Individual"]
    for i in range(0, len(gender_data.index)):
        stringToAdd = ''
        if not pd.isnull(gender_data['Same as'][i]):
            stringToAdd += gender_data['Same as'][i]
            if not pd.isnull(gender_data['name (correct form, seperate multiple with \'/\')'][i]):
                stringToAdd += '/' + gender_data['name (correct form, seperate multiple with \'/\')'][i]
        elif not pd.isnull(gender_data['name (correct form, seperate multiple with \'/\')'][i]):
            stringToAdd += gender_data['name (correct form, seperate multiple with \'/\')'][i]
        if not stringToAdd == '':
            val = gender_data['name'][i]
            print('Searching ' + val)
            for col in colnames:
                res = list(file[file[col].str.contains(val, regex=False, na=False)].index)
                for j in res:
                    file[col][j] = file[col][j] + '/' + stringToAdd
                    if not pd.isnull(gender_data['gender'][i]) and (gender_data['gender'][i] == 'm' or gender_data['gender'][i] == 'f'):
                        file['Gender'][j] = gender_data['gender'][i]




#  TODO: Check for all cells if there are duplicates after / that need to be removed
#getLeffenFira()
#firstNames()
genderData()
file.to_excel("output.xlsx", index="false")

# folderStarts = []
# for i in range(0, len(file.index)):
#     if not pd.isnull(file['Folder name'][i]):
#         val = file['Folder name'][i][0:6]
#         if isinstance(val, str) and len(val) > 5 and val == 'FOLDER':
#             folderStarts.append(i)
# sum = 0
# for j in range(0, len(folderStarts)):
#     sum = 0
#     if j + 1 < len(folderStarts):
#         end = folderStarts[j + 1]
#     else:
#         end = len(file.index)
#     for i in range(folderStarts[j], end):
#         if not pd.isnull(file['Folder name'][i]) and isinstance(file['Folder name'][i], str) and file['Folder name'][i] == 'leffen':
#             file['leffen?'][folderStarts[j]] = 'L'
#             file['Folder name'][i] = ''
#         if not pd.isnull(file['toAdd'][i]):
#             sum += file['toAdd'][i]
#     file['total'][folderStarts[j]] = sum
# file.to_excel("output.xlsx", index="false")
