import pandas as pd
from openpyxl import load_workbook
from pprint import pprint
import numpy as np

file = pd.read_excel('Expatriation Photographs 13.07.21_ZG.xlsx')
file = file.style.applymap(lambda x: "")  # hack to convert file to a styler object


def highlight(index, col):
    file.applymap(lambda x: "background-color: yellow", pd.IndexSlice[index, col])


def getIndicesInColumnsOfVal(colname, val):
    exact = list(file.data.loc[file.data[colname] == val].index)  # finds exact matches
    inList = list(file.data[file.data[colname].str.contains(val + '/', regex=False,
                                                            na=False)].index)  # finds matches in a list e.g. val/val2/val3; val is found
    endOfList = list(file.data[file.data[colname].str.endswith('/' + val,
                                                               na=False)].index)  # finds matches at the end of a list e.g. val2/val3/val; val is found
    return exact + list(set(inList) - set(exact)) + list(
        set(endOfList) - set(inList) - set(exact))  # adding without counting duplicates


def rowIsNotEmpty(index):
    not_na = file.data.notna()
    return not_na.any(axis="columns")[index]


############################FIRST ITERATION###############
def translateEnglishRel(str):
    onto_list = ['daughter-in-law', 'son-in-law', 'grandchild', 'brother-in-law', 'sister-in-law',
                 "sister-in-law's son",
                 "paternal uncle's son", 'paternal uncle', "wife's sister", "maternal uncle's son",
                 "husband's sister",
                 'paternal aunt', "brother-in-law's son", "brother-in-law's daughter", "husband or wife's brother",
                 'brother', "brother's child", "brother's wife", 'grandson', 'granddaughter', 'sister', 'sibling',
                 'mother-in-law', 'daughter', 'mother', 'son', 'self', 'intended', 'father', 'stepmother', 'child',
                 'niece',
                 'nephew', 'husband', 'wife', "wife's brother", "employer", "employee", "non-biological child",
                 "fiancee",
                 "spouse", "one who is under the protection of/in the service of", "brother-in-law's son",
                 'spouses brother', 'child of sibling', 'non-biological daughter', "brother's daughter",
                 'maid, domestic',
                 "brother's son", 'grandmother', 'non-biological son', "sister's child", 'relative',
                 'step paternal uncle', "self and family"]
    str = str.rstrip()
    str = str.lower()
    if not str in onto_list:
        print("Relationshipt not known: " + str)
        return str
    return "relation:" + str


def translateNonEnglishRel(str):
    dic = {"gelin": "daugther-in-law", "damat": "son-in-law", "torun": "grandson", "kayınbirader": "brother-in-law",
           "amcaoğlu": "paternal uncle's son", "amca": "paternal uncle", "baldız": "wife's sister",
           "dayıoğlu": "maternal uncle's son", "görümce": "husband's sister", "hala": "paternal aunt",
           "kayınbiraderinin kızı": "brother-in-law's daughter", "kayınbiraderinin oğlu": "brother-in-law's son",
           "kayın": "husband or wife's brother", "kardeş": "brother", "kardeşinin oğlu": "brother's child",
           "kardeş çocuğu": "brother's child", "biradereşi": "brother's wife", "kiz torun": "granddaughter",
           "bacı": "sister", "abla": "sister", "kızkardeşi": "sister", "kayınvalide": "mother-in-law",
           "kız": "daughter",
           "anne": "mother", "oğlu": "son", "kendi": "self", "erkek eş adayı": "intended", "baba": "father",
           "üvey anne": "stepmother", "çocuk": "child", "kardeşi kızı": "niece", "kardeşi oğlu": "nephew",
           "eşi": "spouse",
           "kardeşi": "sibling", "kızı": "daughter", "kızı ": "daugther", "annesi": "mother", "kız kardeşi": "sister",
           "torunu": "grandchild",
           "gelini": "daughter-in-law", "himayesinde bulunan": "one who is under the protection of/in the service of",
           "çocuğu": "child", "kendi ve ailesi": "self and family", "kaim biraderi oğlu": "brother-in-law's son",
           "kaynı": "spouses brother", "biraderi": "brother", "kayınvalidesi": "mother-in-law",
           "amcası": "paternal uncle",
           "erkek kardeşi": "brother", "ablası, bacısı": "sister", "bacısı": "sister", "kız torunu": "granddaughter",
           "kız çocuğu": "daughter", "kardeşinin eşi": "brother's wife", "halası": "paternal aunt",
           "yeğeni": "child of sibling", "amcası oğlu": "paternal uncle's son", "validesi": "mother",
           "manevi kızı": "non-biological daughter", "kardeşinin kızı": "brother's daughter",
           "kayın validesi": "mother-in-law", "hizmetcisi": "maid, domestic", "görümcesi": "husband's sister",
           "biraderi oğlu": "brother's son", "damadı": "son-in-law", "biraderzadesi": "brother's son",
           "kaynanası": "mother-in-law", "kaynana": "mother-in-law", "erkek kardeşinin kızı": "brother's daughter",
           "biraderi eşi": "brother's wife", "biraderi kızı": "brother's daughter", "büyükannesi": "grandmother",
           "hizmetçisi": "maid, domestic", "manevi oğlu": "non-biological son",
           "kız kardeşinin çocuğu": "sister's child",
           "erkek kardeşinin oğlu": "brother's son", "akrabası": "relative", "üvey amcası": "step paternal uncle"}
    str = str.lower()
    str.replace(" ", "")
    try:
        translation = dic[str]
    except KeyError:
        print("Didn't recognize relation. Might already be english. Calling english translation. " + str)
        return translateEnglishRel(str)
    return "relation:" + translation


def getLeffenFira():
    arr = file.data['      Folder name']
    for i in range(0, len(file.data.index)):
        if not pd.isnull(arr[i]):
            if not arr[i].lower().find('leffen') == -1:
                file.data['      Folder name'][i] = ''
                highlight(i, '      Folder name')
                file.data['Photograph Attached (leffen) '][i - 1] = 'L'
                highlight(i - 1, 'Photograph Attached (leffen) ')
            if not arr[i].lower().find('rar') == -1:  # chose 'rar' as this is in every spelling of firar.
                file.data['      Folder name'][i] = ''
                highlight(i, '      Folder name')
                file.data['Firar-I iade'][i - 1] = 'FI'
                highlight(i - 1, 'Firar-I iade')


def firstNames():
    first_names = pd.read_excel('first_names - Vahakn suggestions (1).xlsx')
    colnames = ['First Name', "Husband's Name", "Father's Name", "Mother's name", "Grandfather's name"]
    for i in range(0, len(first_names.index)):
        if (not pd.isnull(first_names['first_names'][i])) and (
                not pd.isnull(first_names['Vahakn suggestions'][i])) and (
                not first_names['first_names'][i] == first_names['Vahakn suggestions'][i]) and (
                not first_names['Vahakn suggestions'][i] == 'X'):
            val = first_names['first_names'][i]

            # print('Searching ' + val)
            for col in colnames:
                res = getIndicesInColumnsOfVal(col, val)
                for j in res:
                    file.data[col][j] = file.data[col][j] + '/' + first_names['Vahakn suggestions'][i]
                    file.applymap(lambda x: "background-color: yellow", pd.IndexSlice[j, col])


def genderData():
    gender_data = pd.read_excel('gender_data (1).xlsx')
    colnames = ['First Name', "Husband's Name", "Father's Name", "Mother's name", "Grandfather's name"]
    for i in range(0, len(gender_data.index)):
        stringToAdd = ''
        if not pd.isnull(gender_data['Same as'][i]):
            stringToAdd += gender_data['Same as'][i]
            if not pd.isnull(gender_data['name (correct form, seperate multiple with \'/\')'][i]):
                stringToAdd += '/' + gender_data['name (correct form, seperate multiple with \'/\')'][i]
        elif not pd.isnull(gender_data['name (correct form, seperate multiple with \'/\')'][i]):
            stringToAdd += gender_data['name (correct form, seperate multiple with \'/\')'][i]
        if not stringToAdd == '':  # die ziile führt zu problem meinti
            val = gender_data['name'][i]
            print('Searching ' + val)
            for col in colnames:
                res = getIndicesInColumnsOfVal(col, val)

                for j in res:
                    file.data[col][j] = file.data[col][j] + '/' + stringToAdd
                    highlight(j, col)
                    if not pd.isnull(gender_data['gender'][i]) and (
                            gender_data['gender'][i] == 'm' or gender_data['gender'][i] == 'f'):
                        file.data['Gender'][j] = gender_data['gender'][i]
                        highlight(j, 'Gender')


def lastNames():
    last_names = pd.read_excel('POU all last names (1).xlsx')
    columns = ["Turkish Last Name", "Armenian Last name", "Passenger List - Last", "US Documents - Last",
               "Obituary or Gravestone - Last"]
    for i in range(0, len(last_names.index)):
        val = last_names['Ottoman name transliterated into Turkish in Latin script'][i]
        for col in columns:
            res = getIndicesInColumnsOfVal(col, val)
            for j in res:
                if not pd.isnull(last_names['CORRECTION to Column A if necessary'][i]):
                    file.data[col][j] = file.data[col][j].replace(val,
                                                                  last_names['CORRECTION to Column A if necessary'][i])
                    highlight(j, col)
                if not pd.isnull(last_names['Armenian version'][i]):
                    file.data[col][j] = file.data[col][j] + '/' + last_names['Armenian version'][i]
                    highlight(j, col)
                if not pd.isnull(last_names['Alternate spelling'][i]):
                    file.data[col][j] = file.data[col][j] + '/' + last_names['Alternate spelling'][i]
                    highlight(j, col)


def kinships():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data['Kin Relationship'][i]):
            oldval = file.data['Kin Relationship'][i]
            file.data['Kin Relationship'][i] = translateNonEnglishRel(file.data['Kin Relationship'][
                                                                          i])  # some cases have not yet been caught by the translation script (typing differences and actual weird values. Need to talk with Zeynep about this)
            if not oldval == file.data['Kin Relationship'][i]:
                highlight(i, 'Kin Relationship')

    # kinships translated, now assign gender
    gender_list = {'daughter-in-law': 'f', 'son-in-law': 'm', 'brother-in-law': 'm', 'sister-in-law': 'f',
                   "sister-in-law's son": 'm',
                   "paternal uncle's son": 'm', 'paternal uncle': 'm', "wife's sister": 'f',
                   "maternal uncle's son": 'm',
                   "husband's sister": 'f',
                   'paternal aunt': 'f', "brother-in-law's son": 'm', "brother-in-law's daughter": 'f',
                   "husband or wife's brother": 'm',
                   'brother': 'm', "brother's wife": 'f', 'grandson': 'm', 'granddaughter': 'f', 'sister': 'f',
                   'mother-in-law': 'f', 'daughter': 'f', 'mother': 'f', 'son': 'm', 'intended': 'f', 'father': 'm',
                   'stepmother': 'f',
                   'niece': 'f', 'nephew': 'm', 'husband': 'm', 'wife': 'f', "wife's brother": 'm',
                   'spouses brother': 'm', 'non-biological daughter': 'f', "brother's daughter": 'f',
                   'maid, domestic': 'f', "brother's son": 'm', 'grandmother': 'f', 'non-biological son': 'm',
                   'step paternal uncle': 'm'}
    for key in gender_list:
        val = 'relation:' + key
        res = getIndicesInColumnsOfVal('Kin Relationship', val)
        for j in res:
            file.data['Gender'][j] = gender_list[key]
            highlight(j, 'Gender')


def photographersDestination():
    photographer = pd.read_excel('Photographer and Destination Names.xlsx', 'Photographer')
    country = pd.read_excel('Photographer and Destination Names.xlsx', 'Destination Country')
    city = pd.read_excel('Photographer and Destination Names.xlsx', 'Destination City')
    for i in range(0, len(photographer.index)):
        if not pd.isnull(photographer['in Data'][i]) and not pd.isnull(photographer['Correction'][i]):
            val = photographer['in Data'][i]
            res = getIndicesInColumnsOfVal('Photographer Name', val)
            for j in res:
                file.data['Photographer Name'][j] = photographer['Correction'][i]
                highlight(j, 'Photographer Name')
    for i in range(0, len(country.index)):
        if not pd.isnull(country['in Data'][i]) and not pd.isnull(country['Correction'][i]):
            val = country['in Data'][i]
            res = getIndicesInColumnsOfVal('Destination', val)
            for j in res:
                file.data['Destination'][j] = country['Correction'][i]
                highlight(j, 'Destination')
                if not pd.isnull(country['City'][i]):
                    file.data['Destination - city'][j] = country['City'][i]
                    highlight(j, 'Destination - city')
    for i in range(0, len(city.index)):
        if not pd.isnull(city['in Data'][i]) and not pd.isnull(city['Correction'][i]):
            val = city['in Data'][i]
            res = getIndicesInColumnsOfVal('Destination - city', val)
            for j in res:
                file.data['Destination - city'][j] = city['Correction'][i]
                highlight(j, 'Destination - city')


def duplicatesHelper(col, index):
    s = file.data[col][index]
    toDelete = []
    if not pd.isnull(s):
        s = str(s).split('/')
        if (len(s)) == 1:
            return

        for i in range(0, len(s)):
            for j in range(i + 1, len(s)):
                if s[i].strip() == s[j].strip():
                    if not j in toDelete:
                        print('deleting ' + s[i] + ' at position ' + str(j) + '\n')
                        if len(toDelete) == 0:
                            toDelete.append(j)
                        else:
                            k = 0
                            try:
                                while toDelete[k] > j:
                                    k += 1
                                toDelete.insert(k, j)
                            except IndexError:
                                toDelete.append(j)
        if len(toDelete) > 0:
            highlight(index, col)
        for val in toDelete:
            del s[val]
        toReturn = ''

        for i in range(0, len(s) - 1):
            toReturn += s[i].strip() + '/'
        toReturn += s[-1]
        file.data[col][index] = toReturn


def checkDuplicates():
    for col in file.data.columns:
        for i in range(0, len(file.data.index)):
            duplicatesHelper(col, i)


def folderName():
    # make sure that every entry in the folder name column is the exact name of a folder to be extracted - no other information. also make sure that the folder name is on a line with only the folder information, no person info. Remove completely blank lines.

    arr = file.data['      Folder name']
    oldVal = ''
    possibleCols = ['      Folder name', 'Photograph is the same as (if any)', 'DH Page Number',
                    'Photograph Attached (leffen) ',
                    'Wording regarding photography         k=kita (piece)             n=nusha (copy)       p=künye pusulasi (id sheet)\na= aded (unit)',
                    'How many copies of each photograph produced?', 'Firar-I iade',
                    'How many prints enclosed and sent to Istanbul?']
    for i in range(0, len(file.data.index)):
        if not pd.isnull(arr[i]) and arr[i] == oldVal:
            file.data['      Folder name'][i] = ''
            highlight(i, '      Folder name')
        else:
            oldVal = arr[i]
        if not pd.isnull(file.data['      Folder name'][i]) and not file.data['      Folder name'][i] == '':
            for col in file.columns.values.tolist():
                if col in possibleCols:
                    continue
                if not pd.isnull(file.data[col][i]):
                    print('Found value on row ' + col + ', ' + str(i))
                    file.applymap(lambda x: "background-color: red", pd.IndexSlice[i])
                    break

    file.data.dropna(how="all", axis="rows", inplace=True)


def placeNames():
    place = pd.read_excel('Places names POU_vkedits_2SEP (1).xlsx', 'Vilayet names')
    for i in range(0, len(file.index)):
        if not pd.isnull(file['Where are they from?   (vilayet)'][i]):
            res = list(place[place['Sanjak'].str.contains(file['Where are they from?   (vilayet)'][i], regex=False,
                                                          na=False)].index)
            if len(res) == 0:
                print("Didn't find " + str(file['Where are they from?   (vilayet)'][i]))


############################SECOND ITERATION###############
def fixLeffen():
    printL = False
    for i in range(0, len(file.data.index)):
        if not rowIsNotEmpty(i):
            continue
        if not pd.isnull(file.data['      Folder name'][i]):
            printL = False
        if file.data['Photograph Attached (leffen) '][i] == "L":
            printL = True
        if (printL and not file.data['Photograph Attached (leffen) '][i] == "L") or i >= 6560:
            file.data['Photograph Attached (leffen) '][i] = "L"
            highlight(i, 'Photograph Attached (leffen) ')


def mergeTurkishAndArmenian():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Turkish Last Name"][i]) and not pd.isnull(file.data["Armenian Last name"][i]):
            turkish = file.data["Turkish Last Name"][i].split('/')
            armenian = file.data["Armenian Last name"][i].split('/')
            total = turkish + armenian
            s = ""
            for name in total:
                s += name + '/'
            file.data["Turkish Last Name"][i] = s[:-1]
            file.data["Armenian Last name"][i] = ""
            highlight(i, "Turkish Last Name")
            highlight(i, "Armenian Last name")


def genderByName():
    gender_data = pd.read_excel('gender_data (1).xlsx', "people by names")
    for i in range(0, len(gender_data.index)):
        if pd.isnull(gender_data['gender'][i]) or (
                gender_data['gender'][i] != 'm' and gender_data['gender'][i] != 'f'):
            continue
        vals = []
        if not pd.isnull(gender_data['name'][i]):
            vals.append(gender_data['name'][i])
        if not pd.isnull(gender_data['Same as'][i]):
            vals.append(gender_data['Same as'][i])
        if not pd.isnull(gender_data['name (correct form, seperate multiple with \'/\')'][i]):
            vals.append(gender_data['name (correct form, seperate multiple with \'/\')'][i])
        for val in vals:
            res = getIndicesInColumnsOfVal("First Name", val)
            if not res == []:
                pprint("Name: " + val)
                pprint(res)
            for j in res:
                file.data["Gender By Name"][j] = gender_data['gender'][i]
                highlight(j, "Gender By Name")


def genderByRel():
    gender_list = {'daughter-in-law': 'f', 'son-in-law': 'm', 'brother-in-law': 'm', 'sister-in-law': 'f',
                   "sister-in-law's son": 'm',
                   "paternal uncle's son": 'm', 'paternal uncle': 'm', "wife's sister": 'f',
                   "maternal uncle's son": 'm',
                   "husband's sister": 'f',
                   'paternal aunt': 'f', "brother-in-law's son": 'm', "brother-in-law's daughter": 'f',
                   "husband or wife's brother": 'm',
                   'brother': 'm', "brother's wife": 'f', 'grandson': 'm', 'granddaughter': 'f', 'sister': 'f',
                   'mother-in-law': 'f', 'daughter': 'f', 'mother': 'f', 'son': 'm', 'intended': 'f', 'father': 'm',
                   'stepmother': 'f',
                   'niece': 'f', 'nephew': 'm', 'husband': 'm', 'wife': 'f', "wife's brother": 'm',
                   'spouses brother': 'm', 'non-biological daughter': 'f', "brother's daughter": 'f',
                   'maid, domestic': 'f', "brother's son": 'm', 'grandmother': 'f', 'non-biological son': 'm',
                   'step paternal uncle': 'm'}
    for key in gender_list:
        val = 'relation:' + key
        res = getIndicesInColumnsOfVal('Kin Relationship', val)
        for j in res:
            file.data['Gender By Rel'][j] = gender_list[key]
            highlight(j, 'Gender By Rel')


###############THIRD ITERATION#########################
def fixGender():
    val = ""
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Gender"][i]):
            continue
        if not pd.isnull(file.data["Gender By Name"][i]):
            val = file.data["Gender By Name"][i]
        if not pd.isnull(file.data["Gender By Rel"][i]):
            if val != "" and val != file.data["Gender By Rel"][i]:
                file.applymap(lambda x: "background-color: red", pd.IndexSlice[i, "Gender By Name"])
                file.applymap(lambda x: "background-color: red", pd.IndexSlice[i, "Gender By Rel"])
                continue
            val = file.data["Gender By Rel"][i]
        if val != "":
            file.data["Gender"][i] = val
            highlight(i, "Gender")
        val = ""


def markMissingGender():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["First Name"][i]) and pd.isnull(file.data["Gender"][i]):
            file.applymap(lambda x: "background-color: red", pd.IndexSlice[i, "Gender"])


def fillFoldername():
    printName = ""
    for i in range(0, len(file.data.index)):
        if not rowIsNotEmpty(i):
            continue
        if not pd.isnull(file.data['      Folder name'][i]):
            printName = file.data['      Folder name'][i]
        else:
            file.data['      Folder name'][i] = printName
            highlight(i, '      Folder name')


def unifyFolderspelling():
    for i in range(0, len(file.data.index)):
        string = ''
        if not pd.isnull(file.data['      Folder name'][i]) and file.data['      Folder name'][i].find('_') != -1:
            arr = file.data['      Folder name'][i].split('_')
            del (arr[-1])
            j = 0
            while not arr[j].isnumeric():
                string += arr[j] + '.'
                j += 1
            string = string[:-1] + ' '
            while j < len(arr):
                while arr[j][0] == '0':
                    arr[j] = arr[j][1:]
                string += arr[j] + '/'
                j += 1

            file.data['Print ID'][i] = file.data['      Folder name'][i]
            highlight(i, 'Print ID')
            file.data['      Folder name'][i] = string[:-1]
            highlight(i, '      Folder name')


def ftgFoldernames():
    ftg_file = pd.read_excel('FTG Translation.xlsx', 'FTG copies matched with files')
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data['      Folder name'][i]) and file.data['      Folder name'][i].find('FTG') != -1:
            match = list(ftg_file.loc[ftg_file['FTG file name'] == file.data['      Folder name'][i]].index) + list(
                ftg_file.loc[ftg_file['Same glass negative as'] == file.data['      Folder name'][i]].index)
            if len(match) != 1:
                print('ERROR: DID NOT FIND FTG TRANSLATION FOR ' + file.data['      Folder name'][i])
                continue
            term = ftg_file['At least one copy must have originally arrived in Istanbul with this cover letter'][
                match[0]]
            a = term.split(' ')
            file.data['Print ID'][i] = file.data['      Folder name'][i]
            highlight(i, 'Print ID')
            if len(a) > 3:
                file.data['      Folder name'][i] = "NO DH FILE"
                highlight(i, '      Folder name')
                continue
            file.data['      Folder name'][i] = a[0] + ' ' + a[1]
            highlight(i, '      Folder name')


###############FOURTH ITERATION#########################
def cleanBackspaceGender():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["First Name"][i]) and not file.data["First Name"][i].find(
                ' ') == -1 and not pd.isnull(file.data["Gender"][i]) and file.data["Gender"][i] == 'm':
            file.applymap(lambda x: "background-color: red", pd.IndexSlice[i, "Gender"])


def checkHusbandsName():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Husband's Name"][i]):
            if pd.isnull(file.data["Gender"][i]) or not file.data["Gender"][i].lower() == 'f':
                file.applymap(lambda x: "background-color: orange", pd.IndexSlice[i, "Gender"])


def migirdic():
    vals = ["Mıgırdıc", "Mugurdich", "Mıgırdıç", "Mıgırdiç", "Mugurditch"]
    colnames = ['First Name', "Husband's Name", "Father's Name", "Mother's name", "Grandfather's name"]
    for val in vals:
        for col in colnames:
            res = getIndicesInColumnsOfVal(col, val)
            for j in res:
                for printval in vals:
                    if printval == val:
                        continue
                    if file.data[col][j].find(printval) == -1:
                        file.data[col][j] = file.data[col][j] + '/' + printval
                        file.applymap(lambda x: "background-color: yellow", pd.IndexSlice[j, col])


def fillPageNumber():
    printval = ""
    fileval = ""
    for i in range(6600, len(file.data.index)):
        if not pd.isnull(file.data["DH Page Number"][i]):
            printval = file.data["DH Page Number"][i]
            fileval = file.data["Print ID"][i]
        else:
            if not pd.isnull(file.data["Print ID"][i]) and file.data["Print ID"][i] == fileval:
                file.data["DH Page Number"][i] = printval
                highlight(i, "DH Page Number")


def replaceNoDHFile():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["      Folder name"][i]) and file.data["      Folder name"][i] == "NO DH FILE":
            file.data["      Folder name"][i] = "No Known DH File"
            highlight(i, "      Folder name")


def moveHR():
    for i in range(6700, len(file.data.index)):
        if not pd.isnull(file.data["      Folder name"][i]) and file.data["      Folder name"][i][0:2] == "HR":
            if not pd.isnull(file.data["DH Page Number"][i]):
                if file.data['DH Page Number'][i].find('_') != -1:
                    string = ''
                    arr = file.data['DH Page Number'][i].split('_')
                    del (arr[-1])
                    j = 0
                    while not arr[j].isnumeric():
                        string += arr[j] + '.'
                        j += 1
                    string = string[:-1] + ' '
                    while j < len(arr):
                        while arr[j][0] == '0':
                            arr[j] = arr[j][1:]
                        string += arr[j] + '/'
                        j += 1
                    file.data["      Folder name"][i] = string[:-1]
                else:
                    file.data["      Folder name"][i] = file.data['DH Page Number'][i]
            else:
                file.data["      Folder name"][i] = "No Known DH File"
            highlight(i, "      Folder name")


def checkBackSpaces():
    for col in file.data.columns:
        for i in range(0, len(file.data.index)):
            spacesHelper(col, i)


def spacesHelper(col, index):
    s = file.data[col][index]
    if pd.isnull(s):
        return
    try:
        if not s.strip() == s:
            file.applymap(lambda x: "background-color: blue", pd.IndexSlice[index, col])
            file.data[col][index] = s.strip()
    except AttributeError:
        print(s)


############################################FIFTH ITERATION######################################
def daugtherCheck():
    values = ["realtion: daughter", "relation: daughter", "relation:daughter", "relation:daugther"]
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Kin Relationship"][i]) and file.data["Kin Relationship"][i] in values:
            if not pd.isnull(file.data["Husband's Name"][i]):
                file.applymap(lambda x: "background-color: orange", pd.IndexSlice[i, "Husband's Name"])


def fixFTGNames():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Print ID"][i]) and not file.data["Print ID"][i].find("FTG ") == -1:
            file.data["Print ID"][i] = file.data["Print ID"][i].replace("FTG ", "FTG_")
            highlight(i, "Print ID")


def standardizeDate():
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Date (Gregorian) "][i]):
            try:
                a = file.data["Date (Gregorian) "][i].split('.')
            except AttributeError:
                a = [str(file.data["Date (Gregorian) "][i].day), str(file.data["Date (Gregorian) "][i].month),
                     str(file.data["Date (Gregorian) "][i].year)]
            a = [s.replace(' ', '') for s in a]
            if len(a[0]) == 1:
                a[0] = '0' + a[0]
            if len(a[1]) == 1:
                a[1] = '0' + a[1]
            if not len(a[2]) == 4:
                print("problem with ")
                print(a)
                a[2] = a[2].replace('11904', '1904')
            s = a[0] + '.' + a[1] + '.' + a[2]
            if not file.data["Date (Gregorian) "][i] == s:
                file.data["Date (Gregorian) "][i] = s
                highlight(i, "Date (Gregorian) ")


def testLeffen():
    toCheck = []
    oldFile = pd.read_excel('old iterations/Expatriation Photographs 6.4.2021.xlsx', "Photographs attached")
    currVal = ''
    for i in range(0, len(oldFile.index)):
        if not pd.isnull(oldFile['      Folder name'][i]):
            if oldFile['      Folder name'][i].lower().find('leffen') == -1:
                currVal = oldFile['      Folder name'][i]
            else:
                toCheck.append(currVal)
    for i in range(0, len(file.data.index)):
        if not rowIsNotEmpty(i):
            continue
        if not pd.isnull(file.data["      Folder name"][i]):
            if file.data["      Folder name"][i] in toCheck:
                if pd.isnull(file.data["Photograph Attached (leffen) "][i]) or not \
                file.data["Photograph Attached (leffen) "][i] == "L":
                    file.applymap(lambda x: "background-color: red", pd.IndexSlice[i, "Photograph Attached (leffen) "])
            else:
                if not pd.isnull(file.data["Photograph Attached (leffen) "][i]) and i < 6500:
                    file.applymap(lambda x: "background-color: blue", pd.IndexSlice[i, "Photograph Attached (leffen) "])


def createGNColumn():
    currVal = ""
    count = 0
    store = {}
    for i in range(0, len(file.data.index)):
        nmp = False
        value = ""
        if rowIsNotEmpty(i):
            if not pd.isnull(file.data["      Folder name"][i]) and not file.data["      Folder name"][
                                                                            i] == "No Known DH File":
                value = file.data["      Folder name"][i]
            elif not pd.isnull(file.data["Print ID"][i]):
                value = file.data["Print ID"][i]
            elif not pd.isnull(file.data["HR Folder Name"][i]):
                value = file.data["HR Folder Name"][i]
            if not value == currVal:
                store[currVal] = count
                currVal = value
                if value in store:
                    count = store[value]
                else:
                    count = 0
            if not pd.isnull(file.data["How many copies of each photograph produced?"][i]):
                if file.data["How many copies of each photograph produced?"][i] == "NMP":
                    nmp = True
                count += 1
            if not value == "":
                s = "GN_" + value + '_' + str(count)
                if nmp:
                    s += '_NMP'
                file.data["Glass Negative Identifier"][i] = s
                highlight(i, "Glass Negative Identifier")
            if value == "":
                file.data["Glass Negative Identifier"][i] = "No string found"
                file.applymap(lambda x: "background-color: pink", pd.IndexSlice[i, "Glass Negative Identifier"])


def translateRelId():
    dictionary = {"F": "Father", "H": "Husband", "B": "Brother", "S": "Sister", "FIA": "Fiancé", "So": "Son",
                  "MU": "Maternal Uncle", "MA": "Maternal Aunt", "A": "Aunt", "BI": "Brother-in-law",
                  "FL": "Father-in-Law", "X": "Specific relationship unclear"}
    for i in range(0, len(file.data.index)):
        if not pd.isnull(file.data["Joining a family member already abroad (explicit)"][i]):
            if not file.data["Joining a family member already abroad (explicit)"][i] in dictionary:
                file.applymap(lambda x: "background-color: green", pd.IndexSlice[i, "Joining a family member already abroad (explicit)"])
            else:
                file.data["Joining a family member already abroad (explicit)"][i] = dictionary[file.data["Joining a family member already abroad (explicit)"][i]]
                highlight(i, "Joining a family member already abroad (explicit)")

def translateAdresseeAddressor():
    addresse = pd.read_excel('distinctValuesEF.xlsx', "Addresse")
    addressor = pd.read_excel('distinctValuesEF.xlsx', "Addressor")
    for i in range(0, len(addresse.index)):
        if not pd.isnull(addresse["in Data"][i]) and not pd.isnull(addresse["Correction"][i]):
            res = getIndicesInColumnsOfVal("Addressee of document", addresse["in Data"][i])
            for j in res:
                file.data["Addressee of document"][j] = addresse["Correction"][i]
                highlight(j, "Addressee of document")
    for i in range(0, len(addressor.index)):
        if not pd.isnull(addressor["in Data"][i]) and not pd.isnull(addressor["Correction"][i]):
            res = getIndicesInColumnsOfVal("Addressor (who is writing document)", addressor["in Data"][i])
            for j in res:
                file.data["Addressor (who is writing document)"][j] = addressor["Correction"][i]
                highlight(j, "Addressor (who is writing document)")
daugtherCheck()
fixFTGNames()
standardizeDate()
testLeffen()
createGNColumn()
translateRelId()
translateAdresseeAddressor()
file.to_excel("Expatriation Photographs 20.07.21_ES.xlsx", index=False, engine='openpyxl')
