import pandas as pd
from openpyxl import load_workbook
from pprint import pprint

file = pd.read_excel('Expatriation Photographs 23.3.2021 (1).xlsx', 'Photographs attached')

def getIndicesInColumnsOfVal(colname, val):
    exact = list(file.loc[file[colname] == val].index) #finds exact matches
    inList = list(file[file[colname].str.contains(val + '/', regex=False, na=False)].index) #finds matches in a list e.g. val/val2/val3; val is found
    endOfList = list(file[file[colname].str.endswith('/' + val, na=False)].index) #finds matches at the end of a list e.g. val2/val3/val; val is found
    return exact + list(set(inList) - set(exact)) + list(set(endOfList) - set(inList) - set(exact)) #adding without counting duplicates
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
    arr = file['      Folder name']
    for i in range(0, len(file.index)):
        if not pd.isnull(arr[i]):
            if not arr[i].lower().find('leffen') == -1:
                file['      Folder name'][i] = ''
                file['Photograph Attached (leffen) '][i] = 'L'
            if not arr[i].lower().find('rar') == -1:  # chose 'rar' as this is in every spelling of firar.
                file['      Folder name'][i] = ''
                file['Firar-I iade'][i] = 'FI'


def firstNames():
    first_names = pd.read_excel('first_names - Vahakn suggestions (1).xlsx')
    colnames = ['First Name', "Husband's Name", "Father's Name", "Mother's name", "Grandfather's name"]
    for i in range(0, len(first_names.index)):
        if (not pd.isnull(first_names['first_names'][i])) and (
        not pd.isnull(first_names['Vahakn suggestions'][i])) and (
        not first_names['first_names'][i] == first_names['Vahakn suggestions'][i]) and (
        not first_names['Vahakn suggestions'][i] == 'X'):
            val = first_names['first_names'][i]

            #print('Searching ' + val)
            for col in colnames:
                res = getIndicesInColumnsOfVal(col, val)
                for j in res:
                    file[col][j] = file[col][j] + '/' + first_names['Vahakn suggestions'][i]


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
        if not stringToAdd == '':
            val = gender_data['name'][i]
            print('Searching ' + val)
            for col in colnames:
                res = getIndicesInColumnsOfVal(col, val)

                for j in res:
                    file[col][j] = file[col][j] + '/' + stringToAdd
                    if not pd.isnull(gender_data['gender'][i]) and (
                            gender_data['gender'][i] == 'm' or gender_data['gender'][i] == 'f'):
                        file['Gender'][j] = gender_data['gender'][i]


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
                    file[col][j] = file[col][j].replace(val, last_names['CORRECTION to Column A if necessary'][i])
                if not pd.isnull(last_names['Armenian version'][i]):
                    file[col][j] = file[col][j] + '/' + last_names['Armenian version'][i]
                if not pd.isnull(last_names['Alternate spelling'][i]):
                    file[col][j] = file[col][j] + '/' + last_names['Alternate spelling'][i]


def kinships():
    for i in range(0, len(file.index)):
        if not pd.isnull(file['Kin Relationship'][i]):
            file['Kin Relationship'][i] = translateNonEnglishRel(file['Kin Relationship'][
                                                                     i])  # some cases have not yet been caught by the translation script (typing differences and actual weird values. Need to talk with Zeynep about this)

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
        print(res)
        print()
        for j in res:
            file['Gender'][j] = gender_list[key]

def photographersDestination():
    photographer = pd.read_excel('Photographer and Destination Names.xlsx', 'Photographer')
    country = pd.read_excel('Photographer and Destination Names.xlsx', 'Destination Country')
    city = pd.read_excel('Photographer and Destination Names.xlsx', 'Destination City')
    for i in range(0, len(photographer.index)):
        if not pd.isnull(photographer['in Data'][i]) and not pd.isnull(photographer['Correction'][i]):
            val = photographer['in Data'][i]
            res = getIndicesInColumnsOfVal('Photographer Name', val)
            for j in res:
                file['Photographer Name'][j] = photographer['Correction'][i]
    for i in range(0, len(country.index)):
        if not pd.isnull(country['in Data'][i]) and not pd.isnull(country['Correction'][i]):
            val = country['in Data'][i]
            res = getIndicesInColumnsOfVal('Destination', val)
            for j in res:
                file['Destination'][j] = country['Correction'][i]
                if not pd.isnull(country['City'][i]):
                    file['Destination - city'][j] = country['City'][i]
    for i in range(0, len(city.index)):
        if not pd.isnull(city['in Data'][i]) and not pd.isnull(city['Correction'][i]):
            val = city['in Data'][i]
            res = getIndicesInColumnsOfVal('Destination - city', val)
            for j in res:
                file['Destination - city'][j] = city['Correction'][i]
def duplicatesHelper(s):
    toDelete = []
    if not pd.isnull(s):
        s = str(s).split('/')
        if (len(s)) == 1:
            return s[0]

        for i in range(0, len(s)):
            for j in range(i + 1, len(s)):
                if s[i] == s[j]:
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
        for val in toDelete:
            del s[val]
        toReturn = ''

        for i in range(0, len(s) - 1):
            toReturn += s[i] + '/'
        toReturn += s[-1]
        return toReturn
    return s

def checkDuplicates():
    file.applymap(duplicatesHelper)


# def placeNames():
#     place = pd.read_excel('Places names POU_vkedits_2SEP (1).xlsx', 'Vilayet names')
#     for i in range(0, len(file.index)):
#         if not pd.isnull(file['Where are they from?   (vilayet)'][i]):
#             res = list(place[place['Sanjak'].str.contains(file['Where are they from?   (vilayet)'][i], regex=False, na=False)].index)
#             if len(res) == 0:
#                 print("Didn't find " + str(file['Where are they from?   (vilayet)'][i]))



getLeffenFira()
firstNames()
genderData()
lastNames()
kinships()
photographersDestination()
checkDuplicates()
file.to_excel("POU import file.xlsx", index="false")
