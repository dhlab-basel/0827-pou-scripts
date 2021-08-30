import pandas as pd
file = pd.read_excel('Expatriation Photographs 20.07.21_ES.xlsx')
def getList():
    store = []
    #delStack = []
    for i in range(0, len(file.index)):
        #data = {'Vilayet': '', 'Sanjak': '', 'Kaza': '', 'Mahalle': ''}
        data = {'Vilayet': '', 'Sanjak': '', 'Kaza': ''}
        if not pd.isnull(file['Where are they from?   (vilayet)'][i]):
            data["Vilayet"] = file['Where are they from?   (vilayet)'][i]
        if not pd.isnull(file['Where are they from (sanjak)'][i]):
            data["Sanjak"] = file['Where are they from (sanjak)'][i]
        if not pd.isnull(file['Where are they from?   (kaza/Nahiye)'][i]):
            data["Kaza"] = file['Where are they from?   (kaza/Nahiye)'][i]
        # if not pd.isnull(file['Where are they from?  neighborhood (mahalle/köy/karye)'][i]):
        #     data["Mahalle"] = file['Where are they from?  neighborhood (mahalle/köy/karye)'][i]
        if data in store:
            pass
        else:
            store.append(data)
    #newFile = {"Vilayet": [], "Sanjak": [], "Kaza": [], "Mahalle": []}
    newFile = {"Vilayet": [], "Sanjak": [], "Kaza": []}
    for d in store:
        newFile["Vilayet"].append(d["Vilayet"])
        newFile["Sanjak"].append(d["Sanjak"])
        newFile["Kaza"].append(d["Kaza"])
        #newFile["Mahalle"].append(d["Mahalle"])
    df = pd.DataFrame(data=newFile)
    df.to_excel("places.xlsx", index=False, engine='openpyxl')
getList()