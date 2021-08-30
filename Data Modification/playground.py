import pandas as pd
import numpy as np
from pprint import pprint
file = pd.read_excel('test.xlsx')
file = file.style.applymap(lambda x: "") #hack to convert file to a styler object
# print(file.data)
# file.data = file.data.dropna(0, how="all")
# print(file.data)


print(file.data)