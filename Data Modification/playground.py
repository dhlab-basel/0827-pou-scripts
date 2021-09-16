import pandas as pd
import numpy as np
from pprint import pprint
file = pd.read_excel('test.xlsx')
print(file)
file = file.drop(0);
file = file.drop(1);
print(file)