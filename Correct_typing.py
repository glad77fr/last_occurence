import pandas as pd
import re
import numpy as np
from datetime import date
from datetime import datetime

#data = pd.read_csv(r"C:\Users\Sabri.GASMI\Desktop\ALDI\ALDI_ABLIS_STAFF_COMPOSANTS_BASES ASSUJETTIES 042019.csv",encoding="ansi", sep=";")

def convert_date(data):
    try:
       return datetime.strptime(str(data),'%d/%m/%Y')
    except:
        pass

def correct_typing(dataframe):
    nb_columns = len(dataframe.columns)
    r = re.compile(r"^\d*[.,]\d*$")
    j = re.compile(r'^\d*$')
    for y in range(nb_columns):
       for i in range(len(dataframe)):
           if pd.notna(dataframe.iat[i,y]):

               if len(str(dataframe.iat[i,y])) == 10 and str(dataframe.iat[i,y]).count('/') == 2: # Test if is a datetime
                   dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].apply(lambda x: convert_date(x))
                   dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].dt.date
                   #dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].astype('datetime64')
                   break

               if r.match(str(dataframe.iat[i, y])):
                   dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].apply(lambda x: str(x).replace(",", "."))
                   dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].astype('float64')
                   break

               if j.match(str(dataframe.iat[i, y])) and (dataframe[dataframe.columns[y]].dtypes == np.int64) == False:
                   save = dataframe[dataframe.columns[y]]
                   try:
                       dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].apply(lambda x: str(x).replace(",", "."))
                       dataframe[dataframe.columns[y]] = dataframe[dataframe.columns[y]].apply(lambda x: float(x))
                   except ValueError:
                        dataframe[dataframe.columns[y]] = save
               break
    return dataframe