import pandas as pd
from Correct_typing import correct_typing


def loading(rep_source):
    if str(rep_source[-4:]) ==".xls" or str(rep_source[-4:]) == "xlsx":
        source = pd.read_excel(rep_source) # lecture de la source
    elif rep_source[-4:] == ".csv":
        source = pd.read_csv(rep_source, encoding="ansi", sep=";", decimal=',')
    else:
       return
    myresult = [source.columns, source]
    return myresult

def last_occurence(source, mat_salarie, date_ref):
    source = correct_typing(source) # mise en qualité du typage des données
    source = source.sort_values([mat_salarie])
    filter = source.groupby([mat_salarie],sort=False)[date_ref].max() # création d'une série groupby qui va sortir matricule et date max
    filter = pd.DataFrame({mat_salarie : filter.index, date_ref : filter.values}) # on transforme la série en dataframe
    result = pd.merge(source, filter, how='inner') # on réalise une jointure pour filter le résultat de la source
    result[date_ref] = result[date_ref].dt.date # le datetime est transformé en date

    return result

#test = last_occurence(r'C:\Users\Sabri.GASMI\Desktop\Copie de ETATDESORTIE_GEFCOEXP_GTA_AFFECTATION_HORAIRES_CYCLES_20190722150109.xls',"Matricule", "Date Affectation" )

