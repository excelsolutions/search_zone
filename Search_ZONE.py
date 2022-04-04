import pandas # for reading data from xlsx
import os
from termcolor import colored
import numpy as np
path = r'wyszukiwarka.xlsx'
excel_stara_strefa = pandas.read_excel(path, sheet_name='STARE strefy 2022-03-01', usecols=['od kodu', 'do kodu', 'nowa strefa'])
excel_nowa_strefa = pandas.read_excel(path, sheet_name='NOWE strefy 2022-04-01', usecols=['od kodu', 'do kodu', 'nowa strefa'])
excel_przewoznicy = pandas.read_excel(path, sheet_name='Przewoźnicy', usecols=['Pełna strefa', 'Nowa opcja', 'SN'])
excel_kody = pandas.read_excel(path, sheet_name='KODY')

# excel_data.set_index('Waluta')
# excel_data.set_index(["od kodu", "do kodu"], inplace = True, append = True, drop = False)
kod = 55012
nowa_strefa = 0
# PONIŻSZE DZIALA!!!
# df2=excel_data.loc[(excel_data['od kodu'] <= kod) & (excel_data['do kodu'] >= kod), 'od kodu']
os.system('color')
while kod != 0:
    kod = (input(colored('Podaj kod pocztowy: ', 'red', attrs=['reverse', 'blink']) + "\n"))

    try:
        if (kod[2] =="-"):
            kod = str(kod[:2] ) + str(kod[3:])
        kod = int(kod)
        # DANE ADRESOWE
        for index, row in excel_kody.iterrows():
            if (row['PNA'] == kod):
                ulica = excel_kody.loc[index, 'ULICA']
                if str(ulica) == 'nan':
                    miejscowosc = excel_kody.loc[index, 'GMINA']
                else:
                    miejscowosc = excel_kody.loc[index, 'MIEJSCOWOŚĆ'] + ", " + str(excel_kody.loc[index, 'ULICA'])
        print(colored("MIEJSCOWOŚĆ: ", attrs=['bold']), miejscowosc)
        # STARA STREFA
        for index, row in excel_stara_strefa.iterrows():
            if (row['od kodu'] <= kod) and (row['do kodu'] >= kod):
                stara_strefa = excel_stara_strefa.loc[index, 'nowa strefa']
                # print (row['do kodu'], excel_data.loc[index,'do kodu'])
                # print(index, excel_data.loc[index, 'od kodu'], excel_data.loc[index, 'nowa strefa'])
        print(colored("Stara strefa: ", attrs=['bold']), stara_strefa)
        # NOWA STREFA
        for index, row in excel_nowa_strefa.iterrows():
            if (row['od kodu'] <= kod) and (row['do kodu'] >= kod):
                nowa_strefa = excel_nowa_strefa.loc[index, 'nowa strefa']
        print(colored("Nowa strefa: ", attrs=['bold']), nowa_strefa)
        # Przewoźnik
        for index, row in excel_przewoznicy.iterrows():
            if (row['Pełna strefa'] == nowa_strefa):
                przewoznik = excel_przewoznicy.loc[index, 'Nowa opcja']
                przewoznik_sn = excel_przewoznicy.loc[index, 'SN']
        print(colored("Przewoznik: ", attrs=['bold']), przewoznik, colored(", Searchname: ", attrs=['bold']), przewoznik_sn, "\n",  "-------------------", "\n")
    except:
        print('\033[93m' + "Podałeś błedny kod, lub nie można znaleźć rekordu." + '\033[0m')


# print(df2)

# print(excel_data.loc[(excel_data['od kodu'].index[1])])