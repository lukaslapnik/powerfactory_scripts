# -*- coding: utf-8 -*-
"""
Created on Tue Aug  8 07:44:08 2023

@author: slluka
"""

import os
import datetime
import pandas as pd
from tkinter import Tk
from tkinter import filedialog

############### PARAMETRI ##################################################################################
#Ali laufamo v engine mode (spyder) ali se izvaja v PowerFactory
use_powerfactory = True
#Zapis imena datoteke z izhodnimi podatki. Na koncu potrebuje .csv končnico.
#Ker na serverju ni inštaliranega MS office ne gre shranjevati direktno v excel ampak samo v csv tipa datoteke.
file_name = "CIM fuel podatki generatorjev.csv"
##########################################################################################################

#Inicializacija powerfactory
if use_powerfactory:
    import powerfactory as pf    
    app = pf.GetApplication()
    app.ClearOutputWindow()
    user = app.GetCurrentUser()
    app.PrintInfo(f"Trenutni uporabnik: {user}")
    activestudycase = app.GetActiveStudyCase()
    app.PrintInfo(f"Trentuni aktivni study case: {activestudycase}")
    scenario = app.GetActiveScenario()
    app.PrintInfo(f"Trenutni aktivni scenarij: {scenario}")
    
###################################Izpis start cajta skripte##############################################
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
if use_powerfactory: app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
else: print("Pričetek izvajanja programa ob " + str(start_time) + ".")
##########################################################################################################

#Dobi mapo kamor shrani podatke
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
app.PrintPlain("Izberi mapo kamor se shranijo izhodni podatki")
folder_directory = filedialog.askdirectory()
#Dobi path datoteke
file_directory = os.path.join(folder_directory, file_name)
#Inicializacija praznega dataaframe kamor se shranjujejo podatki
df_data = pd.DataFrame()
#Dobi vse generatotrje v modelu
generators = app.GetCalcRelevantObjects("*.ElmSym")
#For zanka, uporabljen enumerate da se spremlja kateri po vrsti je - spremenljivka i
for i, generator in enumerate(generators):
    #dobi ime generatorja
    generator_name = generator.loc_name
    #Če obstaja dobi CIM fuel type v virtualni elektrarni, sicer shrani "BREZ PODATKA"
    try:
        generator_type = generator.pBMU.desc
    except:
        generator_type = "BREZ PODATKA"
    #Izpis v okno consola
    app.PrintPlain(f"Generator {generator} CIM fuel {generator_type}, {i+1}/{len(generators)}")
    #Zapis podatkov v dataframe
    df_data.at[generator_name, "TYPE"] = generator_type

#Na koncu izvozi podatke v csv datoteko
df_data.to_csv(file_directory, index=True, encoding='utf-8')

#################### IZPIS URE ################# KONEC #############################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
if use_powerfactory: app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
else: print("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')