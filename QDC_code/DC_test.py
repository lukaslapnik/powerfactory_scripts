# -*- coding: utf-8 -*-
"""
Created on Sat May 28 08:49:17 2022

@author: SSIMON
"""

import datetime
import os
from tkinter import Tk
from tkinter import filedialog
import powerfactory as pf    
import math
import pandas as pd
app = pf.GetApplication()
app.ClearOutputWindow()

##########################################################################################################################################################################
#############################################################################   PARAMETRI   ##############################################################################
##########################################################################################################################################################################

#Izkoristek omrezja (izgube)
izkoristek_omrezja = 0.97 #(1-izgube)

#Imena uvoženih datotek, glej da se sklada z tistim kar nardi skripta za pretvorbo excel->csv. Ne rabis spreminjati
stringMarketDataFile = "Market Data.csv"
stringBorderFlowFile = "Robna vozlisca P.csv"
stringBorderInfoFile = "Robna vozlisca Info.csv"
stringIzbranaPFile = "Izbrana vozlisca P.csv"
stringIzbranaQFile = "Izbrana vozlisca Q.csv"
stringIzbranaInfoFile = "Izbrana vozlisca Info.csv"

##########################################################################################################################################################################
#############################################################################   PARAMETRI   ##############################################################################
##########################################################################################################################################################################


start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
# else: print("Pričetek izvajanja programa ob " + str(start_time) + ".")

# IMPORT PODATKOV
app.PrintPlain("Izberi mapo z vhodnimi podatki (lahko je v ozadju)!")
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
f_input_data_directory = filedialog.askdirectory()
app.PrintPlain("Mapa izbrana!")
#Beri datoteke
app.PrintPlain("Uvoz market datoteke")
dfMD = pd.read_csv(os.path.join(f_input_data_directory, stringMarketDataFile), index_col = [0])
app.PrintPlain("Uvoz podatkov čezmejnih pretokov")
dfCbFlow = pd.read_csv(os.path.join(f_input_data_directory, stringBorderFlowFile), index_col = [0])
dfCbInfo = pd.read_csv(os.path.join(f_input_data_directory, stringBorderInfoFile), index_col = [0])
app.PrintPlain("Uvoz izbranih vozlisc")
dfIzbP = pd.read_csv(os.path.join(f_input_data_directory, stringIzbranaPFile), header = [0], index_col = [0])
dfIzbQ = pd.read_csv(os.path.join(f_input_data_directory, stringIzbranaQFile), header = [0], index_col = [0])
dfIzbInfo = pd.read_csv(os.path.join(f_input_data_directory, stringIzbranaInfoFile), header = [0])
app.PrintPlain("Datoteke uvozene")
# df_select_nodes_info = pd.read_csv(input_path_select_nodes_info, index_col = [1], header = [0])
# return dfMD, dfCbFlow, dfCbInfo, dfIzbP, dfIzbQ, dfIzbInfo 

app.PrintPlain("Racunanje koeficientov generatorjev")
#NAJDI VSA IZBRANA IN VSE GEN
gen_izbrana = []
gen_other = []
market_grid_type_list = dfMD.columns.tolist()
app.PrintPlain(market_grid_type_list)
izbrana_list = dfIzbInfo.columns.tolist()
dgen_grid_type = {}
for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    generator_name = generator.loc_name
    generator_grid = generator.cpGrid.loc_name
    if generator_name in dfIzbQ:
        gen_izbrana.append(generator)
    else:
        try:generator_type = str(''.join(generator.pBMU.desc))
        except:generator_type = str(''.join(generator.desc))
        generator_grid_type = generator_grid + "_" + generator_type
        if generator_grid_type in market_grid_type_list:
            gen_other.append(generator)
            dgen_grid_type[generator] = generator_grid_type    
#Get ratios SUM and then calc ratio
grid_type_sum = {}
for generator in gen_other:
    try: grid_type_sum[dgen_grid_type[generator]] += generator.pgini
    except: grid_type_sum[dgen_grid_type[generator]] = generator.pgini
#Now calc ratio 
gen_ratio = {}
for generator in gen_other:
    try: gen_ratio[generator] = generator.pgini/grid_type_sum[dgen_grid_type[generator]]
    #Če je error bo 0 in nastavimo 0
    except: gen_ratio[generator] = 0
app.PrintPlain("Izracunani koeficienti generatorjev")
app.PrintPlain(gen_ratio)

app.PrintPlain("Racunanje koeficientov bremen")
#NAJDI VSA IZBRANA IN VSE LOAD
load_izbrana = []
load_other = []
# market_grid_type_list = dfMD.columns.tolist() # ZE MAMO
# izbrana_list = dfIzbInfo.columns.tolist() # ZE MAMO
dload_grid_type = {}
for load in app.GetCalcRelevantObjects("*.ElmLod"):
    load_name = load.loc_name
    load_grid = load.cpGrid.loc_name
    if load_name in izbrana_list:
        load_izbrana.append(load)
    else:
        load_grid_type = load_grid + "_LOAD"
        if load_grid_type in market_grid_type_list:
            load_other.append(load)
            dload_grid_type[load] = load_grid_type
#Get ratios SUM and then calc ratio
# grid_type_sum = {} ze mamo od generatorjev, load je kot type grid_LOAD
for load in load_other:
    try: grid_type_sum[dload_grid_type[load]] += load.plini
    except: grid_type_sum[dload_grid_type[load]] = load.plini
#Now calc ratio 
load_ratio = {}
for load in load_other:
    try: load_ratio[load] = load.plini/grid_type_sum[dload_grid_type[load]]
    #Če je error bo 0 in nastavimo 0
    except: load_ratio[load] = 0
app.PrintPlain("Izracunani koeficientov bremen")
app.PrintPlain(load_ratio)

voltagesource_list = dfCbInfo.index.to_list()
robna_list = []
for voltagesource in app.GetCalcRelevantObjects("*.ElmVac"):
    voltagesource_name = voltagesource.loc_name
    voltagesource_grid = voltagesource.cpGrid.loc_name
    if voltagesource_name in voltagesource_list:
        robna_list.append(voltagesource)
app.PrintPlain("Robna vozlisca:")
app.PrintPlain(voltagesource_list)
app.PrintPlain(robna_list)
        
#Uredi podatke, odstej izbrana od market
global df_checking
df_checking = pd.DataFrame()
global df_izbrana_grid_type_sum
global dfMarketSlo
dfMarketSlo = dfMD.filter(regex='SI00')
df_izbrana_grid_type_sum = pd.DataFrame()
df_checking["Market SUM"] = dfMarketSlo["SI00_sum"]

#suma izbranih voslisc po tipu energenta
for izb_voz in dfIzbP.columns:
    # dfIzbInfo.drop(labels = 'Unnamed: 0', axis = 1, inplace = True)
    izb_grid_type = dfIzbInfo.at[0,izb_voz]
    try:
        df_izbrana_grid_type_sum[izb_grid_type] = df_izbrana_grid_type_sum[izb_grid_type] + dfIzbP[izb_voz]
    except:
        df_izbrana_grid_type_sum[izb_grid_type] = dfIzbP[izb_voz]
        
#Mam df sume izbranih vozlisc
#Se kompletno
app.PrintPlain("zacetek obdelave podatkov")
cols_to_sum = [col for col in df_izbrana_grid_type_sum.columns if "LOAD" not in col]
df_izbrana_grid_type_sum["SI00_sum"] = df_izbrana_grid_type_sum.apply(lambda row: row[cols_to_sum].sum(), axis=1)
df_checking["Izbrana SUM"] = df_izbrana_grid_type_sum["SI00_sum"]


delta = dfMarketSlo.sub(df_izbrana_grid_type_sum, fill_value = 0)
list_type_ignore = ["28","29","44","45"]
for column in delta.columns.to_list():
    if not any(type_ignore in column for type_ignore in list_type_ignore):
        delta[column] = delta[column].clip(lower = 0)
        
prefixes = {col[:4] for col in delta.columns}
for prefix in prefixes:
    cols_to_sum = [col for col in delta.columns if col.startswith(prefix) and "sum" not in col and "LOAD" not in col and "Balance" not in col and "Dump" not in col and "DSR" not in col]
    delta["SI00_NEWmarketsum"] = delta.apply(lambda row: row[cols_to_sum].sum(), axis=1)

df_checking["Market-IzbranaSUM"] = delta["SI00_NEWmarketsum"]
df_checking["New Mark+Izb SUM"] = df_checking["Market-IzbranaSUM"] + df_checking["Izbrana SUM"]
df_checking["New DELTA"] = df_checking["New Mark+Izb SUM"] - df_checking["Market SUM"]

for column in dfMD.columns.to_list():
    if column in delta:
        dfMD[column] = delta[column]
        # app.PrintPlain(f"Subtracted {column} in delta from market")
    
    # replace_negatives = lambda x: 0 if x < 0 and x.name != 'SI00_29' else x
    # delta = delta.apply(replace_negatives)
    # delta = delta.applymap(lambda x: 0 if x < 0 and "28" not in x.name and "29" not in x.name and "44" not in x.name and "45" not in x.name else x)
    # dfMarketData = dfMarketData - df_izbrana_grid_type_sum

    # dfMarketData[] = df.apply(lambda row: row[[col for col in df.columns if col.startswith('C')]].sum()
    
####################################################  UVOZ V POWERFACTORY  ##############################################################################
    
app.PrintPlain("Izdelava karakteristik in uvoz podatkov v powerfactory")
#Make time scale for a year in libry folder
fLibrary = app.GetProjectFolder("lib") #Get library folder
timescale_name = "Time Scale"
timescale = fLibrary.SearchObject(timescale_name)
if not timescale:
    try:
        app.PrintPlain("No timescale named " + timescale_name + " exists, creating")
        fLibrary.CreateObject("TriTime", timescale_name)
        timescale = fLibrary.GetContents(timescale_name + ".TriTime")[0]
        app.PrintPlain("Made " + timescale_name + " vector!")
    except:
        app.PrintWarn("Problem creating timescale")
if timescale:
    app.PrintPlain("Timescale vector " + timescale_name + " exists!")
    timescale.SetAttribute("unit", 3)
    timescale_vector_max = int(len(dfMD.index) + 1)
    app.PrintPlain(f"Timescale vector for {timescale_vector_max-1} hours!")
    timescale_vector = list(range(1,timescale_vector_max))
    timescale.SetAttribute("scale", timescale_vector)
    app.PrintPlain("Edited " + timescale_name + " vector!")

i=1

#Se za gen
for generator in gen_izbrana:
    generator_name = generator.GetAttribute("loc_name")
    try:
        #Assign P vector
        for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
        chaPgini = generator.CreateObject("ChaVec", "pgini")
        chaPgini.SetAttribute("scale", timescale)
        chaPgini.SetAttribute("vector", dfIzbP[generator_name].to_list())
        chaPgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPgini} for {generator} ({i}/{len(gen_izbrana)})")
    except: app.PrintPlain(f"Error creating char P for {generator}")
    try:
        #Assign Q vector
        for chaOld in generator.GetContents("qgini*.ChaVec"): chaOld.Delete() 
        chaQgini = generator.CreateObject("ChaVec", "qgini")
        chaQgini.SetAttribute("scale", timescale)
        chaQgini.SetAttribute("vector", dfIzbQ[generator_name].to_list())
        chaQgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaQgini} for {generator} ({i}/{len(gen_izbrana)})")
    except: app.PrintPlain(f"Error creating char Q for {generator}")
    i=i+1

i=1

for generator in gen_other:
    try:
        #Assign P vector
        for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
        chaPgini = generator.CreateObject("ChaVec", "pgini")
        chaPgini.SetAttribute("scale", timescale)
        p_vector = [val * gen_ratio[generator] for val in dfMD[dgen_grid_type[generator]].to_list()]
        chaPgini.SetAttribute("vector", p_vector)
        chaPgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPgini} for {generator} ({i}/{len(gen_other)})")
    except: app.PrintPlain(f"Error creating char P for {generator}")
    i=i+1

i=1
    
for load in load_izbrana:
    load_name = load.GetAttribute("loc_name")
    try:
        #Assign P vector
        for chaOld in load.GetContents("plini*.ChaVec"): chaOld.Delete() 
        chaPlini = load.CreateObject("ChaVec", "plini")
        chaPlini.SetAttribute("scale", timescale)
        chaPlini.SetAttribute("vector", dfIzbP[load_name].to_list())
        chaPlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPlini} for {load} ({i}/{len(load_izbrana)})")
    except: app.PrintPlain(f"Error creating char P for {load}")
    try:
        #Assign Q vector
        for chaOld in load.GetContents("qlini*.ChaVec"): chaOld.Delete() 
        chaQlini = load.CreateObject("ChaVec", "qlini")
        chaQlini.SetAttribute("scale", timescale)
        chaQlini.SetAttribute("vector", dfIzbQ[load_name].to_list())
        chaQlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaQlini} for {load} ({i}/{len(load_izbrana)})")
    except: app.PrintPlain(f"Error creating char Q for {load}")
    i=i+1
    
i=1

for load in load_other:
    load_name = load.GetAttribute("loc_name")
    try:
        #Assign P vector
        for chaOld in load.GetContents("plini*.ChaVec"): chaOld.Delete() 
        chaPlini = load.CreateObject("ChaVec", "plini")
        chaPlini.SetAttribute("scale", timescale)
        p_vector = [val * load_ratio[load] * izkoristek_omrezja for val in dfMD[dload_grid_type[load]].to_list()]
        chaPlini.SetAttribute("vector", p_vector)
        chaPlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPlini} for {load} ({i}/{len(load_other)})")
    except: app.PrintPlain(f"Error creating char P for {load}")
    i=i+1
       
i=1
 
for voltagesource in robna_list:
    voltagesource_name = voltagesource.GetAttribute("loc_name")
    try:
        #Assign P vector
        for chaOld in voltagesource.GetContents("Pgen*.ChaVec"): chaOld.Delete() 
        chaPgen = voltagesource.CreateObject("ChaVec", "Pgen")
        chaPgen.SetAttribute("scale", timescale)
        p_vector = [val * dfCbInfo.at[voltagesource_name,'DELEZ'] * dfCbInfo.at[voltagesource_name,'POMNOZITI'] for  val in dfCbFlow[dfCbInfo.at[voltagesource_name,'MEJA']].to_list()]
        chaPgen.SetAttribute("vector", p_vector)
        chaPgen.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPgen} for {voltagesource} ({i}/{len(robna_list)})")
    except: app.PrintPlain(f"Error creating char P for {voltagesource}")

#################################################################################### MAIN #######################################################

#################### IZPIS URE #################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
