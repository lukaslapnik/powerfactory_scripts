# -*- coding: utf-8 -*-
"""
Created on Sat May 28 08:49:17 2022

@author: lukc
"""
# New script with optimisations

import pandas as pd
import datetime
import sys
import math
#import numpy
import os
from os import listdir
import glob
import time
from datetime import datetime as dt
from datetime import timedelta as td
from tkinter import Tk
from tkinter import filedialog

############### PARAMETRI ##################################################################################
#Ali laufamo v engine mode (spyder) ali se izvaja v PowerFactory
engine_mode = False
# ČE JE TRUE ŠE ROČNO DEFINIRAJ TOČNO IME PROJEKTA
define_project_name = "ENTSO-E_NT2030_RNPS_2023-2032_v18(4)"

#Parametri za izračun jalovih moči za gen, load, vac..... načeloma če delamo DC loadflow ni važno
# Za AC loadflow je treba porihtat oz najt neke boljše načine dodeljevanja jalovih. 
change_Q_params = False
orignal_cosfi = True
generator_PQ_ratio = 0.2
load_PQ_ratio = 0.1
voltagesource_PQ_ratio = 0
grid_losses = 0.03

#Imena uvoženih datotek, glej da se sklada z tistim kar nardi skripta za pretvorbo excel->csv
stringMarketDataFile = "Market Data.csv"
stringBorderFlowFile = "Border Flow Data.csv"
stringBorderInfoFile = "Border Flow Parameters.csv"
stringIzbranaPFile = "Izbrana vozlisca P.csv"
stringIzbranaQFile = "Izbrana vozlisca Q.csv"
stringIzbranaInfoFile = "Izbrana vozlisca Info.csv"

#Spreminjanje 
# grids_modify_values = ['UKNI','UK00','UA02','UA01','TR00','TN00','SK00','SI00','SE04','SE03','SE02','SE01','SA00','RU00',
#                         'RS00','RO00','PT00','PS00','PL00','NSW0','NOS0','NON1','NOM1','NL00','MT00','MK00','ME00','MD00',
#                         'MA00','LY00','LV00','LUV1','LUG1','LUF1','LUB1','LT00','ITSI','ITSA','ITS1','ITN1','ITCS','ITCN',
#                         'ITCA','IS00','IL00','IE00','HU00','HR00','GR03','GR00','FR15','FR00','FI00','ES00','ELES Interconnectios',
#                         'EG00','EE00','DZ00','DKW1','DKKF','DEKF','DE00','CZ00','CY00','CH00','BG00','BE00','BA00','AT00','AL00']

grids_modify_values = ['AT00','SI00','RS00',
                        'ITN1','HU00','HR00',
                        'ELES Interconnectios','BA00']

grids_write_results = ['SI00', 'ELES Interconnectios']

##########################################################################################################

if engine_mode:
    print("Running in engine mode")
    sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.9")

import powerfactory as pf    
app = pf.GetApplication()
ldf = app.GetFromStudyCase("ComLdf")
qds = app.GetFromStudyCase("ComStatsim")
app.ClearOutputWindow()
mloads = app.GetCalcRelevantObjects("*.ElmLod")
mvoltagesources = app.GetCalcRelevantObjects("*.ElmVac")
fChars = app.GetProjectFolder("chars") #Characters folder

if engine_mode:
    #Če je engine mode funkcije menjamo za navadn print
    app.PrintPlain = print
    app.PrintInfo = print
    app.PrintWarn = print
    app.PrintError = print
    
    #Ime projekta
    app.ActivateProject(define_project_name)
    prj = app.GetActiveProject()
    activestudycase = app.GetActiveStudyCase()
    scenario = app.GetActiveScenario()
    
###################################Izpis start cajta skripte##############################################
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
# else: print("Pričetek izvajanja programa ob " + str(start_time) + ".")
##########################################################################################################

#CREATE AND ACTIVATE A NEW STUDY CASE
fscenarios = app.GetProjectFolder("scen")
scenario = fscenarios.GetContents("Default operation scenario.*")[0] #Characters folder
new_scenario = fscenarios.AddCopy(scenario)
new_scenario.SetAttribute("loc_name", "NOVA STUDIJA - PREIMENUJ")
new_scenario.Activate()
app.PrintPlain(f"Created a new scenario: {new_scenario}")

# IMPORT PODATKOV
app.PrintPlain("Select input data folder (may be hidden behind main screen)!")
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
folder_directory = filedialog.askdirectory()
#Beri datoteke
app.PrintPlain("Importing market file")
dfMD = pd.read_csv(os.path.join(folder_directory, stringMarketDataFile), index_col = [0])
app.PrintPlain("Importing border flow files")
dfCbFlow = pd.read_csv(os.path.join(folder_directory, stringBorderFlowFile), index_col = [0])
dfCbInfo = pd.read_csv(os.path.join(folder_directory, stringBorderInfoFile), index_col = [0])
app.PrintPlain("Importing izbrana vozlisca")
dfIzbP = pd.read_csv(os.path.join(folder_directory, stringIzbranaPFile), header = [0], index_col = [0])
dfIzbQ = pd.read_csv(os.path.join(folder_directory, stringIzbranaQFile), header = [0], index_col = [0])
dfIzbInfo = pd.read_csv(os.path.join(folder_directory, stringIzbranaInfoFile), header = [0])
app.PrintPlain("Files imported")
# df_select_nodes_info = pd.read_csv(input_path_select_nodes_info, index_col = [1], header = [0])
# return dfMD, dfCbFlow, dfCbInfo, dfIzbP, dfIzbQ, dfIzbInfo 

gen_izbrana = []
gen_other = []
market_grid_type_list = dfMD.columns.tolist()
gen_grid_type = {}
gen_ratio = {}

#NAJDI VSA IZBRANA IN VSE GEN
for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    generator_name = generator.loc_name
    generator_grid = generator.cpGrid.loc_name
    if generator_name in dfIzbQ:
        gen_izbrana.append(generator)
    elif generator_grid in grids_modify_values:
        try: generator_type = str(generator.pBMU.desc)
        except: generator_type = str(generator.desc)
        generator_grid_type = generator_grid + generator_type
        if generator_grid_type in market_grid_type_list:
            gen_other.append(generator)

for
    
    
    
def calcIzbrana(dfMarketData, df_select_nodes_p, df_select_nodes_info):
    global df_checking
    df_checking = pd.DataFrame()
    global df_izbrana_grid_type_sum
    global dfMarketSlo
    dfMarketSlo = dfMarketData.filter(regex='SI00')
    df_izbrana_grid_type_sum = pd.DataFrame()
    df_checking["Market SUM"] = dfMarketSlo["SI00_sum"]
    
    #suma izbranih voslisc po tipu energenta
    for izb_voz in df_select_nodes_p.columns:
        # df_select_nodes_info.drop(labels = 'Unnamed: 0', axis = 1, inplace = True)
        izb_grid_type = df_select_nodes_info.at[0,izb_voz]
        try:
            df_izbrana_grid_type_sum[izb_grid_type] = df_izbrana_grid_type_sum[izb_grid_type] + df_select_nodes_p[izb_voz]
        except:
            df_izbrana_grid_type_sum[izb_grid_type] = df_select_nodes_p[izb_voz]
            
    #Mam df sume izbranih vozlisc
    #Se kompletno
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
    
    for column in dfMarketData.columns.to_list():
        if column in delta:
            dfMarketData[column] = delta[column]
            app.PrintPlain(f"Subtracted {column} in delta from market")
    
    # replace_negatives = lambda x: 0 if x < 0 and x.name != 'SI00_29' else x
    # delta = delta.apply(replace_negatives)
    # delta = delta.applymap(lambda x: 0 if x < 0 and "28" not in x.name and "29" not in x.name and "44" not in x.name and "45" not in x.name else x)
    # dfMarketData = dfMarketData - df_izbrana_grid_type_sum
    
    # dfMarketData[] = df.apply(lambda row: row[[col for col in df.columns if col.startswith('C')]].sum(),
    return dfMarketData

def calcGenRatios(df_izbrana_vozlisca):
    if print_basic: app.PrintPlain("Racunanje sum generatorjev po drzavah in energentih")
    generator_list = []
    
    generator_ratios = {}
    generator_sum = {}
    generator_p_final = {}
    generator_grid_energent = {}
    
    for generator in mgenerators:
        generator_name = generator.loc_name
        if generator_name not in df_izbrana_vozlisca.columns:
            #V listi so generatorji ki jim bomo racunali razmerja
            generator_list.append(generator_name)
            generator_grid = generator.cpGrid.loc_name
            try:
                #V PF imajo generatorji desc/energent pod virtual powerplant na evropskem modelu
                generator_energent = str(''.join(generator.pBMU.desc))
            except:
                #V sloveniji je to zapisano direktno v desc, ne delamo virtualnih powerplantov
                generator_energent = str(''.join(generator.desc))
            #Naredimo po DRZAVA_ENERGENT tako kot je v market datoteki (npr. SI00_26 - to so hidroelektrarne ROR)
            generator_grid_energent[generator_name] = generator_grid + "_" + generator_energent
            #Za proizvodno enoto dobimo koliko je pmax
            generator_p = generator.pgini
            #Ce je pmax proizvodne enote enak 0 vzamemo pmin_ucpu, ce je tudi to 0 vzamemo default 2MW 
            # if generator_p == 0:
            #     generator_p = generator.P_max
            # if generator_p == 0:
            #     generator_p = abs(generator.Pmin_ucPU) 
            if generator_p == 0:
                generator_p = 2.0
            #Potem zapisemo se to v dictionary za potem 
            generator_p_final[generator_name] = generator_p
            #Racunamo sum
            generator_sum[generator_grid_energent[generator_name]] = 0
            
    for generator in generators:
        generator_name = generator.loc_name
        if generator_name in generator_list:
            generator_sum[generator_grid_energent[generator_name]] += generator_p_final[generator_name]   
            
    if print_basic: app.PrintPlain("Racunanje razmerja generatorjev") 
    
    for generator in generators:
        generator_name = generator.loc_name
        if generator_name in generator_list:
            generator_ratios[generator_name] = generator_p_final[generator_name] / generator_sum[generator_grid_energent[generator_name]]
                
    if print_all_info: 
        app.PrintPlain("RAZMERJE GENERATORJEV: ")
        app.PrintPlain(generator_ratios)
        
    return generator_ratios

def calculateGeneratorRatios(generators, grids, list_select_nodes, df_market_data):
    app.PrintPlain("Calculating grid GENERATOR sum")
    list_generator_name = []
    
    dict_generator_grid_type_sum = {}
    dict_generator_p_final = {}
    
    dict_generator_grid_type = {}
    dict_generator_ratios = {}
    #dict_generator_PQratio = {}
    
    temp_market_data_grid_type = df_market_data.columns
    
    # df_generator_data = pd.DataFrame()
    
    for generator in generators:
        generator_name = generator.loc_name
        generator_grid = generator.cpGrid.loc_name
        try:
            #V PF imajo generatorji desc/energent pod virtual powerplant na evropskem modelu
            generator_type = str(''.join(generator.pBMU.desc))
        except:
            #V sloveniji je to zapisano direktno v desc, ne delamo virtualnih powerplantov
            generator_type = str(''.join(generator.desc))
        generator_grid_type = generator_grid + "_" + generator_type
        if generator_grid in grids and generator_grid_type in temp_market_data_grid_type and generator_name not in list_select_nodes:
            #V listi so generatorji ki jim bomo racunali razmerja
            #Potem zapisemo se to v dictionary za potem 
            generator_p = generator.P_max
            dict_generator_p_final[generator_name] = generator_p
            #ce je gen z delovno mocjo razlicno od 0 nadaljujemo
            list_generator_name.append(generator_name)
            dict_generator_grid_type[generator_name] = generator_grid_type
            # Delamo sumo typov po drzavah
            if generator_grid_type in dict_generator_grid_type_sum:
                dict_generator_grid_type_sum[generator_grid_type] += generator_p
            else:
                dict_generator_grid_type_sum[generator_grid_type] = generator_p
    
    app.PrintInfo("Sum per generator grid and type: ")
    app.PrintInfo(dict_generator_grid_type_sum)
    
    #print("Sum bremen: " + str(generator_sum)) 
    app.PrintPlain("Calculating generator ratios")
    for generator in generators:
        generator_name = generator.loc_name
        if generator_name in list_generator_name:
            dict_generator_ratios[generator_name] = dict_generator_p_final[generator_name] / dict_generator_grid_type_sum[dict_generator_grid_type[generator_name]]
            #generator_q = generator.qgini
            # if generator_q != 0:
            #     dict_generator_PQratio[generator_name] = generator_q / dict_generator_p_final[generator_name]
            # else:
            #     dict_generator_PQratio[generator_name] = 0
            # df_generator_data.at[generator_name, "grid_type"] = dict_generator_grid_type[generator_name]
            # df_generator_data.at[generator_name, "grid_type_ratio"] = dict_generator_ratios[generator_name]
            # df_generator_data.at[generator_name, "pqratio"] = dict_generator_PQratio[generator_name]
    
    app.PrintInfo("generator radios: ")
    app.PrintInfo(dict_generator_ratios)
    
    return dict_generator_grid_type, dict_generator_ratios
    #return list_generator_name, dict_generator_ratios, dict_generator_grid_type, dict_generator_PQratio

def calculateLoadRatios(loads, grids, list_select_nodes, df_market_data):
    app.PrintPlain("Calculating grid LOAD sum")
    list_load_name = []
    
    dict_load_grid_sum = {}
    dict_load_p_final = {}
    
    dict_load_grid = {}
    dict_load_ratios = {}
    dict_load_PQratio = {}
    
    temp_market_data_grid_type = df_market_data.columns
    
    # df_load_data = pd.DataFrame()
    
    for load in loads:
        load_name = load.loc_name
        load_grid = load.cpGrid.loc_name
        load_grid_LOAD = load_grid + "_LOAD"
        if load_grid in grids and load_grid_LOAD in temp_market_data_grid_type and load_name not in list_select_nodes:
            #V listi so generatorji ki jim bomo racunali razmerja
            #Za load dobimo koliko je p - plini
            #Potem zapisemo se to v dictionary za potem 
            load_p = load.plini
            if load_p != 0.0:
                dict_load_p_final[load_name] = load_p
                #ce je breme z delovno mocjo razlicno od 0 nadaljujemo
                list_load_name.append(load_name)
                dict_load_grid[load_name] = load_grid_LOAD
                #Ce key obstaja v dictionary sestevamo, drugace zapisemo
                if load_grid_LOAD in dict_load_grid_sum:
                    dict_load_grid_sum[load_grid_LOAD] += load_p
                else:
                    dict_load_grid_sum[load_grid_LOAD] = load_p
    
    app.PrintInfo("Grid total load sum: ")
    app.PrintInfo(dict_load_grid_sum)
    
    #print("Sum bremen: " + str(load_sum)) 
    app.PrintPlain("Calculating load ratios")
    for load in loads:
        load_name = load.loc_name
        if load_name in list_load_name:
            dict_load_ratios[load_name] = dict_load_p_final[load_name] / dict_load_grid_sum[dict_load_grid[load_name]]
            load_q = load.qlini
            if load_q != 0:
                dict_load_PQratio[load_name] = load_q / dict_load_p_final[load_name] 
            else:
                dict_load_PQratio[load_name] = 0
            # df_load_data.at[load_name, "grid_type"] = dict_load_grid[load_name]
            # df_load_data.at[load_name, "grid_type_ratio"] = dict_load_ratios[load_name]
            # df_load_data.at[load_name, "pqratio"] = dict_load_PQratio[load_name]
    
    app.PrintInfo("Load radios: ")
    app.PrintInfo(dict_load_ratios)
    return dict_load_grid, dict_load_ratios, dict_load_PQratio
    # return list_load_name, dict_load_ratios, dict_load_grid, dict_load_PQratio

def setGeneratorAndLoadPQ(generators, loads, df_market_data, gengridtype, genrat, loadgrid, loadrat, loadPQrat, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, df_select_node_info, hour, grids_to_set):
    # Ta funkcija nastavi vrednosti moči generatorjem in bremenom po državah in tipih energentov
    # Najprej vzame vrednosti moči po državah in tipih energentov iz market datoteke, in temu odšteje vrednosti izbranih vozlišč
    # Izbrana vozlišča so vozlišča v slo (večje elektrarne in porabniki) za katere načeloma vemo moči po urah na letni ravni, prav tako remonte itd.
    # Od market datoteke odšteta izbrana vozlišča nato razporedi po vozliščh v powerfactory, po enakih razmerjih kot so bila na začetku
    # Jalove moči so določene preko cosfi, kjer lahko to določimo ročno v datoteki parametrov ali pa cosfi ostane enak kot je bil pred nastavitvijo novih moči
    app.PrintPlain("Setting generator and load power for current hour")
    app.PrintPlain("Odstevanje DUMP od market datoteke. Odstevamo od tipa wind in solar")
    #grid_type_list = df_market_data.columns
    
    #Tole je hitrej?
    temp_market_values_hour = df_market_data.loc[hour].to_dict()
    app.PrintInfo(temp_market_values_hour)
    
    ts = time.time()
    
    #for grid_type in grid_type_list:
    for grid_type in temp_market_values_hour:
        # Najprej pomnozimo z (1-default izgube) ker so v market datoteki zajete izgube omrežja v LOAD
        # Trenutno hardcoded 3% izgub
        #if "_LOAD" in grid_type: df_market_data.at[hour, grid_type] *= 0.97 
        if "_LOAD" in grid_type: temp_market_values_hour[grid_type] *= 0.97 
        # Dumped odstejemo enakomrno od sonca/vetra
        if "_Dump" in grid_type and temp_market_values_hour[grid_type] > 0:
            #Ce je dump > 0
            dump_grid = grid_type[0:4] #Vzamemo samo grid npr. SI00
            grid_solar = dump_grid + "_33"
            grid_wind_offshore = dump_grid + "_32"
            grid_wind_onshore = dump_grid + "_31"
            hour_power_dump = temp_market_values_hour[grid_type]
            hour_power_solar = temp_market_values_hour[grid_solar]
            hour_power_wind_offshore = temp_market_values_hour[grid_wind_offshore]
            hour_power_wind_onshore = temp_market_values_hour[grid_wind_onshore]
            power_sum = hour_power_solar + hour_power_wind_offshore + hour_power_wind_onshore
            if power_sum <= 0:
                app.PrintWarn("There is dumped energy and no solar/wind power")
            #Sonca/vetra je nad 0MW
            else:
                if hour_power_solar > 0:
                    temp_solar_before =  temp_market_values_hour[grid_solar]
                    ratio_grid_solar = hour_power_solar/power_sum
                    temp_market_values_hour[grid_solar] -= hour_power_dump * ratio_grid_solar
                    if temp_market_values_hour[grid_solar] < 0: temp_market_values_hour[grid_solar] = 0 #Cap to 0
                    
                if hour_power_wind_offshore > 0:
                    temp_windoff_before =  temp_market_values_hour[grid_wind_offshore]
                    ratio_grid_wind_offshore = hour_power_wind_offshore/power_sum
                    temp_market_values_hour[grid_wind_offshore] -=  hour_power_dump * ratio_grid_wind_offshore
                    if temp_market_values_hour[grid_wind_offshore] < 0: temp_market_values_hour[grid_wind_offshore] = 0 #Cap to 0
                    
                if hour_power_wind_onshore > 0:
                    temp_windon_before =  temp_market_values_hour[grid_wind_onshore]
                    ratio_grid_wind_onshore = hour_power_wind_onshore/power_sum
                    temp_market_values_hour[grid_wind_onshore] -= hour_power_dump * ratio_grid_wind_onshore
                    if temp_market_values_hour[grid_wind_onshore] < 0: temp_market_values_hour[grid_wind_onshore] = 0 #Cap to 0
                if False:
                    app.PrintInfo(dump_grid + " " + str(hour_power_dump))
                    app.PrintInfo("Solar before: " + str(temp_solar_before) + " solar after: " + str(temp_market_values_hour[grid_solar]))
                    app.PrintInfo("Wind_offsh before: " + str(temp_windoff_before) + " wind offsh after: " + str(temp_market_values_hour[grid_wind_offshore]))
                    app.PrintInfo("Wind_onsh before: " + str(temp_windon_before) + " wind onsh after: " + str(temp_market_values_hour[grid_wind_onshore]))
                    app.PrintInfo("Dump total: " + str(temp_solar_before + temp_windoff_before + temp_windon_before - temp_market_values_hour[grid_solar] - temp_market_values_hour[grid_wind_offshore] - temp_market_values_hour[grid_wind_onshore]))
                # app.PrintPlain("Solar before: " + str(df_antares.at[hour, grid_solar]))
                # app.PrintPlain("Dumped: " + str(df_antares.at[hour, grid_type]))
                # app.PrintPlain("Razmerje: " + str(ratio_grid_solar))
                # app.PrintPlain("Solar after: " + str(df_antares.at[hour, grid_solar]))
        # - energente odstejemo od + energentov (za baterije/crpalne/power2gas....)
        if "-" in grid_type:
            # app.PrintInfo("Odstevamo moc neg. energenta")
            grid_type_pos_temp = grid_type[0:5] + grid_type[6:]
            # app.PrintInfo("Negativni grid " + grid_type)
            # app.PrintInfo("Pozitivni grid " + grid_type_pos_temp)
            # app.PrintInfo("Moc pizitivnega prej:")
            # app.PrintInfo(df_market_data.at[hour, grid_type_pos_temp])
            temp_market_values_hour[grid_type_pos_temp] -= temp_market_values_hour[grid_type]
            # app.PrintInfo("Moc pozitivnega za grid " + grid_type_pos_temp + ": " + str(df_market_data.at[hour, grid_type_pos_temp]))
            
    # Racunamo sum izbranih vozlisc da ga odstejemo od market datoteke
    app.PrintPlain("Subtracting select node power sum from market data")
    
    temp_select_nodes_p = df_izbrana_vozlisca_p.loc[hour].to_dict()
    temp_select_nodes_q = df_izbrana_vozlisca_p.loc[hour].to_dict()
    temp_select_nodes_info = df_select_node_info.loc['type'].to_dict()
    
    # Summing known node power
    known_node_type_sum = {}
    list_known_nodes = df_select_node_info.columns
    for node in list_known_nodes:
        if temp_select_nodes_info[node] in known_node_type_sum:
            known_node_type_sum[temp_select_nodes_info[node]] += temp_select_nodes_p[node]
        else:
            known_node_type_sum[temp_select_nodes_info[node]] = temp_select_nodes_p[node]
    #subtracting known node values
    for grid_type in known_node_type_sum:
        if grid_type in df_market_data.columns:
            temp_market_value = temp_market_values_hour[grid_type]
            temp_market_values_hour[grid_type] -= known_node_type_sum[grid_type]
            #Ce je manjsi kot 0 ga nastavimo na 0 CAPPING JA ALI NE???
            if temp_market_values_hour[grid_type] < 0.0: 
                temp_market_values_hour[grid_type] = 0.0
                if False: app.PrintInfo("Value would be negative, capped at 0")
            if False: 
                if False: app.PrintInfo("Type: " + str(grid_type) + " before: " + str(temp_market_value) + " to subtract: " + str(known_node_type_sum[grid_type]) + " after: " + str(df_market_data.at[hour, grid_type]))
        
    # Za uro je zdej odsteto od market datoteke 
    
    default_cosfi_gen = 0.97
    default_PQrat_gen = math.tan(math.acos(default_cosfi_gen))
    mod_gen = 0
    mod_load = 0
    
    app.PrintPlain("Setting generator power")
    #Potem dolocamo moc generatorjev
    for generator in generators:
        generator_name = generator.loc_name
        # Generator from known nodes
        if generator_name in list_known_nodes:
            generator.pgini = float(temp_select_nodes_p[node])
            generator.qgini = float(temp_select_nodes_q[node])
            mod_gen += 1
        # Generator from grids we are setting.
        elif generator_name in gengridtype:
            #Dobimo decription generatorja, izven slo so virtualne elektrarne, v slo so v desc
            # grid_type_temp = df_generator_data.at[generator_name, "grid_type"]
            # app.PrintPlain("Generator grid type" + grid_type_temp)
            # grid_type_neg_temp = grid_type_temp[0:5] + "-" + grid_type_temp[5:]
            # app.PrintPlain(grid_type_neg_temp)
            # if grid_type_neg_temp in grid_type_list:
            #     # Ce mamo negativni energent odstejemo
            #generator.pgini = df_market_data.at[hour,gengridtype[generator_name]] * genrat[generator_name]
            generator.pgini = temp_market_values_hour[gengridtype[generator_name]] * genrat[generator_name]
            #generator.qgini = generator.pgini * df_generator_data.at[generator_name, "pqratio"]
            generator.qgini = generator.pgini * default_PQrat_gen
            mod_gen += 1
            # else:
            #     # Drugace samo racunamo, zaenkrat fiksen cosfi
            #     generator.pgini = df_market_data.at[hour,df_market_data.at[hour,grid_type_temp]] * df_generator_data.at[generator_name, "grid_type_ratio"]
            #     generator.qgini = generator.pgini * df_generator_data.at[generator_name, "pqratio"]
        else:
            if False: app.PrintInfo("Generator " + generator_name + " not set.")
            
    app.PrintInfo("Setting load power")
    #Dolocimo mod bremen
    for load in loads:
        load_name = load.loc_name
        #app.PrintPlain(load_name)
        if load_name in list_known_nodes:
            if False: app.PrintPlain(df_izbrana_vozlisca_p.at[hour, load_name])
            load.plini = float(df_izbrana_vozlisca_p.at[hour, load_name])
            load.qlini = float(df_izbrana_vozlisca_q.at[hour, load_name])
            mod_load += 1
        elif load_name in loadgrid:
            #temp_load_grid_type = loadgrid[load_name]
            #temp_market_p = df_market_data.at[hour, temp_load_grid_type]
            load.plini = temp_market_values_hour[loadgrid[load_name]] * loadrat[load_name]
            #temp_pqratio = loadrat[load_name]
            # app.PrintInfo(temp_load_grid_type)
            # app.PrintInfo(temp_market_p)
            # app.PrintInfo(temp_pqratio)
            
            #load.plini = temp_market_p * temp_pqratio
            load.qlini = load.plini * loadPQrat[load_name]
            mod_load += 1
        else:
            if False: app.PrintInfo("Load " + load_name + " not set")
            
    app.PrintPlain("Time to set: " + str(time.time() - ts))
    
    app.PrintInfo("Modified gen nr: " + str(mod_gen))
    app.PrintInfo("Modified loads nr: " + str(mod_load))
    return

def voltagesourceinfo(voltagesources, grids, df_border_nodes_info):
    for voltagesource in voltagesources:
        voltagesource_name = voltagesource.loc_name
        voltagesource_grid = voltagesource.cpGrid.loc_name
        if voltagesource_name in df_border_nodes_info.index and voltagesource_grid in grids:
            if abs(voltagesource.Pgen) > 0:
                df_border_nodes_info.at[voltagesource_name, "pqratio"] = voltagesource.Qgen/voltagesource.Pgen
            else:
                df_border_nodes_info.at[voltagesource_name, "pqratio"] = 0
    return df_border_nodes_info

def setCrossborderExchanges(voltagesources, df_crossborder_exchanges, df_border_nodes_info, hour, grids_to_set):
    # Funkcija nastavlja moč robnih vozlisc - to so vozlisca na robovih eu (zunanje povezave zunaj ENTSO-E) 
    # ter DC daljnovodi katerim lahko nastavljamo smer in jakost pretoka. Modelirani so kot voltagesource
    #Vhod so "ELMVAC", dataframe crossborder iz antares datoteke, dataframe robnih vozlisc iz excel datoteke in ura simulacije
    app.PrintPlain("Setting external borders")
    for voltagesource in voltagesources:
        #Najprej dobimo ime vozlisca (v PF modelirani kot VAC zato je ime voltagesource)
        voltagesource_name = voltagesource.loc_name
        voltagesource_grid = voltagesource.cpGrid.loc_name
        if voltagesource_name in df_border_nodes_info.index and voltagesource_grid in grids_to_set:
            #app.PrintPlain("External border: " + voltagesource_name + " grid " + voltagesource_grid)
            #Najdemo ime - mejo drzav ki mora bit eneka kot v antares
            border = df_border_nodes_info.at[voltagesource_name, "border"]
            #print("Meja: " + border)
            #Preverimo ce se border nahaja v antaresu
            if border in df_crossborder_exchanges.columns:
                #app.PrintPlain("External border before: " + voltagesource_name + " set P: " + str(voltagesource.Pgen) + " set Q " + str(voltagesource.Qgen))
                voltagesource.Pgen = float(df_crossborder_exchanges.at[hour, border]) * df_border_nodes_info.at[voltagesource_name, "ratio"] * df_border_nodes_info.at[voltagesource_name, "multiplier"]
                voltagesource.Qgen = voltagesource.Pgen * df_border_nodes_info.at[voltagesource_name, "pqratio"]
                #app.PrintPlain("External border before: " + voltagesource_name + " set P: " + str(voltagesource.Pgen) + " set Q " + str(voltagesource.Qgen))
                # print("Moc v antares: " + str(totalpower))
                # print("Ratio: " + str(ratio))
                # print("Multiplier: " + str(multiplier))
                # app.PrintPlain("Koncna moc: " + str(voltagesource_p))   
    return

def calcAndSetGenLoadPower(generators, loads, df_antares, generator_ratios, load_ratios, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour, grids_to_set):
    
    # Ta funkcija nastavi vrednosti moči generatorjem in bremenom po državah in tipih energentov
    # Najprej vzame vrednosti moči po državah in tipih energentov iz market datoteke, in temu odšteje vrednosti izbranih vozlišč
    # Izbrana vozlišča so vozlišča v slo (večje elektrarne in porabniki) za katere načeloma vemo moči po urah na letni ravni, prav tako remonte itd.
    # Od market datoteke odšteta izbrana vozlišča nato razporedi po vozliščh v powerfactory, po enakih razmerjih kot so bila na začetku
    # Jalove moči so določene preko cosfi, kjer lahko to določimo ročno v datoteki parametrov ali pa cosfi ostane enak kot je bil pred nastavitvijo novih moči
    if print_basic: app.PrintPlain("Nastavljanje moci generatorjev ter bremen v trenutnem modelu in scenariju")
    if print_basic: app.PrintPlain("Odstevanje DUMP od market datoteke. Odstevamo od tipa wind in solar")
    
    PQrat_old = {}
    current_hour_sum = {}
    #hour = 8756
    for grid_type in df_antares.columns:
        # Najprej pomnozimo z (1-default izgube) ker so v market datoteki zajete izgube omrežja v LOAD
        if "_LOAD" in grid_type: df_antares.at[hour, grid_type] *= 1 - default_izgube
        # Dumped odstejemo enakomrno od sonca/vetra
        if "_Dump" in grid_type and df_antares.at[hour, grid_type] > 0:
            if print_all_info: app.PrintPlain("Imamo dump in je vecji od 0")
            #Ce je dump > 0
            dump_grid = grid_type[0:4] #Vzamemo samo grid npr. SI00
            grid_solar = dump_grid + "_33"
            grid_wind_offshore = dump_grid + "_32"
            grid_wind_onshore = dump_grid + "_31"
            hour_power_dump = df_antares.at[hour, grid_type]
            hour_power_solar = df_antares.at[hour, grid_solar]
            hour_power_wind_offshore = df_antares.at[hour, grid_wind_offshore]
            hour_power_wind_onshore = df_antares.at[hour, grid_wind_onshore]
            power_sum = hour_power_solar + hour_power_wind_offshore + hour_power_wind_onshore
            if power_sum >= 0: power_sum = 1
            if hour_power_solar > 0:
                ratio_grid_solar = hour_power_solar/power_sum
            else:
                ratio_grid_solar = 0
            if hour_power_wind_offshore > 0:
                ratio_grid_wind_offshore = hour_power_wind_offshore/power_sum
            else:
                ratio_grid_wind_offshore = 0
                
            if hour_power_wind_onshore > 0:
                ratio_grid_wind_onshore = hour_power_wind_onshore/power_sum
            else:
                ratio_grid_wind_onshore = 0
                
            if print_all_info: 
                app.PrintPlain("Solar prej: " + str(df_antares.at[hour, grid_solar]))
                app.PrintPlain("Dumped: " + str(df_antares.at[hour, grid_type]))
                app.PrintPlain("Razmerje: " + str(ratio_grid_solar))
            
            df_antares.at[hour, grid_solar] -= hour_power_dump * ratio_grid_solar
            if df_antares.at[hour, grid_solar] < 0: df_antares.at[hour, grid_solar]  = 0 #Cap to 0
                
            df_antares.at[hour, grid_wind_offshore] -=  hour_power_dump * ratio_grid_wind_offshore
            if df_antares.at[hour, grid_wind_offshore] < 0: df_antares.at[hour, grid_wind_offshore]  = 0 #Cap to 0
                
            df_antares.at[hour, grid_wind_onshore] -= hour_power_dump * ratio_grid_wind_onshore
            if df_antares.at[hour, grid_wind_onshore] < 0: df_antares.at[hour, grid_wind_onshore]  = 0 #Cap to 0
                
            if print_all_info: app.PrintPlain("Solar potem: " + str(df_antares.at[hour, grid_solar]))
    
    #app.PrintPlain("IZRACUNAN LOAD SUM KI GA BI NASTAVU: " + str(df_antares.at[hour, 'SI00_LOAD']))
    
    # Racunamo sum izbranih vozlisc da ga odstejemo od market datoteke
    if print_basic: app.PrintPlain("Racunanje sume izbranih vozlisc in odstevanje od market datoteke")
    # Najprej izdelamo vektor grid_energent kamor bomo sesteval. Torej npr. SI00_LOAD, SI00_26, SI00_30, ...V te najprej napisemo vrednosti 0.0
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] = 0.0   
    #Sestejemo vrednosti v dataframu izbranih vozlisc - sum po grid_tip.
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] += float(df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[hour], vozlisce])
        #app.PrintPlain(current_hour_sum)
    #Potem odstevamo iz market datoteke
    for grid_type in current_hour_sum:
        if True: 
            app.PrintPlain("Grid in energent: " + str(grid_type))
            app.PrintPlain("Vrednost v market datoteki prej: " + str(df_antares.at[hour, grid_type]))
            app.PrintPlain("Vrednost za odstet: " + str(current_hour_sum[grid_type]))
        df_antares.at[hour, grid_type] -= current_hour_sum[grid_type]
        #Ce je manjsi kot 0 ga nastavimo na 0
        if df_antares.at[hour, grid_type] < 0.0 : 
            df_antares.at[hour, grid_type] = 0.0
        if True: app.PrintPlain("Vrednost v market datoteki potem: " + str(df_antares.at[hour, grid_type]))
    # Za uro je zdej odsteto od market datoteke 
    
    if print_basic: app.PrintPlain("NASTAVLJANJE MOCI GENERATORJEV")
    #Potem dolocamo moc generatorjev
    for generator in generators:
        generator_name = generator.loc_name
        generator_grid = generator.cpGrid.loc_name
        #Dobimo decription generatorja, izven slo so virtualne elektrarne, v slo so v desc
        if generator_grid in grids_to_set:
            try:
                generator_type = str(''.join(generator.pBMU.desc))
            except:
                generator_type = str(''.join(generator.desc)) 
            generator_grid_type = generator_grid + "_" + generator_type
        
            #Preverimo ce je generator v izbranih vozliscih in dolocimo tu
            if generator_name in df_izbrana_vozlisca_p.columns:
                if print_all_info: app.PrintPlain("Generator iz izbranega vozlisca")
                #Pri importu so duplikatom v imenih elektrarn (kot npr pri crpalni avce kjer je lahko generator ali breme) dodani .1 na koncu zato preverimo če obstaja,
                # torej avce crpanje so CHEAVC-GEN in avce proizvodnja CHEAVC-GEN.1. To je zacasno hardcoded
                # Naredimo se tako da ce crpa v omrezje oddaja jalovo moc, ce pa deluje kot generator iz omrezja jemlje jalovo (to se dela za uravnavanje napetosti)
                generator_name_type_neg = generator_name + ".1" #To je proizvodnja
                if generator_name_type_neg in df_izbrana_vozlisca_p.columns:
                    #Preverimo ali je večja cifra za proizvodnjo ali porabo in določimo moči na podlagi tega (napjprej za P, potem Q)
                    if float(df_izbrana_vozlisca_p.at[hour, generator_name]) > float(df_izbrana_vozlisca_p.at[hour, generator_name_type_neg]):
                        #Tukaj crpa in jemlje delovno
                        generator.pgini = - float(df_izbrana_vozlisca_p.at[hour, generator_name])
                    else:
                        #Tukaj generira delovno
                        generator.pgini = float(df_izbrana_vozlisca_p.at[hour, generator_name_type_neg])
                    
                    if float(df_izbrana_vozlisca_q.at[hour, generator_name]) > float(df_izbrana_vozlisca_q.at[hour, generator_name_type_neg]):
                        #Tu jemlje jalovo
                        generator.qgini = - float(df_izbrana_vozlisca_q.at[hour, generator_name])
                    else:
                        #Tu oddaja jalovo
                        generator.qgini = float(df_izbrana_vozlisca_q.at[hour, generator_name_type_neg])
                #Ce ni crpalna, baterija itd. in je samo generator nastavimo direktno
                else:
                    generator.pgini = float(df_izbrana_vozlisca_p.at[hour, generator_name])
                    generator.qgini = float(df_izbrana_vozlisca_q.at[hour, generator_name])
                
            #Tu je generator iz market datoteke       
            elif generator_grid_type in df_antares.columns:
                #Negativni energent za potrebe baterij in akumulacijskih in power to gas
                #To deluje tako da vrine - pred cifto
                generator_grid_type_neg = generator_grid_type[0:5] + "-" + generator_grid_type[5:]
                if print_all_info: app.PrintPlain("Generator : " + generator_name + " z grid in energentom: " + generator_grid_type)
                #Izracunamo se tanfi (lazji racun kot cosfi ampak ista zadeva), kar je razmerje med delovno in jalovo
                if generator.pgini == 0.0 or generator.qgini == 0.0:
                    PQrat_old[generator_name] = 0.0
                else:
                    PQrat_old[generator_name] = generator.qgini/generator.pgini
                    
                if generator_grid_type_neg in df_antares.columns and df_antares.at[hour, generator_grid_type_neg] > df_antares.at[hour, generator_grid_type]:
                    #Obstaja negativni tip energenta v antares
                    #Negativni tip ima večjo moč
                    generator_p = - float(df_antares.at[hour, generator_grid_type_neg] * generator_ratios[generator_name])
                    #Ali jalovo proizvaja ali porablja??
                    if oldcosfi:
                        generator_q = - float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * PQrat_old[generator_name])
                    else:
                        generator_q = - float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * default_PQrat_gen)
                else:
                    #Ni negativnega energenta ali ima manjšo moč
                    generator_p = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name])
                    if oldcosfi:
                        generator_q = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * PQrat_old[generator_name])
                    else:
                        generator_q = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * default_PQrat_gen)
                if print_all_info: app.PrintPlain(generator_p)
                if print_all_info: app.PrintPlain(generator_q)
                generator.pgini = generator_p
                generator.qgini = generator_q
                #app.PrintPlain("Generator: " + generator_name + " z grid in energentom: " + generator_grid_type + " ima dodeljeno: " + str(generator_p))
            else:
                if print_all_info: app.PrintPlain("Generator: " + generator_name + " z grid in energentom: " + generator_grid_type + " ni me dodeljeno nič")
        else:
            if print_all_info: app.PrintPlain("Generator je iz grida v katerem ne nastavljamo: " + generator_grid)
    
    if print_basic: app.PrintPlain("NASTAVLJANJE MOCI BREMEN")
    #Se za bremena
    for load in loads:
        load_name = load.loc_name
        load_grid = load.cpGrid.loc_name
        if load_grid in grids_to_set:
            load_grid_type = load_grid + "_LOAD"
            if load_name in df_izbrana_vozlisca_p.columns:
                #Breme iz izbranih vozlisc
                load.plini = float(df_izbrana_vozlisca_p.at[hour, load_name])
                load.qlini = float(df_izbrana_vozlisca_q.at[hour, load_name])
                if print_all_info: app.PrintPlain("Load iz izbranega vozlisca: " + load_name + " nastavljen P: " + str(load.plini) + " in Q: " + str(load.qlini))
                
            elif load_grid_type in df_antares.columns:
                if load.plini == 0.0 or load.qlini == 0.0:
                    PQrat_old[load_name] = 0.0
                else:
                    PQrat_old[load_name] = load.qlini/load.plini
                load.plini = load_ratios[load_name] * df_antares.at[hour, load_grid_type]
                if oldcosfi:
                    #ce je bool true vzamemo cosfi enak kot je bil prej v modelu
                    load.qlini = load.plini * PQrat_old[load_name]
                else:
                    #Drugace nastavimo jalovo po rocno definiranem cosfi
                    load.qlini = load.plini * default_PQrat_load
            else:
                if print_all_info: app.PrintPlain("Breme: " + load_name + " ima drzavo: " + load_grid + " zato je preskocen")
        else:
            if print_all_info: app.PrintPlain("Load je iz grida v katerem ne nastavljamo: " + generator_grid)
    return

def nastaviRobnavozlisca(voltagesources, df_antares_crossborder, df_robna, hour, grids_to_set):
    # Funkcija nastavlja moč robnih vozlisc - to so vozlisca na robovih eu (zunanje povezave zunaj ENTSO-E) 
    # ter DC daljnovodi katerim lahko nastavljamo smer in jakost pretoka. Modelirani so kot voltagesource
    #Vhod so "ELMVAC", dataframe crossborder iz antares datoteke, dataframe robnih vozlisc iz excel datoteke in ura simulacije
    if print_basic: app.PrintPlain("NASTAVLJANJE ROBNIH VOZLISC")
    for voltagesource in voltagesources:
        #Najprej dobimo ime vozlisca (v PF modelirani kot VAC zato je ime voltagesource)
        voltagesource_name = voltagesource.loc_name
        voltagesource_grid = voltagesource.cpGrid.loc_name
        #print("Vozlisce: " + voltagesource_name)
        #ZACASNO HARDCODED ZA TESTIRANJE
        #Ce je ta v excel datoteki robnih vozlisc gremo dalje
        if voltagesource_name in df_robna.index and voltagesource_grid in grids_to_set:
            if print_all_info: app.PrintPlain("Najdeno robno vozlisce: " + voltagesource_name)
            #Najdemo ime - mejo drzav ki mora bit eneka kot v antares
            border = df_robna.at[voltagesource_name, 'MEJA']
            #print("Meja: " + border)
            #Preverimo ce se border nahaja v antaresu
            if border in df_antares_crossborder.columns:
                #totalpower = df_antares_crossborder.at[df_antares_crossborder.index[hour], border]
                totalpower = df_antares_crossborder.at[hour, border]
                ratio = df_robna.at[voltagesource_name, 'DELEZ']
                multiplier = df_robna.at[voltagesource_name, 'POMNOZITI']
                voltagesource_p = totalpower * ratio * multiplier
                #Zakaj točno nastavljamo na int???
                voltagesource.Pgen = voltagesource_p
                voltagesource.Qgen = default_PQrat_vac * voltagesource_p
                if print_all_info: app.PrintPlain("Vozlisce: " + voltagesource_name + " dodeljena moc: " + str(voltagesource_p))
                # print("Moc v antares: " + str(totalpower))
                # print("Ratio: " + str(ratio))
                # print("Multiplier: " + str(multiplier))
                # app.PrintPlain("Koncna moc: " + str(voltagesource_p))   
    return
            
def clearFolders():
    app.PrintPlain("Clearing partial results data")
    # Find folders in partial results folder
    partial_results_fodler = os.getcwd() + r'/Vmesni rezultati'
    #file_list = list()
    for root, dirs, files in os.walk(partial_results_fodler, topdown=False):
        for file in files:
            file_path = os.path.join(root, file)
            os.remove(file_path)
            #file_list.append(file_path)
            if True: app.PrintInfo("Removed file: " + file_path)
    return

#ZAPIS STATUSA IZRACUNA (konvergiralo/ni konvergiralo), mogoče še kaj drugega
def saveCalcStatus(hour, status, t_calc):
    df_results_calcstatus = pd.DataFrame()
    #Shranimo status: 0=OK, 1=Divergenca notranjih zank, 2=Divergenca zunanjih zank
    df_results_calcstatus.at[hour,'convergence'] = int(status)
    df_results_calcstatus.at[hour,'calculation_time'] = int(t_calc)
    loading_file_path = os.getcwd()  + r'/Vmesni rezultati/Calculation/Calcstatus_hour_' + str(hour) +'.csv'
    df_results_calcstatus.to_csv(loading_file_path, encoding='utf-8', index=True)
    return

#PISANJE REZULTATOV V EXCEL
def shraniVmesneRezultateCsv(generators, loads, lines, transformers, terminals, grids_to_write_results, hour):
    if print_basic: app.PrintPlain("Writing element results for hour: " + str(hour))
    
    df_results_line_loading_hourly = pd.DataFrame(data=None)
    df_results_transformer_loading_hourly = pd.DataFrame(data=None)
    df_results_voltage_hourly = pd.DataFrame(data=None)
    df_results_generator_P_set = pd.DataFrame(data=None)
    df_results_generator_Q_set = pd.DataFrame(data=None)
    df_results_load_P_set = pd.DataFrame(data=None)
    df_results_load_Q_set = pd.DataFrame(data=None)
    
    if print_basic: app.PrintPlain("Zapis rezultatov moci generatorjev")
    for generator in generators:
        generator_name = generator.loc_name
        generator_grid = generator.cpGrid.loc_name
        if generator_grid in grids_to_write_results and generator.IsOutOfService() == 0:
            df_results_generator_P_set.at[generator_name, hour] = generator.pgini
            df_results_generator_Q_set.at[generator_name, hour] = generator.qgini
            
    
    if print_basic: app.PrintPlain("Zapis rezultatov moci bremen")
    for load in loads:
        load_name = load.loc_name
        load_grid = load.cpGrid.loc_name
        if load_grid in grids_to_write_results and load.IsOutOfService() == 0:
            df_results_load_P_set.at[load_name, hour] = load.plini
            df_results_load_Q_set.at[load_name, hour] = load.qlini
    
    #Najprej zapisemo rezultate v dataframe za daljnovode
    if print_basic: app.PrintPlain("Zapis rezultatov obremenitev daljnovodov")
    for line in lines:
        line_name = line.loc_name
        line_grid = line.cpGrid.loc_name
        if line_grid in grids_to_write_results and line.IsOutOfService() == 0:
            if line.HasResults() == 1:
                #Ce je in service
                line_loading = round(line.GetAttribute('c:loading'), 2)
            else:
                # Če je element izklopljen
                line_loading = int(0)
            df_results_line_loading_hourly.at[line_name, hour] = line_loading
            
    #Nato se rezultate za transformatorje
    if print_basic: app.PrintPlain("Zapis rezultatov obremenitev transformatorjev")
    for transformer in transformers:
        transformer_name = transformer.loc_name
        transformer_grid = transformer.cpGrid.loc_name
        # hardcoded je da odstanimo trafote kjer je v imenu "/" ker so to 110/xx kV trafoti ponavadi od generatorjev
        if transformer_grid in grids_to_write_results and "/" not in transformer_name and "GT" not in transformer_name and "TES" not in transformer_name and transformer.IsOutOfService() == 0:
            if transformer.HasResults() == 1:
                #Ce je in service
                transformer_loading = round(transformer.GetAttribute('c:loading'), 2)
            else:
                # Če je element izklopljen
                transformer_loading = int(0)
            df_results_transformer_loading_hourly.at[transformer_name, hour] = transformer_loading
            
    if print_basic: app.PrintPlain("Zapis rezultatov napetosti zbiralk")
    #Modelirane so tudi nizjenapetostne zbiralke zato izpisemo samo te z napetnstnega nivoja 110, 220. 400 kV. Ven damo tut "odcepe" ki majo v imenu -
    terminal_voltages_write_out = [110,220,400]
    for terminal in terminals:
        terminal_name = terminal.loc_name
        terminal_grid = terminal.cpGrid.loc_name
        terminal_nominal_voltage = terminal.uknom
        if terminal_grid in grids_to_write_results and "-" not in terminal_name and terminal_nominal_voltage in terminal_voltages_write_out and terminal.IsOutOfService() == 0:
            if terminal.HasResults() == 1:
                #Ce je in services
                #Napetost je per-unit, nazivno lahko nardimo potem v koncnem izpisu?
                #Zaokrozimo na 3 decimalke
                terminal_voltage_pu = round(terminal.GetAttribute('m:u'), 3)
                #terminal_voltage_rated = terminal.uknom
                #terminal_voltage = round((terminal_voltage_rated * terminal_voltage_pu), 2)
            else:
                # Če je element izklopljen
                terminal_voltage_pu = int(0)
            df_results_voltage_hourly.at[terminal_name, hour] = terminal_voltage_pu
    
            
    if print_basic: app.PrintPlain("Zacetek zapisal excel datotek " + str(hour) + ". ure")
    #app.PrintPlain(df_results_lines)
    line_loading_file_path = os.getcwd()  + r'/Vmesni rezultati/Lines/Loading_ura_' + str(hour) +'.csv'
    transformer_loading_file_path = os.getcwd()  + r'/Vmesni rezultati/Transformers/Loading_ura_' + str(hour) +'.csv'
    voltage_file_path = os.getcwd()  + r'/Vmesni rezultati/Terminals/Voltage_ura_' + str(hour) +'.csv'
    generator_P_set_file_path = os.getcwd()  + r'/Vmesni rezultati/Generators/P_set_hour_' + str(hour) +'.csv'
    generator_Q_set_file_path = os.getcwd()  + r'/Vmesni rezultati/Generators/Q_set_hour_' + str(hour) +'.csv'
    load_P_set_file_path = os.getcwd()  + r'/Vmesni rezultati/Loads/P_set_hour_' + str(hour) +'.csv'
    load_Q_set_file_path = os.getcwd()  + r'/Vmesni rezultati/Loads/Q_set_hour_' + str(hour) +'.csv'
    
    df_results_line_loading_hourly.to_csv(line_loading_file_path, encoding='utf-8', index=True)
    df_results_transformer_loading_hourly.to_csv(transformer_loading_file_path, encoding='utf-8', index=True)
    df_results_voltage_hourly.to_csv(voltage_file_path, encoding='utf-8', index=True)
    df_results_generator_P_set.to_csv(generator_P_set_file_path, encoding='utf-8', index=True)
    df_results_generator_Q_set.to_csv(generator_Q_set_file_path, encoding='utf-8', index=True)
    df_results_load_P_set.to_csv(load_P_set_file_path, encoding='utf-8', index=True)
    df_results_load_Q_set.to_csv(load_Q_set_file_path, encoding='utf-8', index=True)
    
    if print_basic: app.PrintPlain("Izpisani in shranjeni rezultati " + str(hour) + ". ure")
    
    return

################################################################ GLAVNI ALGORITEM PROGRAMA ###############################################################

def testset():
    generators = app.GetCalcRelevantObjects("*.ElmSym")
    loads = app.GetCalcRelevantObjects("*.ElmLod")
    voltagesources = app.GetCalcRelevantObjects("*.ElmVac")
    lines = app.GetCalcRelevantObjects("*.ElmLne")
    transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    terminals = app.GetCalcRelevantObjects("*.ElmTerm")
    #hour = 180
    grids_modify_values, grids_write_results, hours_to_calculate = importParametersInPython()
    df_market_data, df_crossborder_exchanges, df_border_nodes_info, df_select_nodes_p, df_select_nodes_q, df_select_nodes_info = importData()
    
    gengridtype, genrat = calculateGeneratorRatios(generators, grids_modify_values, df_select_nodes_info.columns, df_market_data)
    loadgrid, loadrat, loadPQrat = calculateLoadRatios(loads, grids_modify_values, df_select_nodes_info.columns, df_market_data)
    df_border_nodes_info = voltagesourceinfo(voltagesources, grids_modify_values, df_border_nodes_info)
    
    ldf_and_results = False
    
    if ldf_and_results: 
        saveElementData(generators, loads, lines, transformers, terminals, grids_write_results)
        clearFolders()
    
    singlehour = True
    if singlehour:
        hour = 1
        t_start = time.time()
        setGeneratorAndLoadPQ(generators, loads, df_market_data, gengridtype, genrat, loadgrid, loadrat, loadPQrat, df_select_nodes_p, df_select_nodes_q, df_select_nodes_info, hour, grids_modify_values)
        setCrossborderExchanges(voltagesources, df_crossborder_exchanges, df_border_nodes_info, hour, grids_modify_values)
        status = ldf.Execute()
        if status == 0: gridsInterchange()
        t_calc = time.time() - t_start

    else:
        for hour in hours_to_calculate:
            t_start = time.time()
            setGeneratorAndLoadPQ(generators, loads, df_market_data, gengridtype, genrat, loadgrid, loadrat, loadPQrat, df_select_nodes_p, df_select_nodes_q, df_select_nodes_info, hour, grids_modify_values)
            setCrossborderExchanges(voltagesources, df_crossborder_exchanges, df_border_nodes_info, hour, grids_modify_values)
            status = ldf.Execute()
            t_calc = time.time() - t_start
            saveCalcStatus(hour, status, t_calc)
            app.PrintPlain("Loadflowstatus: " + str(status))
            app.PrintPlain("Potrebni cas: " + str(t_calc))
            app.PrintPlain("izveden ldf, izpis vmesnih rezultatov")
            if status == 0: 
                app.PrintPlain("izveden ldf, izpis vmesnih rezultatov")
                shraniVmesneRezultateCsv(generators, loads, lines, transformers, terminals, grids_write_results, hour)
            else: 
                app.PrintPlain("Ura: " + str(hour) + " brez konvergence")
            
    #calculateSumCompareToMarket(generators, loads, hour, df_market_data)
#################################################################################### MAIN #######################################################

#Samo izravun in zapis
#onlycalc()

testset()

#Izračun in združevanje rezultatov
# fullcalcandresults()

#Samo združevanje rezultatov
#combineresults()

#setSingleHour()

#izpisiSoncneSlo()
#################################################################################### MAIN #######################################################

#################### IZPIS URE #################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
