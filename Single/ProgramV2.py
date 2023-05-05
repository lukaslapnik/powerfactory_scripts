# -*- coding: utf-8 -*-
"""
Created on Sat May 28 08:49:17 2022

@author: SSIMON
"""

import pandas as pd
import datetime
import sys
#import numpy
import os
import time
from tkinter import Tk
from tkinter import filedialog
import powerfactory as pf    
app = pf.GetApplication()
app.ClearOutputWindow()
ldf = app.GetFromStudyCase("ComLdf")

##########################################################################################################################################################################
#############################################################################   PARAMETRI   ##############################################################################
##########################################################################################################################################################################
# Parametri za izračun jalovih moči za gen, load, vac..... načeloma če delamo DC loadflow ni važno
# Za AC loadflow je treba porihtat oz najt neke boljše načine dodeljevanja jalovih.
spreminjaj_jalovo = False  # Ali skripta sploh spreminja parametre proizvodnje/porabe jalove moči. False - jalova enaka, True - jalovo spreminja
izhodiscni_cosfi = True     # Ce je true, bo cosfi enak kot v izhodiscnem modelu, sicer vzame vrednosti definirane spodaj (razmerje med Q in P)
contingency_report = True

#Namesto cosfi se vnese razmerje PQ_ratio = tan(acos(cosfi(0.xx)))
#Pri cosfi 0.98 ~ 0.2
#Pri cosfi 0.97 ~ 0.25
#Pri cosfi 0.96 ~ 0.3
generator_PQ_ratio = 0.25 #Delez jalove
load_PQ_ratio = 0.25
voltagesource_PQ_ratio = 0

#Izkoristek omrezja (izgube)
izkoristek_omrezja = 0.97 #(1-izgube)

#Ure za katere skripta naredi izracune. Definiraj zacetno uro, koncno uro in inkrement/korak
zacetna_ura = 1
koncna_ura = 70
inkrement = 3

#Imena uvoženih datotek, glej da se sklada z tistim kar nardi skripta za pretvorbo excel->csv
stringMarketDataFile = "Market Data.csv"
stringBorderFlowFile = "Robna vozlisca P.csv"
stringBorderInfoFile = "Robna vozlisca Info.csv"
stringIzbranaPFile = "Izbrana vozlisca P.csv"
stringIzbranaQFile = "Izbrana vozlisca Q.csv"
stringIzbranaInfoFile = "Izbrana vozlisca Info.csv"

#   ['UKNI','UK00','UA02','UA01','TR00','TN00','SK00','SI00','SE04','SE03','SE02','SE01','SA00','RU00',
#   'RS00','RO00','PT00','PS00','PL00','NSW0','NOS0','NON1','NOM1','NL00','MT00','MK00','ME00','MD00',
#   'MA00','LY00','LV00','LUV1','LUG1','LUF1','LUB1','LT00','ITSI','ITSA','ITS1','ITN1','ITCS','ITCN',
#   'ITCA','IS00','IL00','IE00','HU00','HR00','GR03','GR00','FR15','FR00','FI00','ES00','ELES Interconnectios',
#   'EG00','EE00','DZ00','DKW1','DKKF','DEKF','DE00','CZ00','CY00','CH00','BG00','BE00','BA00','AT00','AL00']

#Drzave/sistemi, ki jim spreminjamo parametre. Vnesi tako kot je v market datoteki ali v powerfactory modelu
sistemi_spreminjanje_parametrov = ['SI00','ITN1','HU00','HR00','ELES Interconnectios']

#Drzave*sistemki, za katere se izpisujejo rezultati
sistemi_izpis_rezultatov = ['SI00', 'ELES Interconnectios']

##########################################################################################################################################################################
#############################################################################   PARAMETRI   ##############################################################################
##########################################################################################################################################################################

start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")

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
f_input_data_directory = filedialog.askdirectory()
app.PrintPlain("Input data folder selected!")
app.PrintPlain("Select output data folder!")
f_output_data_directory = filedialog.askdirectory()
app.PrintPlain("Output data folder selected!")
#Beri datoteke
app.PrintPlain("Importing market file")
dfMD = pd.read_csv(os.path.join(f_input_data_directory, stringMarketDataFile), index_col = [0])
app.PrintPlain("Importing border flow files")
dfCbFlow = pd.read_csv(os.path.join(f_input_data_directory, stringBorderFlowFile), index_col = [0])
dfCbInfo = pd.read_csv(os.path.join(f_input_data_directory, stringBorderInfoFile), index_col = [0])
app.PrintPlain("Importing izbrana vozlisca")
dfIzbP = pd.read_csv(os.path.join(f_input_data_directory, stringIzbranaPFile), header = [0], index_col = [0])
dfIzbQ = pd.read_csv(os.path.join(f_input_data_directory, stringIzbranaQFile), header = [0], index_col = [0])
dfIzbInfo = pd.read_csv(os.path.join(f_input_data_directory, stringIzbranaInfoFile), header = [0])
app.PrintPlain("Files imported")
# df_select_nodes_info = pd.read_csv(input_path_select_nodes_info, index_col = [1], header = [0])
# return dfMD, dfCbFlow, dfCbInfo, dfIzbP, dfIzbQ, dfIzbInfo 

# if clear_output_folder:
#     app.PrintPlain("Clearing partial results data")
#     # Find folders in partial results folder
#     #file_list = list()
#     for root, dirs, files in os.walk(f_output_data_directory, topdown = False):
#         for file in files:
#             file_path = os.path.join(root, file)
#             os.remove(file_path)
#             #file_list.append(file_path)
#             if True: app.PrintInfo("Removed file: " + file_path)

app.PrintPlain("Calculating generator ratios")
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
    elif generator_grid in sistemi_spreminjanje_parametrov:
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
app.PrintPlain("Calculated generator ratios")
app.PrintPlain(gen_ratio)

app.PrintPlain("Calculating load ratios")
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
    elif load_grid in sistemi_spreminjanje_parametrov:
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
    try: load_ratio[load] = generator.pgini/grid_type_sum[dload_grid_type[load]]
    #Če je error bo 0 in nastavimo 0
    except: load_ratio[load] = 0
app.PrintPlain("Calculated load ratios")
app.PrintPlain(load_ratio)

voltagesource_list = dfCbInfo.index.to_list()
robna_list = []
for voltagesource in app.GetCalcRelevantObjects("*.ElmVac"):
    voltagesource_name = voltagesource.loc_name
    voltagesource_grid = voltagesource.cpGrid.loc_name
    if voltagesource_name in voltagesource_list and voltagesource_grid in sistemi_spreminjanje_parametrov:
        robna_list.append(voltagesource)

#Filter for results writing
lgen_results = []
for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    if generator.cpGrid.loc_name in sistemi_izpis_rezultatov:
        lgen_results.append(generator)
        
lload_results = []
for load in app.GetCalcRelevantObjects("*.ElmLod"):
    if load.cpGrid.loc_name in sistemi_izpis_rezultatov:
        lload_results.append(load)
        
lline_results = []
for line in app.GetCalcRelevantObjects("*.ElmLne"):
    if line.cpGrid.loc_name in sistemi_izpis_rezultatov:
        lline_results.append(line)
        
ltra_results = []
for transformer in app.GetCalcRelevantObjects("*.ElmTr2"):
    transformer_name = transformer.loc_name
    if transformer.cpGrid.loc_name in sistemi_izpis_rezultatov and "/" not in transformer_name and "GT" not in transformer_name and "TES" not in transformer_name and transformer.IsOutOfService() == 0:
        ltra_results.append(transformer)
        
lterm_results = []
terminal_voltages_write_out = [110,220,400]
for terminal in app.GetCalcRelevantObjects("*.ElmTerm"):
    terminal_name = terminal.loc_name
    if terminal.cpGrid.loc_name in sistemi_izpis_rezultatov and "-" not in terminal_name and terminal.uknom in terminal_voltages_write_out and terminal.IsOutOfService() == 0:
        lterm_results.append(terminal)
    
#Uredi podatke, odstej izbrana od market itd.
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
        app.PrintPlain(f"Subtracted {column} in delta from market")

#Funkcija za nstavitev vozlisc
def setNodes(hour):
    app.PrintPlain("Nastavitev vozlisc bremen!")
    for load in load_izbrana:
        load_name = load.loc_name
        try:load.plini = dfIzbP.at[hour,load_name]
        except:app.PrintPlain(f"Ni podatka delovne P ali je 0 za {load}, brez sprememb")
        try:load.qlini = dfIzbQ.at[hour,load_name]
        except:app.PrintPlain(f"Ni podatka jalove Q ali je 0 za {load}, brez sprememb")
    for load in load_other:
        load_name = load.loc_name
        load_grid_type = dload_grid_type[load]
        load.plini = float(dfMD.at[hour,load_grid_type]) * load_ratio[load] * izkoristek_omrezja
        # app.PrintPlain(f"{load} P set")
        # load.qlini = load.plini * 
        
    #Se za gen
    app.PrintPlain("Nastavitev vozlisc generatorjev!")
    for generator in gen_izbrana:
        generator_name = generator.loc_name
        try:generator.pgini = dfIzbP.at[hour,generator_name]
        except:app.PrintPlain(f"Ni podatka delovne P ali je 0 za {generator}, brez sprememb")
        try:generator.qgini = dfIzbQ.at[hour,generator_name]
        except:app.PrintPlain(f"Ni podatka jalove Q ali je 0 za {generator}, brez sprememb")
    for generator in gen_other:
        generator_name = generator.loc_name
        generator_grid_type = dgen_grid_type[generator]
        # app.PrintPlain(f"{generator} with {generator_grid_type} P set") 
        generator.pgini = float(dfMD.at[hour,generator_grid_type]) * gen_ratio[generator]
        
    app.PrintPlain("Nastavitev robnih vozlisc!")
    for voltagesource in robna_list:
        voltagesource_name = voltagesource.loc_name
        border = dfCbInfo.at[voltagesource_name,'MEJA']
        ratio = dfCbInfo.at[voltagesource_name,'DELEZ']
        multiplier = dfCbInfo.at[voltagesource_name,'POMNOZITI']
        try: voltagesource.Pgen = float(dfCbFlow.at[hour, border]) * ratio * multiplier
        except: app.PrintPlain("Error pri nastravljanju robnega!")

def saveElementData():
    df_data_generators = pd.DataFrame(data=None)
    df_data_loads = pd.DataFrame(data=None)
    df_data_lines = pd.DataFrame(data=None)
    df_data_transformers = pd.DataFrame(data=None)
    df_data_terminals = pd.DataFrame(data=None)
    path_info_gen = os.path.join(f_output_data_directory,("Gen_Info.csv"))
    path_info_load = os.path.join(f_output_data_directory,("Load_Info.csv"))
    path_info_line = os.path.join(f_output_data_directory,("Line_Info.csv"))
    path_info_tra = os.path.join(f_output_data_directory,("Transformer_Info.csv"))
    path_info_term = os.path.join(f_output_data_directory,("Terminal_Info.csv"))
    
    app.PrintPlain("Izpis parametrov generatorjev")
    for generator in lgen_results:
        generator_name = generator.loc_name
        generator_grid = generator.cpGrid.loc_name
        try: generator_area = generator.cpArea.loc_name
        except: generator_area = "NOAREA"
        try: generator_zone = generator.cpZone.loc_name
        except: generator_area = "NOZONE"
        df_data_generators.at[generator_name, 'grid'] = generator_grid
        df_data_generators.at[generator_name, 'area'] = generator_area
        df_data_generators.at[generator_name, 'zone'] = generator_zone
    df_data_generators.to_csv(path_info_gen, encoding='utf-8', index=True)
            
    app.PrintPlain("Izpis parametrov bremen")
    for load in lload_results:
        load_name = load.loc_name
        load_grid = load.cpGrid.loc_name
        try: load_area = load.cpArea.loc_name
        except: load_area = "NOAREA"
        try: load_zone = load.cpZone.loc_name
        except: load_area = "NOZONE"
        df_data_loads.at[load_name, 'grid'] = load_grid
        df_data_loads.at[load_name, 'area'] = load_area
        df_data_loads.at[load_name, 'zone'] = load_zone
    df_data_loads.to_csv(path_info_load, encoding='utf-8', index=True)
    
    app.PrintPlain("Izpis parametrov daljnovodov")
    for line in lline_results:
        line_name = line.loc_name
        line_grid = line.cpGrid.loc_name
        try: line_area = line.cpArea.loc_name
        except: line_area = "NOAREA"
        try: line_zone = line.cpZone.loc_name
        except: line_area = "NOZONE"
        line_rated_voltage = line.typ_id.uline #Nazivna napetost v kV
        line_rated_current = round(line.typ_id.sline * 1000) #Nazivni tok v A
        line_rated_power = round(line_rated_voltage * line_rated_current * 1.73205 / 1000) # Nazivna moč MW
        df_data_lines.at[line_name, 'rated_voltage'] = line_rated_voltage
        df_data_lines.at[line_name, 'rated_current'] = line_rated_current
        df_data_lines.at[line_name, 'rated_power'] = line_rated_power
        df_data_lines.at[line_name, 'grid'] = line_grid
        df_data_lines.at[line_name, 'area'] = line_area
        df_data_lines.at[line_name, 'zone'] = line_zone
    df_data_lines.to_csv(path_info_line, encoding='utf-8', index=True)
    
    app.PrintPlain("Izpis parametrov transformatorjev")
    for transformer in ltra_results:
        transformer_name = transformer.loc_name
        transformer_grid = transformer.cpGrid.loc_name
        try: transformer_area = transformer.cpArea.loc_name
        except: transformer_area = "NOAREA"
        try: transformer_zone = transformer.cpZone.loc_name
        except: transformer_area = "NOZONE"
        df_data_transformers.at[transformer_name, 'grid'] = transformer_grid
        df_data_transformers.at[transformer_name, 'area'] = transformer_area
        df_data_transformers.at[transformer_name, 'zone'] = transformer_zone
    df_data_transformers.to_csv(path_info_tra, encoding='utf-8', index=True)
            
    app.PrintPlain("Izpis informacij zbiralk")
    for terminal in lterm_results:
        terminal_name = terminal.loc_name
        terminal_grid = terminal.cpGrid.loc_name
        terminal_nominal_voltage = terminal.uknom
        try: terminal_area = terminal.cpArea.loc_name
        except: terminal_area = "NOAREA"
        try: terminal_zone = terminal.cpZone.loc_name
        except: terminal_zone = "NOZONE"
        df_data_terminals.at[terminal_name, 'nominal_voltage'] = terminal_nominal_voltage
        df_data_terminals.at[terminal_name, 'grid'] = terminal_grid
        df_data_terminals.at[terminal_name, 'area'] = terminal_area
        df_data_terminals.at[terminal_name, 'zone'] = terminal_zone
    df_data_terminals.to_csv(path_info_term, encoding='utf-8', index=True)
    return

def contingencyCalcReport(hour):
    # Run contingency analysis
    app.PrintPlain("Izvedba contingency analiz")
    ctg = app.GetFromStudyCase("ComSimoutage")
    ctg.Execute()
    app.PrintPlain("Ivoz contingency datotek")
    if not os.path.exists(os.path.join(f_output_data_directory,"Contingencies")):
        os.makedirs(os.path.join(f_output_data_directory,"Contingencies"))
    # Export v S:\SlapnikL_Mag\Programi za analizo podatkov iz PowerFactory\Vmesni rezultati\Contingencies
    contingency_report = app.GetFromStudyCase("ComRes")
    contingency_report.iopt_exp = 6
    contingency_report.f_name = os.path.join(f_output_data_directory,"Contingencies",("Contingency_"+str(hour)+".csv"))
    # contingency_report.f_name = os.getcwd() + '\Vmesni rezultati\Contingencies\Contingency_' + str(hour) + '.csv'
    contingency_report.Execute()
    app.PrintInfo("Izvozeni contingency rezultati")
    # Clear contingency results data
    #contingency_results = app.GetFromStudyCase("Contingency Analysis AC.ElmRes")
    #contingency_results.bClear()
    return

# #PISANJE REZULTATOV V DATAFRAME
def resultExport(hour):
    app.PrintPlain("Writing element results for hour: " + str(hour))
    app.PrintPlain("Zapis rezultatov moci generatorjev")
    for generator in lgen_results:
        if generator.IsOutOfService() == 0:
            generator_name = generator.loc_name
            df_results_generator_P_set.at[generator_name, hour] = generator.GetAttribute("pgini")
            df_results_generator_Q_set.at[generator_name, hour] = generator.GetAttribute("qgini")
            
    app.PrintPlain("Zapis rezultatov moci bremen")
    for load in lload_results:
        if load.IsOutOfService() == 0:
            load_name = load.loc_name
            df_results_load_P_set.at[load_name, hour] = load.GetAttribute("plini")
            df_results_load_Q_set.at[load_name, hour] = load.GetAttribute("qlini")
    
    #Najprej zapisemo rezultate v dataframe za daljnovode
    app.PrintPlain("Zapis rezultatov obremenitev daljnovodov")
    for line in lline_results:
        line_name = line.loc_name
        if line.HasResults() == 1:
            line_loading = line.GetAttribute('c:loading')
        else:
            line_loading = int(0)
        df_results_line_loading_hourly.at[line_name, hour] = line_loading
            
    #Nato se rezultate za transformatorje
    app.PrintPlain("Zapis rezultatov obremenitev transformatorjev")
    for transformer in ltra_results:
        transformer_name = transformer.loc_name
        # hardcoded je da odstanimo trafote kjer je v imenu "/" ker so to 110/xx kV trafoti ponavadi od generatorjev
        if transformer.HasResults() == 1:
            transformer_loading = transformer.GetAttribute('c:loading')
        else:
            transformer_loading = int(0)
        df_results_transformer_loading_hourly.at[transformer_name, hour] = transformer_loading
            
    app.PrintPlain("Zapis rezultatov napetosti zbiralk")
    #Modelirane so tudi nizjenapetostne zbiralke zato izpisemo samo te z napetnstnega nivoja 110, 220. 400 kV. Ven damo tut "odcepe" ki majo v imenu -
    for terminal in lterm_results:
        terminal_name = terminal.loc_name
        if terminal.HasResults() == 1:
            terminal_voltage_pu = terminal.GetAttribute('m:u')
        else:
            terminal_voltage_pu = int(0)
        df_results_voltage_hourly.at[terminal_name, hour] = terminal_voltage_pu
    
    app.PrintPlain("Izpisani rezultati " + str(hour) + ". ure")
    return

################################################################ GLAVNI ALGORITEM PROGRAMA ###############################################################

    #calculateSumCompareToMarket(generators, loads, hour, df_market_data)
#################################################################################### MAIN #######################################################

df_results_convergence = pd.DataFrame(data=None)
df_results_line_loading_hourly = pd.DataFrame(data=None)
df_results_transformer_loading_hourly = pd.DataFrame(data=None)
df_results_voltage_hourly = pd.DataFrame(data=None)
df_results_generator_P_set = pd.DataFrame(data=None)
df_results_generator_Q_set = pd.DataFrame(data=None)
df_results_load_P_set = pd.DataFrame(data=None)
df_results_load_Q_set = pd.DataFrame(data=None)

saveElementData()

for hour in range(zacetna_ura, koncna_ura + 1, inkrement):
    t_start = time.time()
    setNodes(hour)
    t_setting = time.time()
    app.PrintPlain(f"Time needed for setting for hour {hour} data: {t_setting - t_start} seconds")
    result = ldf.Execute()
    t_calc = time.time()
    app.PrintPlain(f"Time needed for calculation for hour {hour} data: {t_calc - t_setting} seconds")
    if result == 0:
        df_results_convergence.at[hour,"Konvergenca"] = "DA"
        df_results_convergence.at[hour,"Cas"] = t_calc-t_start
        resultExport(hour)
        t_export = time.time()
        app.PrintPlain(f"Time needed for resultss for hour {hour} data: {t_export - t_calc} seconds")
        app.PrintPlain(f"Total time for hour {hour} data: {t_export - t_start} seconds")
        if contingency_report: contingencyCalcReport(hour)
    else:
        df_results_convergence.at[hour,"Konvergenca"] = "NE"
        df_results_convergence.at[hour,"Cas"] = t_calc-t_start
    app.PrintPlain(f"Total time for hour {hour} data: {t_calc - t_start} seconds")

app.PrintPlain("Shranjevanje csv datotek rezultatov")

line_loading_file_path = os.path.join(f_output_data_directory,("Line_loading.csv"))
transformer_loading_file_path = os.path.join(f_output_data_directory,("Transformer_loading.csv"))
voltage_file_path = os.path.join(f_output_data_directory,("Terminal_voltage.csv"))
generator_P_set_file_path = os.path.join(f_output_data_directory,("Generator_P.csv"))
generator_Q_set_file_path = os.path.join(f_output_data_directory,("Generator_Q.csv"))
load_P_set_file_path = os.path.join(f_output_data_directory,("Load_P.csv"))
load_Q_set_file_path = os.path.join(f_output_data_directory,("Load_Q.csv"))

df_results_line_loading_hourly.to_csv(line_loading_file_path, encoding='utf-8', index=True)
df_results_transformer_loading_hourly.to_csv(transformer_loading_file_path, encoding='utf-8', index=True)
df_results_voltage_hourly.to_csv(voltage_file_path, encoding='utf-8', index=True)
df_results_generator_P_set.to_csv(generator_P_set_file_path, encoding='utf-8', index=True)
df_results_generator_Q_set.to_csv(generator_Q_set_file_path, encoding='utf-8', index=True)
df_results_load_P_set.to_csv(load_P_set_file_path, encoding='utf-8', index=True)
df_results_load_Q_set.to_csv(load_Q_set_file_path, encoding='utf-8', index=True)

app.PrintPlain("CSV shranjen")
app.PrintPlain("Shranjevanje excel datotek rezultatov")

line_loading_file_path = os.path.join(f_output_data_directory,("Line_loading.xlsx"))
transformer_loading_file_path = os.path.join(f_output_data_directory,("Transformer_loading.xlsx"))
voltage_file_path = os.path.join(f_output_data_directory,("Terminal_voltage.xlsx"))
generator_P_set_file_path = os.path.join(f_output_data_directory,("Generator_P.xlsx"))
generator_Q_set_file_path = os.path.join(f_output_data_directory,("Generator_Q.xlsx"))
load_P_set_file_path = os.path.join(f_output_data_directory,("Load_P.xlsx"))
load_Q_set_file_path = os.path.join(f_output_data_directory,("Load_Q.xlsx"))

df_results_line_loading_hourly.to_excel(line_loading_file_path)
df_results_transformer_loading_hourly.to_excel(transformer_loading_file_path)
df_results_voltage_hourly.to_excel(voltage_file_path)
df_results_generator_P_set.to_excel(generator_P_set_file_path)
df_results_generator_Q_set.to_excel(generator_Q_set_file_path)
df_results_load_P_set.to_excel(load_P_set_file_path)
df_results_load_Q_set.to_excel(load_Q_set_file_path)    
    
app.PrintPlain("Excel shranjen")

#################################################################################### MAIN #######################################################

#################### IZPIS URE #################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
