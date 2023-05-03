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

############### PARAMETRI ##################################################################################
#Ali laufamo v engine mode (spyder) ali se izvaja v PowerFactory
engine_mode = False
# ČE JE TRUE ŠE ROČNO DEFINIRAJ TOČNO IME PROJEKTA
define_project_name = "ENTSO-E_NT2030_RNPS_2023-2032_v18(4)"

#Parametri za izračun jalovih moči za gen, load, vac..... načeloma če delamo DC loadflow ni važno
# Za AC loadflow je treba porihtat oz najt neke boljše načine dodeljevanja jalovih.
clear_output_folder = False 
change_Q_params = False
orignal_cosfi = True
generator_PQ_ratio = 0.3
load_PQ_ratio = 0.3
voltagesource_PQ_ratio = 0
grid_efficiency = 0.97 #(1-losses)

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

# grids_modify_values = ['AT00','DE00','SI00','RS00',
#                         'ITN1','HU00','HR00',
#                         'ELES Interconnectios','BA00']

grids_modify_values = ['SI00','ITN1','HU00','HR00','ELES Interconnectios']

grids_to_write_results = ['SI00', 'ELES Interconnectios']
##########################################################################################################

if engine_mode:
    print("Running in engine mode")
    sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.9")

import powerfactory as pf    
app = pf.GetApplication()
app.ClearOutputWindow()
ldf = app.GetFromStudyCase("ComLdf")

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

if clear_output_folder:
    app.PrintPlain("Clearing partial results data")
    # Find folders in partial results folder
    #file_list = list()
    for root, dirs, files in os.walk(f_output_data_directory, topdown = False):
        for file in files:
            file_path = os.path.join(root, file)
            os.remove(file_path)
            #file_list.append(file_path)
            if True: app.PrintInfo("Removed file: " + file_path)

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
    elif generator_grid in grids_modify_values:
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
    elif load_grid in grids_modify_values:
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

# app.PrintPlain(dfIzbP)
# app.PrintPlain(dfIzbQ)
# app.PrintPlain(dfMD)

#Filter for results writing
lgen_results = []
for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    if generator.cpGrid.loc_name in grids_to_write_results:
        lgen_results.append(generator)
        
lload_results = []
for load in app.GetCalcRelevantObjects("*.ElmLod"):
    if load.cpGrid.loc_name in grids_to_write_results:
        lload_results.append(load)
        
lline_results = []
for line in app.GetCalcRelevantObjects("*.ElmLne"):
    if line.cpGrid.loc_name in grids_to_write_results:
        lline_results.append(line)
        
ltra_results = []
for transformer in app.GetCalcRelevantObjects("*.ElmTr2"):
    transformer_name = transformer.loc_name
    if transformer.cpGrid.loc_name in grids_to_write_results and "/" not in transformer_name and "GT" not in transformer_name and "TES" not in transformer_name and transformer.IsOutOfService() == 0:
        ltra_results.append(transformer)
        
lterm_results = []
terminal_voltages_write_out = [110,220,400]
for terminal in app.GetCalcRelevantObjects("*.ElmTerm"):
    terminal_name = terminal.loc_name
    if terminal.cpGrid.loc_name in grids_to_write_results and "-" not in terminal_name and terminal.uknom in terminal_voltages_write_out and terminal.IsOutOfService() == 0:
        lterm_results.append(terminal)
        
        
def setGenLoad(hour):
    for load in load_izbrana:
        load_name = load.loc_name
        try:load.plini = dfIzbP.at[hour,load_name]
        except:app.PrintPlain(f"No P data or 0 for {load}")
        try:load.qlini = dfIzbQ.at[hour,load_name]
        except:app.PrintPlain(f"No Q data or 0 for {load}")
    for load in load_other:
        load_name = load.loc_name
        load_grid_type = dload_grid_type[load]
        load.plini = float(dfMD.at[hour,load_grid_type]) * load_ratio[load] * grid_efficiency
        # app.PrintPlain(f"{load} P set")
        # load.qlini = load.plini * 
    #Se za gen
    for generator in gen_izbrana:
        generator_name = generator.loc_name
        try:generator.pgini = dfIzbP.at[hour,generator_name]
        except:app.PrintPlain(f"No P data or 0 for {generator}, not changing value")
        try:generator.qgini = dfIzbQ.at[hour,generator_name]
        except:app.PrintPlain(f"No Q data or 0 for {generator}, not changing value")
    for generator in gen_other:
        generator_name = generator.loc_name
        generator_grid_type = dgen_grid_type[generator]
        # app.PrintPlain(f"{generator} with {generator_grid_type} P set") 
        generator.pgini = float(dfMD.at[hour,generator_grid_type]) * gen_ratio[generator]

# #ZAPIS STATUSA IZRACUNA (konvergiralo/ni konvergiralo), mogoče še kaj drugega
# def saveCalcStatus(hour, status, t_calc):
#     df_results_calcstatus = pd.DataFrame()
#     #Shranimo status: 0=OK, 1=Divergenca notranjih zank, 2=Divergenca zunanjih zank
#     df_results_calcstatus.at[hour,'convergence'] = int(status)
#     df_results_calcstatus.at[hour,'calculation_time'] = int(t_calc)
#     loading_file_path = os.getcwd()  + r'/Vmesni rezultati/Calculation/Calcstatus_hour_' + str(hour) +'.csv'
#     df_results_calcstatus.to_csv(loading_file_path, encoding='utf-8', index=True)
#     return

#PISANJE REZULTATOV V EXCEL
def shraniVmesneRezultateCsv(hour):
    app.PrintPlain("Writing element results for hour: " + str(hour))
    
    df_results_line_loading_hourly = pd.DataFrame(data=None)
    df_results_transformer_loading_hourly = pd.DataFrame(data=None)
    df_results_voltage_hourly = pd.DataFrame(data=None)
    df_results_generator_P_set = pd.DataFrame(data=None)
    df_results_generator_Q_set = pd.DataFrame(data=None)
    df_results_load_P_set = pd.DataFrame(data=None)
    df_results_load_Q_set = pd.DataFrame(data=None)
    
    app.PrintPlain("Zapis rezultatov moci generatorjev")
    for generator in lgen_results:
        if generator.IsOutOfService() == 0:
            df_results_generator_P_set.at[generator_name, hour] = generator.pgini
            df_results_generator_Q_set.at[generator_name, hour] = generator.qgini
            
    app.PrintPlain("Zapis rezultatov moci bremen")
    for load in lload_results:
        load_name = load.loc_name
        if load.IsOutOfService() == 0:
            df_results_load_P_set.at[load_name, hour] = load.plini
            df_results_load_Q_set.at[load_name, hour] = load.qlini
    
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
    
            
    app.PrintPlain("Zacetek zapisal excel datotek " + str(hour) + ". ure")
    #app.PrintPlain(df_results_lines)
    line_loading_file_path = os.path.join(f_output_data_directory,("Line_loading_" + str(hour) +".csv"))
    transformer_loading_file_path = os.path.join(f_output_data_directory,("Transformer_loading_" + str(hour) +".csv"))
    voltage_file_path = os.path.join(f_output_data_directory,("Terminal_voltage_" + str(hour) +".csv"))
    generator_P_set_file_path = os.path.join(f_output_data_directory,("Generator_P_" + str(hour) +".csv"))
    generator_Q_set_file_path = os.path.join(f_output_data_directory,("Generator_Q_" + str(hour) +".csv"))
    load_P_set_file_path = os.path.join(f_output_data_directory,("Load_P_" + str(hour) +".csv"))
    load_Q_set_file_path = os.path.join(f_output_data_directory,("Load_Q_" + str(hour) +".csv"))
    
    df_results_line_loading_hourly.to_csv(line_loading_file_path, encoding='utf-8', index=True)
    df_results_transformer_loading_hourly.to_csv(transformer_loading_file_path, encoding='utf-8', index=True)
    df_results_voltage_hourly.to_csv(voltage_file_path, encoding='utf-8', index=True)
    df_results_generator_P_set.to_csv(generator_P_set_file_path, encoding='utf-8', index=True)
    df_results_generator_Q_set.to_csv(generator_Q_set_file_path, encoding='utf-8', index=True)
    df_results_load_P_set.to_csv(load_P_set_file_path, encoding='utf-8', index=True)
    df_results_load_Q_set.to_csv(load_Q_set_file_path, encoding='utf-8', index=True)
    
    app.PrintPlain("Izpisani in shranjeni rezultati " + str(hour) + ". ure")
    return

################################################################ GLAVNI ALGORITEM PROGRAMA ###############################################################

    #calculateSumCompareToMarket(generators, loads, hour, df_market_data)
#################################################################################### MAIN #######################################################

for hour in range(1,20):
    t_start = time.time()
    setGenLoad(hour)
    t_setting = time.time()
    app.PrintPlain(f"Time needed for setting for hour {hour} data: {t_setting - t_start} seconds")
    result = ldf.Execute()
    t_calc = time.time()
    app.PrintPlain(f"Time needed for calculation for hour {hour} data: {t_calc - t_setting} seconds")
    if result == 0:
        shraniVmesneRezultateCsv(hour)
        t_export = time.time()
        app.PrintPlain(f"Time needed for resultss for hour {hour} data: {t_export - t_calc} seconds")
        app.PrintPlain(f"Total time for hour {hour} data: {t_export - t_start} seconds")
    app.PrintPlain(f"Total time for hour {hour} data: {t_calc - t_start} seconds")

#################################################################################### MAIN #######################################################

#################### IZPIS URE #################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
