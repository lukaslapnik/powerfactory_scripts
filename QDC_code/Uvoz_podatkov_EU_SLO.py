# -*- coding: utf-8 -*-
"""
Created on Fri Jan 27 14:44:29 2023

@author: SSIMON
"""

import os
import datetime
import pandas as pd
from tkinter import Tk
from tkinter import filedialog

############### PARAMETRI ##################################################################################
#Ali laufamo v engine mode (spyder) ali se izvaja v PowerFactory
use_powerfactory = True
# ČE JE TRUE ŠE ROČNO DEFINIRAJ TOČNO IME PROJEKTA
clear_old_data = True

voltage_table = {}
voltage_table["0"] = 750.0
voltage_table["1"] = 400.0
voltage_table["2"] = 220.0
voltage_table["3"] = 150.0
voltage_table["4"] = 120.0
voltage_table["5"] = 110.0
voltage_table["6"] = 70.0
voltage_table["7"] = 27.0
voltage_table["8"] = 330.0
voltage_table["9"] = 500.0
##########################################################################################################

# if engine_mode:
#     print("Running in engine mode")
#     sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.9")

if use_powerfactory:
    import powerfactory as pf    
    app = pf.GetApplication()
    ldf = app.GetFromStudyCase("ComLdf")
    qds = app.GetFromStudyCase("ComStatsim")
    app.ClearOutputWindow()
    user = app.GetCurrentUser()
    app.PrintInfo(f"Current user: {user}")
    activestudycase = app.GetActiveStudyCase()
    app.PrintInfo(f"Current study case: {activestudycase}")
    scenario = app.GetActiveScenario()
    app.PrintInfo(f"Current scenario: {scenario}")
    fChars = app.GetProjectFolder("chars") #Characters folder in PowerFactory software
    fLibrary = app.GetProjectFolder("lib") #Get library folder
    
###################################Izpis start cajta skripte##############################################
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
if use_powerfactory: app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
else: print("Pričetek izvajanja programa ob " + str(start_time) + ".")
##########################################################################################################

if True:
    #Open excel files
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    app.PrintPlain("Izberi vhodno mapo za SLO podatke")
    dirFolder_SLO = filedialog.askdirectory()
    file_gen_list_SLO = []
    file_load_list_SLO = []
    # Daj na true ce mamo .csv datoteke
    havecsvfiles = True
    for root, dirs, files in os.walk(dirFolder_SLO):
        for file in files:
            if havecsvfiles:
                if "PROIZVODNJA" in file and file.endswith(".csv"):
                    file_gen_list_SLO.append(os.path.join(root, file))
                if "ODJEM" in file and file.endswith(".csv"):
                    file_load_list_SLO.append(os.path.join(root, file))
            else:
                if "PROIZVODNJA" in file and file.endswith(".xlsx"):
                    file_gen_list_SLO.append(os.path.join(root, file))
                if "ODJEM" in file and file.endswith(".xlsx"):
                    file_load_list_SLO.append(os.path.join(root, file))
        
    app.PrintPlain("Izbrane datoteke SLO podatkov")
    app.PrintPlain(file_gen_list_SLO)
    app.PrintPlain(file_load_list_SLO)

    #MAPA ZA EU PODATKE
    #Open excel files
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    app.PrintPlain("Izberi vhodno mapo za ostale EU podatke")
    dirFolder_EU = filedialog.askdirectory()
    file_gen_list_EU = []
    file_load_list_EU = []
    # Daj na true ce mamo .csv datoteke
    havecsvfiles = True
    for root, dirs, files in os.walk(dirFolder_EU):
        for file in files:
            if havecsvfiles:
                if "GEN" in file and file.endswith(".csv"):
                    file_gen_list_EU.append(os.path.join(root, file))
                if "LOAD" in file and file.endswith(".csv"):
                    file_load_list_EU.append(os.path.join(root, file))
            else:
                if "GEN" in file and file.endswith(".xlsx"):
                    file_gen_list_EU.append(os.path.join(root, file))
                if "LOAD" in file and file.endswith(".xlsx"):
                    file_load_list_EU.append(os.path.join(root, file))
        
    app.PrintPlain("Izbrane datoteke EU")
    app.PrintPlain(file_gen_list_EU)
    app.PrintPlain(file_load_list_EU)

    ############################################# UVAŽANJE IN OBDELAVA PODATKOV ##############################################
    
    #FOR GENS
    dfGenP_SLO = pd.DataFrame()
    dfGenQ_SLO = pd.DataFrame()
    
    app.PrintPlain(f"Import fajla {file_gen_list_SLO[0]}")
    if havecsvfiles:
        dfDataGen_Temp_SLO = pd.read_csv(file_gen_list_SLO[0], index_col = 0)
    else:
        file1 = pd.ExcelFile(file_gen_list_SLO[0])
        file1_sheets = file1.sheet_names
        dfDataGen_Temp_SLO = pd.DataFrame()
        dfDataGen_Temp_SLO = file1.parse(file1_sheets[0], index_col = 0)
        
    dfGenP_SLO = dfDataGen_Temp_SLO[dfDataGen_Temp_SLO['P/Q'] == "MW"].drop(["P/Q"], axis = 'columns')
    dfGenQ_SLO = dfDataGen_Temp_SLO[dfDataGen_Temp_SLO['P/Q'] == "Mvar"].drop(["P/Q"], axis = 'columns')
    #Replace missing data with 0
    dfGenP_SLO = dfGenP_SLO.interpolate(method='linear', axis = 1)
    dfGenQ_SLO = dfGenQ_SLO.interpolate(method='linear', axis = 1)
    dfGenP_SLO = dfGenP_SLO.fillna(0.0)
    dfGenQ_SLO = dfGenQ_SLO.fillna(0.0)
        
    app.PrintPlain(dfGenP_SLO) 
    app.PrintPlain(dfGenQ_SLO)   
    
    #FOR LOADS
    dfLoadP_SLO = pd.DataFrame()
    dfLoadQ_SLO = pd.DataFrame()
    
    app.PrintPlain(f"Import fajla {file_load_list_SLO[0]}")
    if havecsvfiles:
        dfDataLoad_Temp_SLO = pd.read_csv(file_load_list_SLO[0], index_col = 0)
    else:
        file2 = pd.ExcelFile(file_load_list_SLO[0])
        file2_sheets = file2.sheet_names
        dfDataLoad_Temp_SLO = pd.DataFrame()
        dfDataLoad_Temp_SLO = file2.parse(file2_sheets[0], index_col = 0)
            
    dfLoadP_SLO = dfDataLoad_Temp_SLO[dfDataLoad_Temp_SLO['P/Q'] == "MW"].drop(["P/Q"], axis = 'columns')
    dfLoadQ_SLO = dfDataLoad_Temp_SLO[dfDataLoad_Temp_SLO['P/Q'] == "Mvar"].drop(["P/Q"], axis = 'columns')
    #Replace missing data with 0
    dfLoadP_SLO = dfLoadP_SLO.interpolate(method='linear', axis = 1)
    dfLoadQ_SLO = dfLoadQ_SLO.interpolate(method='linear', axis = 1)
    dfLoadP_SLO = dfLoadP_SLO.fillna(0.0)
    dfLoadQ_SLO = dfLoadQ_SLO.fillna(0.0)

    app.PrintPlain(dfLoadP_SLO) 
    app.PrintPlain(dfLoadQ_SLO)   
    
    app.PrintPlain("Datoteke za SLO uvozene in obdelane")

################### UVOZENO ZA SLO, ZDAJ SE OSTALE DRZAVE #############################################
 
app.PrintPlain("Zacetek uvoza EU podatkov")
#FOR GENS
dfDataGen_EU = pd.DataFrame()
dfGenNodeType_EU = pd.DataFrame()
dfGenU_EU = pd.DataFrame()
dfGenP_EU = pd.DataFrame()
dfGenQ_EU = pd.DataFrame()

for file1_loc in file_gen_list_EU:
    app.PrintPlain(f"Import fajla {file1_loc}")
    if havecsvfiles:
        dfDataGen_Temp_EU = pd.read_csv(file1_loc, index_col = 0)
    else:
        file1 = pd.ExcelFile(file1_loc)
        file1_sheets = file1.sheet_names
        dfDataGen_Temp_EU = pd.DataFrame()
        dfDataGen_Temp_EU = file1.parse(file1_sheets[0], index_col = 0)
    
    dfGenNodeType_Temp_EU = pd.DataFrame()
    dfGenU_Temp_EU = pd.DataFrame()
    dfGenU_Temp_unom_EU = pd.DataFrame()
    dfGenP_Temp_EU = pd.DataFrame()
    dfGenQ_Temp_EU = pd.DataFrame()
    
    #Shranimo node type (PV/PQ)
    dfGenNodeType_Temp_EU = dfDataGen_Temp_EU[dfDataGen_Temp_EU['U_P/Q'] == "Node Type"].drop(["U_P/Q"], axis = 'columns')
    
    # dfGenU_Temp = dfDataGen_Temp[dfDataGen_Temp['U_P/Q'] == "U (kV)"]
    # dfGenU_Temp_unom['Napetost'] = dfGenU_Temp['Napetost']
    
    #Shranimo podatke napetosti uporabljene za PV tip
    dfGenU_Temp_EU = dfDataGen_Temp_EU[dfDataGen_Temp_EU['U_P/Q'] == "U (kV)"].drop(["U_P/Q"], axis = 'columns')
    
    #Tu delimo napetost z nazivno da dobimo p.u
    #Nazivna napetost generatorja/proizvodnje je 7. znak v imenu
    dfGenU_Temp_unom_EU.index = dfGenU_Temp_EU.index
    dfGenU_Temp_unom_EU['Napetost'] = dfGenU_Temp_EU.index.astype(str).str[6].map(voltage_table) 
    # app.PrintPlain(dfGenU_Temp_unom_EU)
    #Napetost nato delimo z nazivno
    dfGenU_Temp_EU = dfGenU_Temp_EU.divide(dfGenU_Temp_unom_EU['Napetost'], axis=0)
    
    #Dobi podatke P in Q in dropni stolpec tipa
    dfGenP_Temp_EU = dfDataGen_Temp_EU[dfDataGen_Temp_EU['U_P/Q'] == "Gen (MW)"].drop(["U_P/Q"], axis = 'columns')
    dfGenQ_Temp_EU = dfDataGen_Temp_EU[dfDataGen_Temp_EU['U_P/Q'] == "Gen (Mvar)"].drop(["U_P/Q"], axis = 'columns')
    
    dfDataGen_EU = pd.concat([dfDataGen_EU, dfDataGen_Temp_EU])
    dfGenNodeType_EU = pd.concat([dfGenNodeType_EU, dfGenNodeType_Temp_EU])
    dfGenU_EU = pd.concat([dfGenU_EU, dfGenU_Temp_EU])
    dfGenP_EU = pd.concat([dfGenP_EU, dfGenP_Temp_EU])
    dfGenQ_EU = pd.concat([dfGenQ_EU, dfGenQ_Temp_EU])
    
# Handle missing data
dfGenNodeType_EU = dfGenNodeType_EU.fillna(2) #Replace with 2 for missing data (PV)
# V datoteki 0 = PQ, 2 = PV spremenimo v 0 = PV, 1 = PQ
dfGenNodeType_EU = dfGenNodeType_EU.replace(to_replace=0, value=1)
dfGenNodeType_EU = dfGenNodeType_EU.replace(to_replace=2, value=0)

dfGenU_EU = dfGenU_EU.interpolate(method='linear', axis = 1) #To pa interpolira manjkajoče vrednosti
dfGenP_EU = dfGenP_EU.interpolate(method='linear', axis = 1)
dfGenQ_EU = dfGenQ_EU.interpolate(method='linear', axis = 1)
dfGenU_EU = dfGenU_EU.fillna(1.0)
dfGenP_EU = dfGenP_EU.fillna(0.0)
dfGenQ_EU = dfGenQ_EU.fillna(0.0)
dfGenP_EU = dfGenP_EU.multiply(-1)
dfGenQ_EU = dfGenQ_EU.multiply(-1)

app.PrintPlain(dfGenP_EU) 
app.PrintPlain(dfGenU_EU)   
app.PrintPlain(dfGenQ_EU)   
    
#FOR LOADS
dfDataLoad_EU = pd.DataFrame()
dfLoadP_EU = pd.DataFrame()
dfLoadQ_EU = pd.DataFrame()

for file2_loc in file_load_list_EU:
    app.PrintPlain(f"Import fajla {file2_loc}")
    if havecsvfiles:
        dfDataLoad_Temp_EU = pd.read_csv(file2_loc, index_col = 0)
    else:
        file2 = pd.ExcelFile(file2_loc)
        file2_sheets = file2.sheet_names
        dfDataLoad_Temp_EU = pd.DataFrame()
        dfDataLoad_Temp_EU = file2.parse(file2_sheets[0], index_col = 0)
        
    dfLoadP_Temp_EU = pd.DataFrame()
    dfLoadQ_Temp_EU = pd.DataFrame()
    
    dfLoadP_Temp_EU = dfDataLoad_Temp_EU[dfDataLoad_Temp_EU['P/Q'] == "MW"].drop(["P/Q"], axis = 'columns')
    dfLoadQ_Temp_EU = dfDataLoad_Temp_EU[dfDataLoad_Temp_EU['P/Q'] == "Mvar"].drop(["P/Q"], axis = 'columns')
    
    dfDataLoad_EU = pd.concat([dfDataLoad_EU, dfDataLoad_Temp_EU])
    dfLoadP_EU = pd.concat([dfLoadP_EU, dfLoadP_Temp_EU])
    dfLoadQ_EU = pd.concat([dfLoadQ_EU, dfLoadQ_Temp_EU])

dfLoadP_EU = dfLoadP_EU.interpolate(method='linear', axis = 1)
dfLoadQ_EU = dfLoadQ_EU.interpolate(method='linear', axis = 1)
dfLoadP_EU = dfLoadP_EU.fillna(0.0)
dfLoadQ_EU = dfLoadQ_EU.fillna(0.0)
    

app.PrintPlain("Datoteke uvozene in obdelane")
    
############################ DATA IMPORTED ######################

app.PrintPlain("Zacenjam uvoz v powerfactory")

#Delete current station(external) controllers in project
if clear_old_data: 
    for stationcontroller in app.GetCalcRelevantObjects("*.ElmStactrl"): 
        stationcontroller.Delete()

generator_station_voltage_list = []
generators_slo = []
generators_adaptive = []
generators_pq = []

#Izpisi generatorje na isti zbiralki
for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    #Station kontrolerje za prvi generator v vsakem RTP za dan napetostni nivo
    generator_name = generator.loc_name
    #Preveri da generator ni v slo
    if generator_name in dfGenP_SLO.index:
        #Generatorji v slo za katere uporabimo podatke
        generators_slo.append(generator)
        app.PrintPlain(f"Generator SLO {generator}")
    elif generator_name in dfGenP_EU.index:
        try:
            #Generator ma okej določeno vse
            generator_station_voltage = generator.bus1.cterm.cpSubstat.loc_name + "_" + generator_name[6]
            if generator_station_voltage in generator_station_voltage_list:
                generators_pq.append(generator)
                app.PrintPlain(f"generator {generator} bo SAMO PQ substation {generator_station_voltage}")
            else:
                generators_adaptive.append(generator)
                generator_station_voltage_list.append(generator_station_voltage)
                app.PrintPlain(f"generator {generator} bo ADAPTIVE substation {generator_station_voltage}")
        except:
            app.PrintWarn(f"napaka za {generator}, nima dolocenega substation oz nek drugi problem....")
    else:
        app.PrintWarn(f"Neznani generator {generator}")
        
############################ DATA IMPORTED ######################

app.PrintPlain("Zacenjam uvoz v powerfactory")

#Make time scale for a year in libry folder
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
    timescale_vector = list(range(0,8760))
    timescale.SetAttribute("scale", timescale_vector)
    app.PrintPlain("Edited " + timescale_name + " vector!")
    
#Če je slo se nastavi samo PQ
for generator in generators_slo:
    generator_name = generator.GetAttribute("loc_name")
    #Klasična bremena/odjem
    app.PrintPlain(f"Proizvodnja {generator}")
    try:
        #Assign P vector
        app.PrintPlain(f"Generator/proizvodnja {generator}")
        # Remove old data
        if clear_old_data:
            for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaPgini = generator.CreateObject("ChaVec", "pgini")
        chaPgini.SetAttribute("scale", timescale)
        chaPgini.SetAttribute("vector", dfGenP_SLO.loc[generator_name].to_list())
        chaPgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Nastavil podatke P {chaPgini} za {generator}")
    except:
        app.PrintWarn(f"Napaka nastavljanja P za {generator}, preveri vhodne datoteke!")
    
    try:
        #Assign Q vector
        # Remove old data
        if clear_old_data:
            for chaOld in generator.GetContents("qgini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaQgini = generator.CreateObject("ChaVec", "qgini")
        chaQgini.SetAttribute("scale", timescale)
        chaQgini.SetAttribute("vector", dfGenQ_SLO.loc[generator_name].to_list())
        chaQgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Nastavil podatke Q {chaQgini} za {generator}")
    except:
        app.PrintWarn(f"Napaka nastavljanja Q za {generator}, preveri vhodne datoteke!")
    
    
#Če je adaptive se dela station control
for generator in generators_adaptive:
    generator_name = generator.GetAttribute("loc_name")
    #Assign P vector
    # Remove old data
    if clear_old_data:
        for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
    # Assign controller to generator
    chaPgini = generator.CreateObject("ChaVec", "pgini")
    chaPgini.SetAttribute("scale", timescale)
    chaPgini.SetAttribute("vector", dfGenP_EU.loc[generator_name].to_list())
    chaPgini.SetAttribute("usage", 2)
    app.PrintPlain(f"Katkteristika P {chaPgini} za generator {generator}")
    
    #Create station controller for generator
    stationcontroller = generator.GetParent().CreateObject("ElmStactrl", generator_name + "_SC.ElmStactrl") 
    # Remove old data
    if clear_old_data:
        for elmStacontOld in generator.GetContents("c_pstac*.ElmStactrl"): elmStacontOld.Delete() 
    # Assign controller to generator
    generator.SetAttribute("c_pstac", stationcontroller)
    # Assign control nodes (generator terminal for voltage and cub for reactive power Q) 
    stationcontroller.SetAttribute("rembar", generator.GetAttribute("bus1").GetAttribute("cterm"))
    stationcontroller.SetAttribute("p_cub", generator.GetAttribute("bus1"))
    app.PrintPlain(f"Narejen {stationcontroller} za generator {generator}")
    
    # Characteristic for Type
    # Remove old data
    if clear_old_data:
        for chaOld in stationcontroller.GetContents("i_ctrl*.ChaVec"): chaOld.Delete() 
    # Assign controller to generator
    chaSC_Type = stationcontroller.CreateObject("ChaVec", "i_ctrl") #attribute name, has to be the name of the ChaRef to make the link
    chaSC_Type.SetAttribute("scale", timescale)
    chaSC_Type.SetAttribute("vector", dfGenNodeType_EU.loc[generator_name].to_list())
    chaSC_Type.SetAttribute("usage", 2)
    app.PrintPlain(f"Katkteristika PV/PQ {chaSC_Type} za station kontroler {stationcontroller}")
    
    # Create characteritic for voltage U
    # Remove old data
    if clear_old_data:
        for chaOld in stationcontroller.GetContents("usetp*.ChaVec"): chaOld.Delete() 
    # Assign controller to generator
    chaSC_U = stationcontroller.CreateObject("ChaVec", "usetp") #attribute name, has to be the name of the ChaRef to make the link
    chaSC_U.SetAttribute("scale", timescale)
    chaSC_U.SetAttribute("vector", dfGenU_EU.loc[generator_name].to_list())
    chaSC_U.SetAttribute("usage", 2)
    app.PrintPlain(f"Karakteristika U {chaSC_U} za station kontroler {stationcontroller}")
    
    # Characteristic for Q
    # Remove old data
    if clear_old_data:
        for chaOld in stationcontroller.GetContents("qsetp*.ChaVec"): chaOld.Delete() 
    # Assign controller to generator
    chaSC_Q = stationcontroller.CreateObject("ChaVec", "qsetp") #attribute name, has to be the name of the ChaRef to make the link
    chaSC_Q.SetAttribute("scale", timescale)
    chaSC_Q.SetAttribute("vector", dfGenQ_EU.loc[generator_name].to_list())
    chaSC_Q.SetAttribute("usage", 2)
    app.PrintPlain(f"Karakteristika Q {chaSC_Q}za station kontroler {stationcontroller}")
        
for generator in generators_pq:
    generator_name = generator.GetAttribute("loc_name")
    #Klasična bremena/odjem
    app.PrintPlain(f"Proizvodnja {generator} samo kot PQ brez station controllerkja, nastavljeno na constQ mode")
    generator.av_mode = "constq"
    
    try:
        #Assign P vector
        app.PrintPlain(f"Generator/proizvodnja {generator}")
        # Remove old data
        if clear_old_data:
            for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaPgini = generator.CreateObject("ChaVec", "pgini")
        chaPgini.SetAttribute("scale", timescale)
        chaPgini.SetAttribute("vector", dfGenP_EU.loc[generator_name].to_list())
        chaPgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Nastavil podatke P {chaPgini} za {generator}")
    except:
        app.PrintWarn(f"Napaka nastavljanja P za {generator}, preveri vhodne datoteke!")
    
    try:
        #Assign Q vector
        # Remove old data
        if clear_old_data:
            for chaOld in generator.GetContents("qgini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaQgini = generator.CreateObject("ChaVec", "qgini")
        chaQgini.SetAttribute("scale", timescale)
        chaQgini.SetAttribute("vector", dfGenQ_EU.loc[generator_name].to_list())
        chaQgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Nastavil podatke Q {chaQgini} za {generator}")
    except:
        app.PrintWarn(f"Napaka nastavljanja Q za {generator}, preveri vhodne datoteke!")
    
    
for load in app.GetCalcRelevantObjects("*.ElmLod"):
    load_name = load.GetAttribute("loc_name")
    if load_name in dfLoadP_SLO.index:
        #Assign P vector
        # Remove old data
        if clear_old_data:
            for chaOld in load.GetContents("plini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaPlini = load.CreateObject("ChaVec", "plini")
        chaPlini.SetAttribute("scale", timescale)
        chaPlini.SetAttribute("vector", dfLoadP_SLO.loc[load_name].to_list())
        chaPlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPlini} for {load}")
        
        #Assign Q vector
        # Remove old data
        if clear_old_data:
            for chaOld in load.GetContents("qlini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaQlini = load.CreateObject("ChaVec", "qlini")
        chaQlini.SetAttribute("scale", timescale)
        chaQlini.SetAttribute("vector", dfLoadQ_SLO.loc[load_name].to_list())
        chaQlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaQlini} for {load}")
        
    elif load_name in dfLoadP_EU.index:
        #Assign P vector
        # Remove old data
        if clear_old_data:
            for chaOld in load.GetContents("plini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaPlini = load.CreateObject("ChaVec", "plini")
        chaPlini.SetAttribute("scale", timescale)
        chaPlini.SetAttribute("vector", dfLoadP_EU.loc[load_name].to_list())
        chaPlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPlini} for {load}")
        
        #Assign Q vector
        # Remove old data
        if clear_old_data:
            for chaOld in load.GetContents("qlini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaQlini = load.CreateObject("ChaVec", "qlini")
        chaQlini.SetAttribute("scale", timescale)
        chaQlini.SetAttribute("vector", dfLoadQ_EU.loc[load_name].to_list())
        chaQlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaQlini} for {load}")
    else:
        app.PrintWarn(f"Breme/Odjem {load} nima podatkov")
    

#################### IZPIS URE ################# KONEC #############################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
if use_powerfactory: app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
else: print("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')