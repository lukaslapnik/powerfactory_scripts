# -*- coding: utf-8 -*-
"""
Created on Fri Jan 27 14:44:29 2023

@author: SSIMON
"""

import os
import datetime
import sys
import pandas as pd
from tkinter import Tk
from tkinter import filedialog

############### PARAMETRI ##################################################################################
#Ali laufamo v engine mode (spyder) ali se izvaja v PowerFactory
engine_mode = False
# ČE JE TRUE ŠE ROČNO DEFINIRAJ TOČNO IME PROJEKTA
define_project_name = "your_powerfactory_project_name"

#Parametri za izračun jalovih moči za gen, load, vac..... načeloma če delamo DC loadflow ni važno
# Za AC loadflow je treba porihtat oz najt neke boljše načine dodeljevanja jalovih. 
##########################################################################################################

if engine_mode:
    print("Running in engine mode")
    sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.9")

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
    
    print("User: " + str(user))
    print("Project: " + str(prj))
    print("Study Case: " + str(activestudycase))
    print("Scenario: " + str(scenario))
    
###################################Izpis start cajta skripte##############################################
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
# else: print("Pričetek izvajanja programa ob " + str(start_time) + ".")
##########################################################################################################

# if True:
#     # PF ma omejitev 40 znakov zato krajšamo zarad imen karakteristik 
#     # na koncu se dodaja P in Q in če ma element 40 znakov bi mela karakteristika 41 kar vrže error
#     for generator in generators:
#         if len(generator.loc_name) > 38: generator.loc_name = generator.loc_name[:-1]
#         if len(generator.loc_name) > 38: generator.loc_name = generator.loc_name[:-1]
#     for load in loads:
#         if len(load.loc_name) > 38: load.loc_name = load.loc_name[:-1]
#         if len(load.loc_name) > 38: load.loc_name = load.loc_name[:-1]
#     for voltagesource in voltagesources:
#         if len(voltagesource.loc_name) > 38: voltagesource.loc_name = voltagesource.loc_name[:-1]
#         if len(voltagesource.loc_name) > 38: voltagesource.loc_name = voltagesource.loc_name[:-1]

if True:
    #Open excel files
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    app.PrintPlain("Izberi vhodno mapo")
    dirFolder = filedialog.askdirectory()
    file_gen_list = []
    file_load_list = []
    # Daj na true ce mamo .csv datoteke
    havecsvfiles = True
    for root, dirs, files in os.walk(dirFolder):
        for file in files:
            if havecsvfiles:
                if "GEN" in file and file.endswith(".csv"):
                    file_gen_list.append(os.path.join(root, file))
                if "LOAD" in file and file.endswith(".csv"):
                    file_load_list.append(os.path.join(root, file))
            else:
                if "GEN" in file and file.endswith(".xlsx"):
                    file_gen_list.append(os.path.join(root, file))
                if "LOAD" in file and file.endswith(".xlsx"):
                    file_load_list.append(os.path.join(root, file))
        
    app.PrintPlain("Mapa izbrana, uvažam datoteke")
    # print(file_gen_list)
    # print(file_load_list)
 
    #FOR GENS
    dfDataGen = pd.DataFrame()
    dfGenNodeType = pd.DataFrame()
    dfGenU = pd.DataFrame()
    dfGenP = pd.DataFrame()
    dfGenQ = pd.DataFrame()
    
    for file1_loc in file_gen_list:
        print(f"Import fajla {file1_loc}")
        if havecsvfiles:
            dfDataGen_Temp = pd.read_csv(file1_loc, index_col = 1)
        else:
            file1 = pd.ExcelFile(file1_loc)
            file1_sheets = file1.sheet_names
            dfDataGen_Temp = pd.DataFrame()
            dfDataGen_Temp = file1.parse(file1_sheets[0], index_col = 1)
        
        dfGenNodeType_Temp = pd.DataFrame()
        dfGenU_Temp = pd.DataFrame()
        dfGenU_Temp_unom = pd.DataFrame()
        dfGenP_Temp = pd.DataFrame()
        dfGenQ_Temp = pd.DataFrame()
        
        dfGenNodeType_Temp = dfDataGen_Temp[dfDataGen_Temp['U_P/Q'] == "Node Type"].drop(["Node name", "Napetost", "U_P/Q"], axis = 'columns')
        dfGenU_Temp = dfDataGen_Temp[dfDataGen_Temp['U_P/Q'] == "U (kV)"]
        dfGenU_Temp_unom['Napetost'] = dfGenU_Temp['Napetost']
        dfGenU_Temp = dfDataGen_Temp[dfDataGen_Temp['U_P/Q'] == "U (kV)"].drop(["Node name", "Napetost", "U_P/Q"], axis = 'columns')
        dfGenU_Temp = dfGenU_Temp.divide(dfGenU_Temp_unom['Napetost'], axis=0)
        dfGenP_Temp = dfDataGen_Temp[dfDataGen_Temp['U_P/Q'] == "Gen (MW)"].drop(["Node name", "Napetost", "U_P/Q"], axis = 'columns')
        dfGenQ_Temp = dfDataGen_Temp[dfDataGen_Temp['U_P/Q'] == "Gen (Mvar)"].drop(["Node name", "Napetost", "U_P/Q"], axis = 'columns')
        
        dfDataGen = pd.concat([dfDataGen, dfDataGen_Temp])
        dfGenNodeType = pd.concat([dfGenNodeType, dfGenNodeType_Temp])
        dfGenU = pd.concat([dfGenU, dfGenU_Temp])
        dfGenP = pd.concat([dfGenP, dfGenP_Temp])
        dfGenQ = pd.concat([dfGenQ, dfGenQ_Temp])
        
        # Handle missing data
        dfGenNodeType = dfGenNodeType.fillna(2.0) #Replace with 2 for missing data (PV)
        # V datoteki 0 = PQ, 2 = PV spremenimo v 0 = PV, 1 = PQ
        dfGenNodeType = dfGenNodeType.replace(to_replace=0, value=1)
        dfGenNodeType = dfGenNodeType.replace(to_replace=2, value=0)
        
        dfGenU = dfGenU.interpolate(method='linear', axis = 1) #To pa interpolira manjkajoče vrednosti
        dfGenP = dfGenP.interpolate(method='linear', axis = 1)
        dfGenQ = dfGenQ.interpolate(method='linear', axis = 1)
        dfGenU = dfGenU.fillna(1.0)
        dfGenP = dfGenP.fillna(0.0)
        dfGenQ = dfGenQ.fillna(0.0)
        
        
    #FOR LOADS
    dfDataLoad = pd.DataFrame()
    dfLoadP = pd.DataFrame()
    dfLoadQ = pd.DataFrame()
    
    for file2_loc in file_load_list:
        print(f"Import fajla {file1_loc}")
        if havecsvfiles:
            dfDataLoad_Temp = pd.read_csv(file2_loc, index_col = 1)
        else:
            file2 = pd.ExcelFile(file2_loc)
            file2_sheets = file2.sheet_names
            dfDataLoad_Temp = pd.DataFrame()
            dfDataLoad_Temp = file2.parse(file2_sheets[0], index_col = 1)
            
        dfLoadP_Temp = pd.DataFrame()
        dfLoadQ_Temp = pd.DataFrame()
        
        dfLoadP_Temp = dfDataLoad_Temp[dfDataLoad_Temp['P/Q'] == "MW"].drop(["Node name", "P/Q"], axis = 'columns')
        dfLoadQ_Temp = dfDataLoad_Temp[dfDataLoad_Temp['P/Q'] == "Mvar"].drop(["Node name", "P/Q"], axis = 'columns')
        
        dfDataLoad = pd.concat([dfDataLoad, dfDataLoad_Temp])
        dfLoadP = pd.concat([dfLoadP, dfLoadP_Temp])
        dfLoadQ = pd.concat([dfLoadQ, dfLoadQ_Temp])
        
        dfLoadP = dfLoadP.interpolate(method='linear', axis = 1)
        dfLoadQ = dfLoadQ.interpolate(method='linear', axis = 1)
        dfLoadP = dfLoadP.fillna(0.0)
        dfLoadQ = dfLoadQ.fillna(0.0)

    app.PrintPlain("Datoteke uvozene in obdelane")
    
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
    timescale_vector = list(range(1,8761))
    timescale.SetAttribute("scale", timescale_vector)
    app.PrintPlain("Edited " + timescale_name + " vector!")

#Delete current station(external) controllers in project
if False: 
    for stationcontroller in app.GetCalcRelevantObjects("*.ElmStactrl"): 
        stationcontroller.Delete()

for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    generator_name = generator.GetAttribute("loc_name")
    if generator_name in dfGenNodeType.index:
        #Assign P vector
        # Remove old data
        for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaPgini = generator.CreateObject("ChaVec", "pgini")
        chaPgini.SetAttribute("scale", timescale)
        chaPgini.SetAttribute("vector", dfGenP.loc[generator_name].to_list())
        chaPgini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPgini} for {generator}")
        
        #Create station controller for generator
        stationcontroller = generator.GetParent().CreateObject("ElmStactrl", generator_name + "_SC.ElmStactrl") 
        # Remove old data
        for elmStacontOld in generator.GetContents("c_pstac*.ElmStactrl"): elmStacontOld.Delete() 
        # Assign controller to generator
        generator.SetAttribute("c_pstac", stationcontroller)
        # Assign control nodes (generator terminal for voltage and cub for reactive power Q) 
        stationcontroller.SetAttribute("rembar", generator.GetAttribute("bus1").GetAttribute("cterm"))
        stationcontroller.SetAttribute("p_cub", generator.GetAttribute("bus1"))
        app.PrintPlain(f"Created and assigned {stationcontroller} for {generator}")
        
        # Characteristic for Type
        # Remove old data
        for chaOld in stationcontroller.GetContents("i_ctrl*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaSC_Type = stationcontroller.CreateObject("ChaVec", "i_ctrl") #attribute name, has to be the name of the ChaRef to make the link
        chaSC_Type.SetAttribute("scale", timescale)
        chaSC_Type.SetAttribute("vector", dfGenNodeType.loc[generator_name].to_list())
        chaSC_Type.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned type characteristic {chaSC_Type}")
        
        # Create characteritic for voltage U
        # Remove old data
        for chaOld in stationcontroller.GetContents("usetp*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaSC_U = stationcontroller.CreateObject("ChaVec", "usetp") #attribute name, has to be the name of the ChaRef to make the link
        chaSC_U.SetAttribute("scale", timescale)
        chaSC_U.SetAttribute("vector", dfGenU.loc[generator_name].to_list())
        chaSC_U.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned voltage U characteristic {chaSC_U}")
        
        # Characteristic for Q
        # Remove old data
        for chaOld in stationcontroller.GetContents("qsetp*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaSC_Q = stationcontroller.CreateObject("ChaVec", "qsetp") #attribute name, has to be the name of the ChaRef to make the link
        chaSC_Q.SetAttribute("scale", timescale)
        chaSC_Q.SetAttribute("vector", dfGenQ.loc[generator_name].to_list())
        chaSC_Q.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned reactive power Q characteristic {chaSC_Q}")
        
for load in app.GetCalcRelevantObjects("*.ElmLod"):
    load_name = load.GetAttribute("loc_name")
    if load_name in dfLoadP.index:
        #Assign P vector
        # Remove old data
        for chaOld in load.GetContents("plini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaPlini = load.CreateObject("ChaVec", "plini")
        chaPlini.SetAttribute("scale", timescale)
        chaPlini.SetAttribute("vector", dfLoadP.loc[load_name].to_list())
        chaPlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaPlini} for {load}")
        
        #Assign Q vector
        # Remove old data
        for chaOld in load.GetContents("qlini*.ChaVec"): chaOld.Delete() 
        # Assign controller to generator
        chaQlini = load.CreateObject("ChaVec", "qlini")
        chaQlini.SetAttribute("scale", timescale)
        chaQlini.SetAttribute("vector", dfLoadQ.loc[load_name].to_list())
        chaQlini.SetAttribute("usage", 2)
        app.PrintPlain(f"Created and assigned {chaQlini} for {load}")
    

#################### IZPIS URE ################# KONEC #############################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
# else: print("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
