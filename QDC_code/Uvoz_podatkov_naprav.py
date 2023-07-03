# -*- coding: utf-8 -*-
"""
Created on Wed Jun 28 10:32:33 2023

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
use_powerfactory = True

#Parametri za izračun jalovih moči za gen, load, vac..... načeloma če delamo DC loadflow ni važno
# Za AC loadflow je treba porihtat oz najt neke boljše načine dodeljevanja jalovih. 

#Imena uvoženih datotek, glej da se sklada z tistim kar nardi skripta za pretvorbo excel->csv
stringLineStateFile = "line_state.csv"
stringTransformerStateFile = "transformer_state.csv"
stringTransformerStepFile = "transformer_step.csv"
stringBreakerStateFile = "circbreakter_state.csv"
stringSwitchStateFile = "switch_state.csv"
##########################################################################################################
    
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

#Open excel files
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
app.PrintPlain("Izberi vhodno mapo")
dirFolder = filedialog.askdirectory()
# Daj na true ce mamo .csv datoteke

df_line_state = pd.read_csv(os.path.join(dirFolder, stringLineStateFile), index_col = 0)
df_transformer_state = pd.read_csv(os.path.join(dirFolder, stringTransformerStateFile), index_col = 0)
df_transformer_step = pd.read_csv(os.path.join(dirFolder, stringTransformerStepFile), index_col = 0)
df_circbreakter_state = pd.read_csv(os.path.join(dirFolder, stringBreakerStateFile), index_col = 0)
df_switch_state = pd.read_csv(os.path.join(dirFolder, stringSwitchStateFile), index_col = 0)

df_line_state = df_line_state.replace({0:1, 1:0})
app.PrintPlain(df_line_state)

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
    
done_list = []

#Karakteristike daljnovodov
i=0

for line in app.GetCalcRelevantObjects("*.ElmLne"):
    line_name = line.GetAttribute("loc_name")
    if line_name in df_line_state.index:
        app.PrintPlain(f"DV {line}")
        try:
            # Remove old data
            for chaOld in line.GetContents("outserv*.ChaVec"): chaOld.Delete() 
            # Assign controller to switch
            chaState = line.CreateObject("ChaVec", "outserv")
            chaState.SetAttribute("scale", timescale)
            chaState.SetAttribute("vector", df_line_state.loc[line_name].to_list())
            chaState.SetAttribute("usage", 2)
            app.PrintPlain(f"Nastavil podatke {chaState} za {line}")
        except:
            app.PrintWarn(f"Napaka nastavljanja podatkov za {line}, preveri vhodne datoteke!")
        done_list.append(line_name)
    else:
        app.PrintWarn(f"{line} je brez vhodnih podatkov...")
    i+=1

app.PrintPlain(i)
        
#Karakteristike trafotov

#Karakteristike breakerjev

#Karakteristike switchev ()

    

# for switch in app.GetCalcRelevantObjects("*.ElmCoup"):
#     switch_name = switch.GetAttribute("loc_name")
#     if switch_name in df_switch_state.index:
#         app.PrintPlain(f"Stikalo {switch}")
#         try:
#             # Remove old data
#             for chaOld in switch.GetContents("on_off*.ChaVec"): chaOld.Delete() 
#             # Assign controller to switch
#             chaState = switch.CreateObject("ChaVec", "on_off")
#             chaState.SetAttribute("scale", timescale)
#             chaState.SetAttribute("vector", df_switch_state.loc[switch_name].to_list())
#             chaState.SetAttribute("usage", 2)
#             app.PrintPlain(f"Nastavil podatke {chaState} za {switch}")
#         except:
#             app.PrintWarn(f"Napaka nastavljanja podatkov za {switch}, preveri vhodne datoteke!")
#         done_list.append(switch_name)
#     else:
#         app.PrintWarn(f"{switch} je brez vhodnih podatkov...")

#################### IZPIS URE ################# KONEC #############################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
if use_powerfactory: app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')
else: print("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')