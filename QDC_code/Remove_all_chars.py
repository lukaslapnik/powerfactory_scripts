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
# Parametri za izračun jalovih moči za gen, load, vac..... načeloma če delamo DC loadflow ni važno
# Za AC loadflow je treba porihtat oz najt neke boljše načine dodeljevanja jalovih.
# spreminjaj_jalovo = False  # Ali skripta sploh spreminja parametre proizvodnje/porabe jalove moči. False - jalova enaka, True - jalovo spreminja
# izhodiscni_cosfi = True     # Ce je true, bo cosfi enak kot v izhodiscnem modelu, sicer vzame vrednosti definirane spodaj (razmerje med Q in P)

#Namesto cosfi se vnese razmerje PQ_ratio = tan(acos(cosfi(0.xx)))
#Pri cosfi 0.98 ~ 0.2
#Pri cosfi 0.97 ~ 0.25
#Pri cosfi 0.96 ~ 0.3
# generator_PQ_ratio = 0.25 #Delez jalove
# load_PQ_ratio = 0.25
# voltagesource_PQ_ratio = 0

#Izkoristek omrezja (izgube)

#   ['UKNI','UK00','UA02','UA01','TR00','TN00','SK00','SI00','SE04','SE03','SE02','SE01','SA00','RU00',
#   'RS00','RO00','PT00','PS00','PL00','NSW0','NOS0','NON1','NOM1','NL00','MT00','MK00','ME00','MD00',
#   'MA00','LY00','LV00','LUV1','LUG1','LUF1','LUB1','LT00','ITSI','ITSA','ITS1','ITN1','ITCS','ITCN',
#   'ITCA','IS00','IL00','IE00','HU00','HR00','GR03','GR00','FR15','FR00','FI00','ES00','ELES Interconnectios',
#   'EG00','EE00','DZ00','DKW1','DKKF','DEKF','DE00','CZ00','CY00','CH00','BG00','BE00','BA00','AT00','AL00']

#Drzave/sistemi, ki jim spreminjamo parametre. Vnesi tako kot je v market datoteki ali v powerfactory modelu
# sistemi_spreminjanje_parametrov = ['SI00','ITN1','HU00','HR00','ELES Interconnectios']

##########################################################################################################################################################################
#############################################################################   PARAMETRI   ##############################################################################
##########################################################################################################################################################################


start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
# else: print("Pričetek izvajanja programa ob " + str(start_time) + ".")

app.PrintPlain("Start removing chars!")

#Se za gen
for generator in app.GetCalcRelevantObjects("*.ElmSym"):
    for chaOld in generator.GetContents("pgini*.ChaVec"): chaOld.Delete() 
    for chaOld in generator.GetContents("qgini*.ChaVec"): chaOld.Delete()
    app.PrintPlain(f"Removed data for {generator}")

    
for load in app.GetCalcRelevantObjects("*.ElmLod"):
    for chaOld in load.GetContents("plini*.ChaVec"): chaOld.Delete() 
    for chaOld in load.GetContents("qlini*.ChaVec"): chaOld.Delete() 
    app.PrintPlain(f"Removed data for {load}")
    
for voltagesource in app.GetCalcRelevantObjects("*.ElmVac"):
    for chaOld in voltagesource.GetContents("Pgen*.ChaVec"): chaOld.Delete() 
    for chaOld in voltagesource.GetContents("Qgen*.ChaVec"): chaOld.Delete() 
    app.PrintPlain(f"Removed data for {voltagesource}")

#################################################################################### MAIN #######################################################

#################### IZPIS URE #################

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')