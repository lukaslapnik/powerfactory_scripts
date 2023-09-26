# -*- coding: utf-8 -*-
"""
Created on Wed Sep 4 12:54:52 2023

@author: SLUKA
"""
# Uvoz knjižnic in modulov
import pandas as pd
import datetime
import sys
import math
import os

# Append poti do powerfactory modula in uvoz.
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.9") 
import powerfactory as pf

# Print časa začetka programa
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
print("Pričetek izvajanja programa ob " + str(start_time) + ".")

# Inicializacija/dobitev objekta app
app = pf.GetApplication()

# Aktivacija projekta. To se more ujemat z že obstoječim projektom v powerfactory
ime_projekta = "LV Distribution Network"
app.ActivateProject(ime_projekta)

# Pridobitev podatkov projekta, uporabnika, study case, itd.
user = app.GetCurrentUser()
prj = app.GetActiveProject()
activestudycase = app.GetActiveStudyCase()

#Izpis teh podatkov v okno spyderja/console
print("Uporabnik: " + str(user))
print("Projekt: " + str(prj))
print("StudyCase: " + str(activestudycase))

######################################################################################################
########################################### GLAVNI PROGRAM ###########################################
######################################################################################################

for line in app.GetCalcRelevantObjects("*.ElmLne"):
    print(f"Vod: {line.loc_name}")

# TU DODAŠ KODO ZA INTERAKCIJO Z MODELI, IZVEDBO ANALIZ ITD...

######################################################################################################
##################################### KONEC GLAVNI PROGRAM ###########################################
######################################################################################################

# Izpis časa konca skripte in potrebnega skupnega časa za izvedbo
end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
print("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')