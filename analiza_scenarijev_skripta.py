# -*- coding: utf-8 -*-
"""
Created on Thu Jul  6 10:21:31 2023

@author: SSIMON
"""

import os
import datetime
from tkinter import Tk
from tkinter import filedialog

import powerfactory as pf    
app = pf.GetApplication()
app.ClearOutputWindow()
active_project = app.GetActiveProject()
active_study_case = app.GetActiveStudyCase()

###################################Izpis start cajta skripte##############################################
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
app.PrintPlain("Pričetek izvajanja programa ob " + str(start_time) + ".")
##########################################################################################################

#Scenariji v projektu
scenario_folder = app.GetProjectFolder("scen")
scenarios = scenario_folder.GetContents()
#Contingency analiza
contanlysis = app.GetFromStudyCase("ComSimoutage")

app.PrintPlain("Izberi mapo kamor izvozimo rezultate")
# Open excel files
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
export_folder = filedialog.askdirectory()
simulation_fail_list = []

if export_folder:  
    app.PrintPlain(f"Izbrana mapa {export_folder}")
    for scenario in scenarios:
        scenario.Activate()
        converged = contanlysis.Execute()
        if converged == 0:
            # Contingnency analiza uspešna, izvozi rezultate
            # results = app.GetActiveStudyCase("Contingency Analysis AC.ElmRes")
            file_name = active_project.loc_name + "_" + active_study_case.loc_name + "_" + scenario.loc_name + ".csv"
            export_location = os.path.join(export_folder,file_name)
            contingency_report = app.GetFromStudyCase("ComRes")
            contingency_report.iopt_exp = 6
            contingency_report.f_name = export_location
            contingency_report.Execute()
            app.PrintPlain(f"Izvoženi rezultati contingency analize v {export_location}")
            # Clear contingency results data
            # contingency_results = app.GetFromStudyCase("Contingency Analysis AC.ElmRes")
            # contingency_results.bClear()
        elif converged == 1:
            #Neuspešna
            app.PrintWarn("Contingency analiza ni bila uspešna")
            simulation_fail_list.append(active_project.loc_name + "_" + active_study_case.loc_name + "_" + scenario.loc_name)
            
    fail_list_location = os.path.join(export_folder,"fail_list.txt")
    with open(fail_list_location, 'w') as outfile:
      outfile.write('\n'.join(str(i) for i in simulation_fail_list))

else:
    app.PrintError("Izbrana ni bila nobena mapa, zaključujem")
    
################################################################# KONEC #################################################
end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintPlain("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')