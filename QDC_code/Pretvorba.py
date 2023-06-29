# -*- coding: utf-8 -*-
"""
Created on Wed Jun 28 12:47:45 2023

@author: slluka
"""

import os
import pandas as pd
from tkinter import Tk
from tkinter import filedialog

# file_path = "C:/Users/slluka/Documents/Podatki Advance/202201_SN_NDT_ZDB/20220101_0000_ABB_SCADA.zdb"

# with open(file_path, "r") as file:
#     # Read the entire contents of the file
#     contents = file.read()
#     print(contents)

window_select = False
if window_select:
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    #Open excel files
    print("Izberi vhodno mapo. lahko se pojavi v ozadju (za oknom Spyder-ja) in je v tem primeru potrebno pomanjsati!")
    dirFolder = filedialog.askdirectory()
else:
    dirFolder = "C:/Users/slluka/Documents/Podatki Advance/202201_SN_NDT_ZDB"

file_list = []
file_contents = []
df_line_state = pd.DataFrame()
df_transformer_state = pd.DataFrame() 
df_transformer_step = pd.DataFrame()
df_circbreakter_state = pd.DataFrame()
df_switch_state = pd.DataFrame()

i=1
for root, dirs, files in os.walk(dirFolder):
    for file in files:
        if file.endswith(".zdb"):
            time = file.split("_")[1]
            if time[-2:] == "00":
                #Ce so minute 00 torej vsaka polna ura
                hour = int(time[:2])
                # print(f"first part {first_part}")
                file_path = os.path.join(root, file)
                file_list.append(file_path)
                with open(file_path, "r") as file:
                    data = file.read()
                    file_contents.append(data)
                    #Split rows
                    rows = data.split('\n')
                    for row in rows[:-1]:
                        #Split columns ;
                        values = row.split(";")
                        if values[1] == "LINE":
                            #DV
                            df_line_state.at[values[2], i] = values[3]
                        elif values[1] == "TRANSFORMER":
                            #Trafo
                            df_transformer_state.at[values[2], i] = values[3]
                            df_transformer_step.at[values[2], i] = values[5]
                        elif values[1] == "CIRC_BREAKER":
                            #Odklopnik
                            df_circbreakter_state.at[values[2], i] = values[3]
                        elif values[1] == "DISCSWITCH":
                            #Stikalo
                            df_switch_state.at[values[2], i] = values[3]
                    i+=1
            # file_loc = os.path.join(root, file)
            print(f"Datoteka {file}")
            
df_line_state.to_csv(os.path.join(dirFolder, "line_state.csv"))
df_transformer_state.to_csv(os.path.join(dirFolder, "transformer_state.csv"))
df_transformer_step.to_csv(os.path.join(dirFolder, "transformer_step.csv"))
df_circbreakter_state.to_csv(os.path.join(dirFolder, "circbreakter_state.csv"))
df_switch_state.to_csv(os.path.join(dirFolder, "switch_state.csv"))


# i=1
# for data in file_contents:
#     #Split rows
#     rows = data.split('\n')
#     for row in rows[:-1]:
#         #Split columns ;
#         values = row.split(";")
#         if values[1] == "LINE":
#             #DV
#             df_line_state.at[values[2], i] = values[3]
#         elif values[1] == "TRANSFORMER":
#             #Trafo
#             df_transformer_state.at[values[2], i] = values[3]
#             df_transformer_step.at[values[2], i] = values[5]
#         elif values[1] == "CIRC_BREAKER":
#             #Odklopnik
#             df_circbreakter_state.at[values[2], i] = values[3]
#         elif values[1] == "DISCSWITCH":
#             #Stikalo
#             df_switch_state.at[values[2], i] = values[3]
#     i+=1
        
# rows = file_contents[0].split('\n')
# # Iterate over each row
# for row in rows:
#     print(row)        

print("Konec pretvorbe datotek!")
