# -*- coding: utf-8 -*-
"""
Created on Fri Mar 31 11:14:27 2023

@author: slluka
"""

import os
import pandas as pd
from tkinter import Tk
from tkinter import filedialog

#Open excel files
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
print("Izberi vhodno mapo. lahko se pojavi v ozadju (za oknom Spyder-ja) in je v tem primeru potrebno pomanjsati!")
dirFolder = filedialog.askdirectory()
file_gen_list = []
file_load_list = []
for root, dirs, files in os.walk(dirFolder):
    for file in files:
        if file.endswith(".xlsx"):
            file_loc = os.path.join(root, file)
            print(f"Uvozena datoteka {file}")
            file_excel = pd.ExcelFile(file_loc)
            file_sheets = file_excel.sheet_names
            dfData = pd.DataFrame()
            dfData = file_excel.parse(file_sheets[0], index_col = 0)
            csv_loc = os.path.splitext(file_excel)[0] + ".csv"
            dfData.to_csv(csv_loc)
            print(f"Izvozena datoteka {csv_loc}")
print("Konec pretvorbe datotek!")