# -*- coding: utf-8 -*-
"""
Created on Sat May 28 08:49:17 2022

@author: lukc

some poorly written code, collecting result files and creating a final report. 
"""
import pandas as pd
import datetime
import sys
import powerfactory as pf
import math
import numpy
import os
from os import listdir
import glob
import time
import xlsxwriter
from datetime import datetime as dt
from datetime import timedelta as td

app = pf.GetApplication()
user = app.GetCurrentUser()
ldf = app.GetFromStudyCase("ComLdf")
app.ClearOutputWindow()

#Izpis start cajta
start_time = datetime.datetime.now().time().strftime('%H:%M:%S')
print("Pričetek izvajanja programa ob " + str(start_time) + ".")
app.PrintInfo("Pričetek izvajanja programa ob " + str(start_time) + ".")
################
    
#Naredimo dataframe za sumo pmax generatorjev po državah in energetnih 
#Zacasno hardcoded da ne importamo excel datotek 
hmdindex = ['AL00', 'AT00', 'BA00', 'BE00', 'BG00', 'CH00', 'CY00', 'CZ00', 'DE00', 'DEKF', 'DKE1', 'DKKF', 'DKW1', 'DZ00', 'EE00', 'EG00', 'ES00', 'FI00', 'FR00', 'FR15', 'GR00', 'GR03', 'HR00', 'HU00', 'IE00', 'IL00', 'IS00', 'ITCA', 'ITCN', 'ITCS', 'ITN1', 'ITS1', 'ITSA', 'ITSI', 'LT00', 'LUB1', 'LUF1', 'LUG1', 'LUV1', 'LV00', 'LY00', 'MA00', 'MD00', 'ME00', 'MK00', 'MT00', 'NL00', 'NOM1', 'NON1', 'NOS0', 'NSW0', 'PL00', 'PS00', 'PT00', 'RO00', 'RS00', 'SE01', 'SE02', 'SE03', 'SE04', 'SI00', 'SK00', 'TN00', 'TR00', 'UA01', 'UA02', 'UK00', 'UKNI']
hmdcolumns = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '-28', '29', '-29', '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '44', '-44', '45', '-45', '46', 'DSR', 'LOAD', 'Balance', 'Dump', 'ENS', 'Mgl Cost']
#Konec zacasnega hardcoda, potem se lahko to vzame iz antares excela
#tu naredimo dataframe ki se uporablja za sum instaliranih moci

#Lista izbranih vozlisc zacasno hardcoded
#izbrana_vozlisca = ['load1', 'load2', 'load3', 'generator1', 'generator2', 'generator3']
#izbrana_vozlisca_file_excel = r"S:\SlapnikL_Mag\Excel_datoteke\Izbrana_vozlisca_RN2023-2032 25-5-2022_shorter.xlsx"
#antares_file_excel = r"S:\SlapnikL_Mag\Excel_datoteke\MMStandardOutputFile_NT2030_ANTARES_CY2009_shorter.xlsx"
#antares_file_excel = glob.glob("Vhodni podatki\Market datoteka\*.xlsx")
#antares_file_excel = os.getcwd() + antares_file_excel[0] 
#izbrana_vozlisca_file_excel = glob.glob("Vhodni podatki\Izbrana vozlisca\*.xlsx")
#izbrana_vozlisca_file_excel = os.getcwd() + izbrana_vozlisca_file_excel[0]

debug_all_info = False

def importParameters():
    #Uvozi datoteko z parametri
    #app.PrintInfo("Uvazanje parametrov iz excel datoteke Parametri.xlsx")
    file_input_parameters = os.getcwd() + r'\Parametri.xlsx'
    #app.PrintInfo(file_input_parameters)
    df_input_parameters = pd.read_excel(file_input_parameters, sheet_name = 'Parametri', index_col = 0)
    #app.PrintInfo(df_input_parameters)
    df_input_countries = pd.read_excel(file_input_parameters, sheet_name = 'Drzave', index_col = 0, header = 0)
    df_input_hours = pd.read_excel(file_input_parameters, sheet_name = 'Ure', index_col = 0, header = 0)
    app.PrintInfo(df_input_countries)
    grids_to_set_pf = []
    grids_to_write_results = []
    for country in df_input_countries.index:
        if df_input_countries.at[country, 'Nastavljanje moci'] == 'DA':
            grids_to_set_pf.append(country)
        if df_input_countries.at[country, 'Nastavljanje moci'] == 'DA':
            grids_to_write_results.append(country)
    #app.PrintInfo(grids_to_set_pf + ', ' + grids_to_write_results)
    
    #Zacasno hardcoded da vzame original cosfi za generatorje in bremena
    global PFoldcosfi
    PFoldcosfi = True
    
    #Mogoce cekiramo ce so vrednosti med 0-1 in float format in vzamemo default če niso
    default_cosfi_gen = df_input_parameters.at['cosfigen', 'Vrednost']
    default_cosfi_load = df_input_parameters.at['cosfiload', 'Vrednost']
    default_cosfi_vac = df_input_parameters.at['cosfirobna', 'Vrednost']
    global default_izgube
    default_izgube = df_input_parameters.at['izgubeomr', 'Vrednost']
    #Iz cosfi izracunamo razmerje PQ
    global default_PQrat_gen
    default_PQrat_gen = math.tan(math.acos(default_cosfi_gen))
    global default_PQrat_vac
    default_PQrat_vac = math.tan(math.acos(default_cosfi_vac))
    global default_PQrat_load
    default_PQrat_load = math.tan(math.acos(default_cosfi_load))
    #Ker v antares load uposteva tudi izgube daljnovodov, so bremena v modelu za izkoristek (izgube na daljnovodu) manjsa ker so modelirani daljnovodi posebej
    #Bremena modeliramo 0.02 manj oz faktor 0.98
    
    return grids_to_set_pf, grids_to_write_results

def importCSVFiles():
    app.PrintInfo("Funkcija za uvazanje CSV")
    #Funkcija za uvoz podatkov iz CSV datotek. Ce CSV datotek ni, jih naredi iz excel datotek
    #Zaenkrat se rocno dolocamo ime datoteke in sheetov(listov)
    #input_data_folder = r"S:\SlapnikL_Mag\Vecscenarijska_Analiza\Vhodni podatki"
    input_data_folder_market = os.getcwd() + r'/Vhodni podatki/Market datoteka'
    input_data_folder_izbrana_vozlisca = os.getcwd() + r'/Vhodni podatki/Izbrana vozlisca'
    
    #Find excel and csv files in folder
    #TO-DO AVTOMATSKO NAJDI ANTARES IN IZBRANA VOZLISCA FILE
    antares_file_excel = glob.glob("Vhodni podatki\Market datoteka\*.xlsx")
    #excel_fajli_market_datoteka = os.getcwd() + excel_fajli_market_datoteka[0]
    izbrana_vozlisca_file_excel = glob.glob("Vhodni podatki\Izbrana vozlisca\*.xlsx")
    #izbrana_vozlisca_file_excel = os.getcwd() + izbrana_vozlisca_file_excel[0]    
    
    antares_file_excel = antares_file_excel[0]
    antares_file_excel = antares_file_excel.replace(".xlsx","",1)
    antares_file_excel = antares_file_excel.replace("Vhodni podatki\\Market datoteka\\","",1)

    izbrana_vozlisca_file_excel = izbrana_vozlisca_file_excel[0]
    izbrana_vozlisca_file_excel = izbrana_vozlisca_file_excel.replace(".xlsx","",1)
    izbrana_vozlisca_file_excel = izbrana_vozlisca_file_excel.replace("Vhodni podatki\\Izbrana vozlisca\\","",1)
    #print(izbrana_vozlisca_file_excel)
    
    #Antares and robna vozlisca sheet names
    antares_sheet_list = ['Hourly Market Data', 'Crossborder exchanges']
    robna_vozlisca_sheet_list = ['P', 'Q', 'Robna vozlisca']
    
    #Najdemo cvs in excel datoteke v mapi vhodnih podatkov
    all_files_market = listdir(input_data_folder_market)
    data_folder_csv_files_market = list(filter(lambda f: f.endswith('.csv'), all_files_market))
    data_folder_xlsx_files_market = list(filter(lambda f: f.endswith('.xlsx'), all_files_market))
    all_files_izbrana_vozlisca = listdir(input_data_folder_izbrana_vozlisca)
    data_folder_csv_files_izbrana_vozlisca = list(filter(lambda f: f.endswith('.csv'), all_files_izbrana_vozlisca))
    data_folder_xlsx_files_izbrana_vozlisca = list(filter(lambda f: f.endswith('.xlsx'), all_files_izbrana_vozlisca))
    
    #Cekiraj CSV-je za antares datoteko
    for sheet in antares_sheet_list:
        end_file_name_type = antares_file_excel + "_" + sheet + ".csv"
        end_file_path = os.path.join(input_data_folder_market, end_file_name_type)
        # print(end_file_name_type)
        # print(data_folder_csv_files)
        # Cekiramo ce so te datoteke v direktoriji, sicer jih nardimo
        if end_file_name_type not in data_folder_csv_files_market:
            app.PrintInfo("File ne obstaja, izdelujem csv")
            read_file_add_type = antares_file_excel + ".xlsx"
            read_file_path = os.path.join(input_data_folder_market, read_file_add_type)
            app.PrintInfo(":" + str(read_file_path))
            app.PrintInfo("Sheet name: " + str(sheet))
            read_file = pd.read_excel(read_file_path, sheet_name=sheet, index_col=0)
            read_file.to_csv(end_file_path)
        # Ce datoteke obstajajo beremo
        if sheet == "Hourly Market Data":
            app.PrintInfo("Berem ANTARES Hourly Market Data")
            df_hourly_market_data = pd.read_csv(end_file_path, index_col = [0], skiprows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
        if sheet == "Crossborder exchanges":
            app.PrintInfo("Berem ANTARES Crossborder exchanges")
            df_crossborder_exchanges = pd.read_csv(end_file_path, index_col = [0], skiprows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9])

    # Cekiraj csv-je za robna vozlisca
    for sheet in robna_vozlisca_sheet_list:
        end_file_name_type = izbrana_vozlisca_file_excel + "_" + sheet + ".csv"
        end_file_path = os.path.join(input_data_folder_izbrana_vozlisca, end_file_name_type)
        # print(end_file_name_type)
        # print(data_folder_csv_files)
        # Cekiramo ce so te datoteke v direktoriji, sicer jih nardimo
        if end_file_name_type not in data_folder_csv_files_izbrana_vozlisca:
            app.PrintInfo("File ne obstaja, izdelujem csv")
            read_file_add_type = izbrana_vozlisca_file_excel + ".xlsx"
            read_file_path = os.path.join(input_data_folder_izbrana_vozlisca, read_file_add_type)
            app.PrintInfo(":" + str(read_file_path))
            app.PrintInfo("Sheet name: " + str(sheet))
            read_file = pd.read_excel(read_file_path, sheet_name=sheet, index_col = 0)
            read_file.to_csv(end_file_path)
        #Ce datoteke obstajajo beremo
        if sheet == "P":
            app.PrintInfo("Berem izbrana vozlisca P")
            df_izbrana_vozlisca_p = pd.read_csv(end_file_path, index_col = [0], skiprows=[1, 3, 4, 5], header = [0])
        if sheet == "Q":
            app.PrintInfo("Berem izbrana vozlisca Q")
            df_izbrana_vozlisca_q = pd.read_csv(end_file_path, index_col = [0], skiprows=[1, 3, 4, 5], header = [0])
        if sheet == "Robna vozlisca":
            app.PrintInfo("Berem izbrana vozlisca Robna vozlisca")
            df_robna_vozlisca = pd.read_csv(end_file_path, index_col=[4], header=0)

    return df_hourly_market_data, df_crossborder_exchanges, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, df_robna_vozlisca

def calcGenRatios(generators, df_izbrana_vozlisca):
    app.PrintInfo("Racunanje sum generatorjev po drzavah in energentih")
    generator_list = []
    generator_ratios = {}
    generator_sum = {}
    generator_p_final = {}
    generator_grid_energent = {}
    
    for generator in generators:
        generator_name = generator.loc_name
        if generator_name not in df_izbrana_vozlisca.columns:
            #V listi so generatorji ki jim bomo racunali razmerja
            generator_list.append(generator_name)
            generator_grid = generator.cpGrid.loc_name
            try:
                #V PF imajo generatorji desc/energent pod virtual powerplant na evropskem modelu
                generator_energent = str(''.join(generator.pBMU.desc))
            except:
                #V sloveniji je to zapisano direktno v desc, ne delamo virtualnih powerplantov
                generator_energent = str(''.join(generator.desc))
            #Naredimo po DRZAVA_ENERGENT tako kot je v market datoteki (npr. SI00_26 - to so hidroelektrarne ROR)
            generator_grid_energent[generator_name] = generator_grid + "_" + generator_energent
            #Za proizvodno enoto dobimo koliko je pmax
            generator_p = generator.P_max
            #Ce je pmax proizvodne enote enak 0 vzamemo pmin_ucpu, ce je tudi to 0 vzamemo default 2MW 
            if generator_p == 0:
                generator_p = abs(generator.Pmin_ucPU) 
            if generator_p == 0:
                generator_p = 2.0
            #Potem zapisemo se to v dictionary za potem 
            generator_p_final[generator_name] = generator_p
            #Racunamo sum
            generator_sum[generator_grid_energent[generator_name]] = 0
            
    for generator in generators:
        generator_name = generator.loc_name
        if generator_name in generator_list:
            generator_sum[generator_grid_energent[generator_name]] += generator_p_final[generator_name]   
            
    app.PrintInfo("Racunanje razmerja generatorjev") 
    
    for generator in generators:
        generator_name = generator.loc_name
        if generator_name in generator_list:
            generator_ratios[generator_name] = generator_p_final[generator_name] / generator_sum[generator_grid_energent[generator_name]]
                             
    return generator_ratios

def calcLoadRatios(loads, df_izbrana_vozlisca):
    app.PrintInfo("Racunanje sum bremen")
    load_list = []
    load_ratios = {}
    load_sum = {}
    load_p_final = {}
    load_grid = {}
    
    for load in loads:
        load_name = load.loc_name
        if load_name not in df_izbrana_vozlisca.columns:
            #V listi so generatorji ki jim bomo racunali razmerja
            load_list.append(load_name)
            load_grid[load_name] = load.cpGrid.loc_name
            #Za load dobimo koliko je p - plini
            #Potem zapisemo se to v dictionary za potem 
            load_p_final[load_name] = load.plini
            if load_p_final[load_name] == 0.0:
                #ce je slucajno 0 damo vrednost 1MW da se ne pojavi 0/0. To je samo za racunanje razmerja
                load_p_final[load_name] = 1.0
            #Racunamo sum
            load_sum[load_grid[load_name]] = 0
    
    for load in loads:
        load_name = load.loc_name
        if load_name in load_list:
            load_sum[load_grid[load_name]] += load_p_final[load_name]
            
    app.PrintInfo("Racunanje razmerja generatorjev")
    
    for load in loads:
        load_name = load.loc_name
        if load_name in load_list:
            load_ratios[load_name] = load_p_final[load_name] / load_sum[load_grid[load_name]]
                
    # {'8103_P2W_jrecg': 1.0, '8103_PV_ezfrp': 2.093744791190286, '8103_BM_wzrah': 0.7349871499544447, '8103_LW_vdibu': 0.01849206288655599, '8103BS_bbrgq': 2.0937557893494385, '8103_KWK_deuof': 0.2881395250922762,               
    return load_ratios

def calcAndSetGenPower(generators, df_antares, generator_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour):
    app.PrintInfo("Nastavljanje moci generatorjev in zapis v PF")
    print_info = False
    PQrat_old = {}
    current_hour_sum = {}
    #hour = 8756
    
    app.PrintInfo("Racunanje sume izbranih vozlisc")
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] = 0.0   
    #Sestejemo vrednosti v dataframu izbranih vozlisc
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] += float(df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[hour], vozlisce])
        #app.PrintInfo(current_hour_sum)
    #Potem odstevamo iz market datoteke
    for grid_type in current_hour_sum:
        if print_info: app.PrintInfo("Grid in energent: " + str(grid_type))
        if print_info: app.PrintInfo("Vrednost v market datoteki prej: " + str(df_antares.at[hour, grid_type]))
        if print_info: app.PrintInfo("Vrednost za odstet: " + str(current_hour_sum[grid_type]))
        df_antares.at[hour, grid_type] -= current_hour_sum[grid_type]
        #Ce je manjsi kot 0 ga nastavimo na 0
        if df_antares.at[hour, grid_type] < 0.0 : 
            df_antares.at[hour, grid_type] = 0.0
        if print_info: app.PrintInfo("Vrednost v market datoteki potem: " + str(df_antares.at[hour, grid_type]))
    # Za uro je zdej odsteto od market datoteke 
    
    for generator in generators:
        generator_name = generator.loc_name
        generator_grid = generator.cpGrid.loc_name
        #Dobimo decription generatorja, izven slo so virtualne elektrarne, v slo so v desc
        try:
            generator_type = str(''.join(generator.pBMU.desc))
        except:
            generator_type = str(''.join(generator.desc)) 
        generator_grid_type = generator_grid + "_" + generator_type
        
        #Preverimo ce je generator v izbranih vozliscih
        if generator_name in df_izbrana_vozlisca_p.columns:
            if print_info: app.PrintInfo("Generator iz izbranega vozlisca")
            #Pri importu so duplikatom v imenih elektrarn (kot npr pri crpalni avce kjer je lahko generator ali breme) dodani .1 na koncu zato preverimo če obstaja,
            # torej avce crpanje so CHEAVC-GEN in avce proizvodnja CHEAVC-GEN.1. To je zacasno hardcoded
            generator_name_type_neg = generator_name + ".1"
            if generator_name_type_neg in df_izbrana_vozlisca_p.columns:
                #Preverimo ali je večja cifra za proizvodnjo ali porabo in določimo moči na podlagi tega (napjprej za P, potem Q)
                if float(df_izbrana_vozlisca_p.at[hour, generator_name]) > float(df_izbrana_vozlisca_p.at[hour, generator_name_type_neg]):
                    generator.pgini = float(df_izbrana_vozlisca_p.at[hour, generator_name])
                else:
                    generator.pgini = - float(df_izbrana_vozlisca_p.at[hour, generator_name_type_neg])
                    
                if float(df_izbrana_vozlisca_q.at[hour, generator_name]) > float(df_izbrana_vozlisca_q.at[hour, generator_name_type_neg]):
                    generator.qgini = float(df_izbrana_vozlisca_q.at[hour, generator_name])
                else:
                    generator.qgini = - float(df_izbrana_vozlisca_q.at[hour, generator_name_type_neg])
            #Ce ni crpalna, baterija itd. in je samo generator nastavimo direktno
            else:
                generator.pgini = float(df_izbrana_vozlisca_p.at[hour, generator_name])
                generator.qgini = float(df_izbrana_vozlisca_q.at[hour, generator_name])
                
        #Tu je generator iz market datoteke       
        elif generator_grid_type in df_antares.columns:
            #Negativni energent za potrebe baterij in akumulacijskih in power to gas
            #To deluje tako da vrine - pred cifto
            generator_grid_type_neg = generator_grid_type[0:5] + "-" + generator_grid_type[5:]
            if print_info: app.PrintInfo("Generator : " + generator_name + " z grid in energentom: " + generator_grid_type)
            #Izracunamo se tanfi (lazji racun kot cosfi ampak ista zadeva), kar je razmerje med delovno in jalovo
            if generator.pgini == 0.0 or generator.qgini == 0.0:
                PQrat_old[generator_name] = 0.0
            else:
                PQrat_old[generator_name] = generator.qgini/generator.pgini
            if generator_grid_type_neg in df_antares.columns and df_antares.at[hour, generator_grid_type_neg] > df_antares.at[hour, generator_grid_type]:
                #Obstaja negativni tip energenta v antares
                #Negativni tip ima večjo moč
                generator_p = - float(df_antares.at[hour, generator_grid_type_neg] * generator_ratios[generator_name])
                #Ali jalovo proizvaja ali porablja??
                if PFoldcosfi:
                    generator_q = - float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * PQrat_old[generator_name])
                else:
                    generator_q = - float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * default_PQrat_gen)
            else:
                #Ni negativnega energenta ali ima manjšo moč
                generator_p = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name])
                if PFoldcosfi:
                    generator_q = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * PQrat_old[generator_name])
                else:
                    generator_q = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * default_PQrat_gen)
            if print_info: app.PrintInfo(generator_p)
            if print_info: app.PrintInfo(generator_q)
            generator.pgini = generator_p
            generator.qgini = generator_q
        else:
            app.PrintInfo("Generator: " + generator_name + " z grid in energentom: " + generator_grid_type + " ni me dodeljeno nič")
    return

def calcAndSetLoadPower(loads, df_antares, load_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour):
    #Vhod je dataobject bremen, dataframe antaresa za določeno uro, dictionary koeficientov moči bremen, vektor drzav katere nastavljamo in bool ali upoštevamo stare cosfi
    PQrat_old = {}
    print_info=False
    current_hour_sum = {}
    app.PrintInfo("Nastavljanje moci bremen in zapis v PF")
    
    app.PrintInfo("Racunanje sume izbranih vozlisc")
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] = 0.0   
    #Sestejemo vrednosti v dataframu izbranih vozlisc
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] += float(df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[hour], vozlisce])
        #app.PrintInfo(current_hour_sum)
    #Potem odstevamo iz market datoteke
    for grid_type in current_hour_sum:
        if print_info: app.PrintInfo("Grid in energent: " + str(grid_type))
        if print_info: app.PrintInfo("Vrednost v market datoteki prej: " + str(df_antares.at[hour, grid_type]))
        if print_info: app.PrintInfo("Vrednost za odstet: " + str(current_hour_sum[grid_type]))
        df_antares.at[hour, grid_type] -= current_hour_sum[grid_type]
        #Ce je manjsi kot 0 ga nastavimo na 0
        if df_antares.at[hour, grid_type] < 0.0 : 
            df_antares.at[hour, grid_type] = 0.0
        if print_info: app.PrintInfo("Vrednost v market datoteki potem: " + str(df_antares.at[hour, grid_type]))
    # Za uro je zdej odsteto od market datoteke
    
    for load in loads:
        load_name = load.loc_name
        load_grid = load.cpGrid.loc_name
        load_grid_type = load_grid + "_LOAD"
        if load_name in df_izbrana_vozlisca_p.columns:
            #Breme iz izbranih vozlisc
            load.plini = float(df_izbrana_vozlisca_p.at[hour, load_name])
            load.qlini = float(df_izbrana_vozlisca_q.at[hour, load_name])
            if print_info: app.PrintInfo("Load iz izbranega vozlisca: " + load_name + " nastavljen P: " + str(load.plini) + " in Q: " + str(load.qlini))
                
        elif load_grid_type in df_antares.columns:
            if load.plini == 0.0 or load.qlini == 0.0:
                PQrat_old[load_name] = 0.0
            else:
                PQrat_old[load_name] = load.qlini/load.plini
            load.plini = load_ratios[load_name] * df_antares.at[hour, load_grid_type] * (1-default_izgube)
            if PFoldcosfi:
                #ce je bool true vzamemo cosfi enak kot je bil prej v modelu
                load.qlini = load.plini * PQrat_old[load_name]
            else:
                #Drugace nastavimo jalovo po rocno definiranem cosfi
                load.qlini = load.plini * default_PQrat_load
        else:
            if print_info: app.PrintInfo("Breme: " + load_name + " ima drzavo: " + load_grid + " zato je preskocen")
    return

def calcAndSetGenLoadPower(generators, loads, df_antares, generator_ratios, load_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour, grids_to_set):
    app.PrintInfo("Nastavljanje moci generatorjev in zapis v PF")
    print_info = False
    PQrat_old = {}
    current_hour_sum = {}
    #hour = 8756
    #Najprej pomnozimo z (1-default izgube) ker so v market datoteki zajete izgube omrežja v LOAD
    for grid_type in df_antares.columns:
        if "_LOAD" in grid_type:
            df_antares.at[hour, grid_type] = df_antares.at[hour, grid_type] * (1 - default_izgube)
        # Prav tako odstejemo dumped energy od soncne (antares dela tako da dumped energy je energija ki je ni možno izkoristit, torej je proizvodnje preveč in gre stran)
        # Dumped 
        if "_Dump" in grid_type:
            if df_antares.at[hour, grid_type] > 0:
                app.PrintInfo("Imamo dump in je vecji od 0")
                #Ce je dump > 0
                dump_grid = grid_type[0:4]
                grid_solar = dump_grid + "_33"
                grid_wind_offshore = dump_grid + "_32"
                grid_wind_onshore = dump_grid + "_31"
                hour_power_dump = df_antares.at[hour, grid_type]
                hour_power_solar = df_antares.at[hour, grid_solar]
                hour_power_wind_offshore = df_antares.at[hour, grid_wind_offshore]
                hour_power_wind_onshore = df_antares.at[hour, grid_wind_onshore]
                power_sum = hour_power_solar + hour_power_wind_offshore + hour_power_wind_onshore
                if hour_power_solar > 0:
                    ratio_grid_solar = hour_power_solar/power_sum
                else:
                    ratio_grid_solar = 0
                    
                if hour_power_wind_offshore > 0:
                    ratio_grid_wind_offshore = hour_power_wind_offshore/power_sum
                else:
                    ratio_grid_wind_offshore = 0
                    
                if hour_power_wind_onshore > 0:
                    ratio_grid_wind_onshore = hour_power_wind_onshore/power_sum
                else:
                    ratio_grid_wind_onshore = 0
                    
                app.PrintInfo("Solar prej: " + str(df_antares.at[hour, grid_solar]))
                app.PrintInfo("Dumped: " + str(df_antares.at[hour, grid_type]))
                app.PrintInfo("Razmerje: " + str(ratio_grid_solar))
                
                df_antares.at[hour, grid_solar] = hour_power_solar - hour_power_dump * ratio_grid_solar
                if df_antares.at[hour, grid_solar] < 0:
                    df_antares.at[hour, grid_solar]  = 0
                    
                df_antares.at[hour, grid_wind_offshore] = hour_power_wind_offshore - hour_power_dump * ratio_grid_wind_offshore
                if df_antares.at[hour, grid_wind_offshore] < 0:
                    df_antares.at[hour, grid_wind_offshore]  = 0
                    
                df_antares.at[hour, grid_wind_onshore] = hour_power_wind_onshore - hour_power_dump * ratio_grid_wind_onshore
                if df_antares.at[hour, grid_wind_onshore] < 0:
                    df_antares.at[hour, grid_wind_onshore]  = 0
                    
                app.PrintInfo("Solar potem: " + str(df_antares.at[hour, grid_solar]))
    
    #app.PrintInfo("IZRACUNAN LOAD SUM KI GA BI NASTAVU: " + str(df_antares.at[hour, 'SI00_LOAD']))
    
    # Racunamo sum izbranih vozlisc da ga odstejemo od market datoteke
    
    app.PrintInfo("Racunanje sume izbranih vozlisc")
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] = 0.0   
    #Sestejemo vrednosti v dataframu izbranih vozlisc
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] += float(df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[hour], vozlisce])
        #app.PrintInfo(current_hour_sum)
    #Potem odstevamo iz market datoteke
    for grid_type in current_hour_sum:
        if print_info: app.PrintInfo("Grid in energent: " + str(grid_type))
        if print_info: app.PrintInfo("Vrednost v market datoteki prej: " + str(df_antares.at[hour, grid_type]))
        if print_info: app.PrintInfo("Vrednost za odstet: " + str(current_hour_sum[grid_type]))
        df_antares.at[hour, grid_type] -= current_hour_sum[grid_type]
        #Ce je manjsi kot 0 ga nastavimo na 0
        if df_antares.at[hour, grid_type] < 0.0 : 
            df_antares.at[hour, grid_type] = 0.0
        if print_info: app.PrintInfo("Vrednost v market datoteki potem: " + str(df_antares.at[hour, grid_type]))
    # Za uro je zdej odsteto od market datoteke 
    
    
    #Potem dolocamo moc generatorjev
    for generator in generators:
        generator_name = generator.loc_name
        generator_grid = generator.cpGrid.loc_name
        #Dobimo decription generatorja, izven slo so virtualne elektrarne, v slo so v desc
        if generator_grid in grids_to_set:
            try:
                generator_type = str(''.join(generator.pBMU.desc))
            except:
                generator_type = str(''.join(generator.desc)) 
            generator_grid_type = generator_grid + "_" + generator_type
        
            #Preverimo ce je generator v izbranih vozliscih in dolocimo tu
            if generator_name in df_izbrana_vozlisca_p.columns:
                if print_info: app.PrintInfo("Generator iz izbranega vozlisca")
                #Pri importu so duplikatom v imenih elektrarn (kot npr pri crpalni avce kjer je lahko generator ali breme) dodani .1 na koncu zato preverimo če obstaja,
                # torej avce crpanje so CHEAVC-GEN in avce proizvodnja CHEAVC-GEN.1. To je zacasno hardcoded
                # Naredimo se tako da ce crpa v omrezje oddaja jalovo moc, ce pa deluje kot generator iz omrezja jemlje jalovo (to se dela za uravnavanje napetosti)
                generator_name_type_neg = generator_name + ".1" #To je proizvodnja
                if generator_name_type_neg in df_izbrana_vozlisca_p.columns:
                    #Preverimo ali je večja cifra za proizvodnjo ali porabo in določimo moči na podlagi tega (napjprej za P, potem Q)
                    if float(df_izbrana_vozlisca_p.at[hour, generator_name]) > float(df_izbrana_vozlisca_p.at[hour, generator_name_type_neg]):
                        #Tukaj crpa in jemlje delovno
                        generator.pgini = - float(df_izbrana_vozlisca_p.at[hour, generator_name])
                    else:
                        #Tukaj generira delovno
                        generator.pgini = float(df_izbrana_vozlisca_p.at[hour, generator_name_type_neg])
                    
                    if float(df_izbrana_vozlisca_q.at[hour, generator_name]) > float(df_izbrana_vozlisca_q.at[hour, generator_name_type_neg]):
                        #Tu jemlje jalovo
                        generator.qgini = - float(df_izbrana_vozlisca_q.at[hour, generator_name])
                    else:
                        #Tu oddaja jalovo
                        generator.qgini = float(df_izbrana_vozlisca_q.at[hour, generator_name_type_neg])
                #Ce ni crpalna, baterija itd. in je samo generator nastavimo direktno
                else:
                    generator.pgini = float(df_izbrana_vozlisca_p.at[hour, generator_name])
                    generator.qgini = float(df_izbrana_vozlisca_q.at[hour, generator_name])
                
            #Tu je generator iz market datoteke       
            elif generator_grid_type in df_antares.columns:
                #Negativni energent za potrebe baterij in akumulacijskih in power to gas
                #To deluje tako da vrine - pred cifto
                generator_grid_type_neg = generator_grid_type[0:5] + "-" + generator_grid_type[5:]
                if print_info: app.PrintInfo("Generator : " + generator_name + " z grid in energentom: " + generator_grid_type)
                #Izracunamo se tanfi (lazji racun kot cosfi ampak ista zadeva), kar je razmerje med delovno in jalovo
                if generator.pgini == 0.0 or generator.qgini == 0.0:
                    PQrat_old[generator_name] = 0.0
                else:
                    PQrat_old[generator_name] = generator.qgini/generator.pgini
                    
                if generator_grid_type_neg in df_antares.columns and df_antares.at[hour, generator_grid_type_neg] > df_antares.at[hour, generator_grid_type]:
                    #Obstaja negativni tip energenta v antares
                    #Negativni tip ima večjo moč
                    generator_p = - float(df_antares.at[hour, generator_grid_type_neg] * generator_ratios[generator_name])
                    #Ali jalovo proizvaja ali porablja??
                    if PFoldcosfi:
                        generator_q = - float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * PQrat_old[generator_name])
                    else:
                        generator_q = - float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * default_PQrat_gen)
                else:
                    #Ni negativnega energenta ali ima manjšo moč
                    generator_p = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name])
                    if PFoldcosfi:
                        generator_q = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * PQrat_old[generator_name])
                    else:
                        generator_q = float(df_antares.at[hour, generator_grid_type] * generator_ratios[generator_name] * default_PQrat_gen)
                if print_info: app.PrintInfo(generator_p)
                if print_info: app.PrintInfo(generator_q)
                generator.pgini = generator_p
                generator.qgini = generator_q
            else:
                if print_info: app.PrintInfo("Generator: " + generator_name + " z grid in energentom: " + generator_grid_type + " ni me dodeljeno nič")
        else:
            if print_info: app.PrintInfo("Generator je iz grida v katerem ne nastavljamo: " + generator_grid)
    
    #Se za bremena
    for load in loads:
        load_name = load.loc_name
        load_grid = load.cpGrid.loc_name
        if load_grid in grids_to_set:
            load_grid_type = load_grid + "_LOAD"
            if load_name in df_izbrana_vozlisca_p.columns:
                #Breme iz izbranih vozlisc
                load.plini = float(df_izbrana_vozlisca_p.at[hour, load_name])
                load.qlini = float(df_izbrana_vozlisca_q.at[hour, load_name])
                if print_info: app.PrintInfo("Load iz izbranega vozlisca: " + load_name + " nastavljen P: " + str(load.plini) + " in Q: " + str(load.qlini))
                
            elif load_grid_type in df_antares.columns:
                if load.plini == 0.0 or load.qlini == 0.0:
                    PQrat_old[load_name] = 0.0
                else:
                    PQrat_old[load_name] = load.qlini/load.plini
                load.plini = load_ratios[load_name] * df_antares.at[hour, load_grid_type]
                if PFoldcosfi:
                    #ce je bool true vzamemo cosfi enak kot je bil prej v modelu
                    load.qlini = load.plini * PQrat_old[load_name]
                else:
                    #Drugace nastavimo jalovo po rocno definiranem cosfi
                    load.qlini = load.plini * default_PQrat_load
            else:
                if print_info: app.PrintInfo("Breme: " + load_name + " ima drzavo: " + load_grid + " zato je preskocen")
        else:
            if print_info: app.PrintInfo("Load je iz grida v katerem ne nastavljamo: " + generator_grid)
    return

def nastaviRobnavozlisca(voltagesources, df_antares_crossborder, df_robna, hour):
    print_info = False
    #Vhod so "ELMVAC", dataframe crossborder iz antares datoteke, dataframe robnih vozlisc iz excel datoteke in ura simulacije
    if print_info: app.PrintInfo("Nastavljanje robnih vozlisc")
    for voltagesource in voltagesources:
        #Najprej dobimo ime vozlisca (v PF modelirani kot VAC zato je ime voltagesource)
        voltagesource_name = voltagesource.loc_name
        #voltagesource_name = voltagesource
        #print("Vozlisce: " + voltagesource_name)
        #ZACASNO HARDCODED ZA TESTIRANJE
        #Ce je ta v excel datoteki robnih vozlisc gremo dalje
        if voltagesource_name in df_robna.index:
            if debug_all_info: app.PrintInfo("Najdeno robno")
            #Najdemo ime - mejo drzav ki mora bit eneka kot v antares
            border = df_robna.at[voltagesource_name, 'MEJA']
            #print("Meja: " + border)
            #Preverimo ce se border nahaja v antaresu
            if border in df_antares_crossborder.columns:
                totalpower = df_antares_crossborder.at[df_antares_crossborder.index[hour], border]
                ratio = df_robna.at[voltagesource_name, 'DELEZ']
                multiplier = df_robna.at[voltagesource_name, 'POMNOZITI']
                voltagesource_p = totalpower * ratio * multiplier
                #Zakaj točno nastavljamo na int???
                voltagesource.Pgen = voltagesource_p
                voltagesource.Qgen = default_PQrat_vac * voltagesource_p
                if debug_all_info: app.PrintInfo("Vozlisce: " + voltagesource_name + " dodeljena moc: " + str(voltagesource_p))
                # print("Moc v antares: " + str(totalpower))
                # print("Ratio: " + str(ratio))
                # print("Multiplier: " + str(multiplier))
                # app.PrintInfo("Koncna moc: " + str(voltagesource_p))
    return

def odstejSumIzbranaVozlisca(df_antares, df_izbrana_vozlisca_p, hours):
    hours = range(1,873)
    print_info = False
    current_hour_sum = {}
    #Prvi loop nardimo da naredimo vektor in ga zafilamo z 0
    for vozlisce in df_izbrana_vozlisca_p.columns:
        vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
        if 'SI00_' in vozlisce_grid_energent:
            current_hour_sum[vozlisce_grid_energent] = 0.0
    #Potem delamo loope za vsako uro
    for hour in hours:
        #Pocistimo vrednosti
        for grid_type in current_hour_sum:
            current_hour_sum[grid_type] = 0.0
        #Sestejemo vrednosti v dataframu izbranih vozlisc
        for vozlisce in df_izbrana_vozlisca_p.columns:
            vozlisce_grid_energent = df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[0],vozlisce]
            if 'SI00_' in vozlisce_grid_energent:
                current_hour_sum[vozlisce_grid_energent] += float(df_izbrana_vozlisca_p.at[df_izbrana_vozlisca_p.index[hour], vozlisce])
        #app.PrintInfo(current_hour_sum)
        #Potem odstevamo iz market datoteke
        for grid_type in current_hour_sum:
            if print_info: app.PrintInfo("Grid in energent: " + str(grid_type))
            if print_info: app.PrintInfo("Vrednost v market datoteki prej: " + str(df_antares.at[hour, grid_type]))
            if print_info: app.PrintInfo("Vrednost za odstet: " + str(current_hour_sum[grid_type]))
            df_antares.at[hour, grid_type] -= current_hour_sum[grid_type]
            if print_info: app.PrintInfo("Vrednost v market datoteki potem: " + str(df_antares.at[hour, grid_type]))
    
    return df_antares
       
def checkAllGenSlo(generators, grid_to_check):
    power_value_sum = 0
    for generator in generators:
        generator_grid = generator.cpGrid.loc_name
        if generator_grid in grid_to_check:
            power_value_sum += generator.pgini
    return power_value_sum
            
def checkAllLoadSlo(loads, grid_to_check):
    power_value_sum = 0
    for load in loads:
        load_grid = load.cpGrid.loc_name
        if load_grid in grid_to_check:
            power_value_sum += load.plini
    return power_value_sum

############ NEKO PISANJE REZULTATOV

# def izpisRezultatovGeneratorji(df_vrednosti):
#     app.PrintInfo("Izpisovanje rezultatov generatorjev")
#     global df_results_generators
#     df_results_generators = pd.DataFrame()
#     for generator in generators:
#         if generator.cpZone.loc_name in powerfactorygensum.index and str(''.join(generator.pBMU.desc)) in powerfactorygensum.columns:
#             df_results_generators.at[generator.loc_name, 'Drzava'] = generator.cpZone.loc_name
#             df_results_generators.at[generator.loc_name, 'Energent'] = str(''.join(generator.pBMU.desc))
#             df_results_generators.at[generator.loc_name, 'P_max'] = generator.P_max
#             df_results_generators.at[generator.loc_name, 'P_ratio'] = generator_ratio[generator.loc_name]
#             df_results_generators.at[generator.loc_name, 'P_gini'] = generator.pgini
#             df_results_generators.at[generator.loc_name, 'P_sum'] = powerfactorygensum.at[generator.cpZone.loc_name, str(''.join(generator.pBMU.desc))]
#             if ('-' + str(''.join(generator.pBMU.desc))) in df_vrednosti.columns:
#                 #app.PrintInfo("Obstaja negativni generator")
#                 if df_vrednosti.at[generator.cpZone.loc_name, ('-' + str(''.join(generator.pBMU.desc)))] > 0:
#                     #Ce je pozitivna cifra v neg vrednosti energenta zapisemo tega
#                     #app.PrintInfo(generator_energent_neg)
#                     #app.PrintInfo(df_vrednosti.at[generator_zone, generator_energent_neg])
#                     #app.PrintInfo(powerfactorygensum.at[generator_zone,generator_energent_neg])
#                     df_results_generators.at[generator.loc_name, 'P_Antares'] = df_vrednosti.at[generator.cpZone.loc_name, ('-' + str(''.join(generator.pBMU.desc)))]
#                     #app.PrintInfo(generator_p_new)
#                 elif df_vrednosti.at[generator.cpZone.loc_name, ('-' + str(''.join(generator.pBMU.desc)))] == 0:
#                     df_results_generators.at[generator.loc_name, 'P_Antares'] = df_vrednosti.at[generator.cpZone.loc_name, str(''.join(generator.pBMU.desc))]
#             else:
#                 df_results_generators.at[generator.loc_name, 'P_Antares'] = df_vrednosti.at[generator.cpZone.loc_name, str(''.join(generator.pBMU.desc))]

            
#PISANJE REZULTATOV V EXCEL
def shraniVmesneRezultateCsv(lines, transformers, terminals, grids_to_write, hour):
    df_results_loading_hourly = pd.DataFrame(data=None)
    df_results_voltage_hourly = pd.DataFrame(data=None)
    app.PrintInfo("Zapis v csv za trenutno uro")
    #Najprej zapisemo rezultate v dataframe za daljnovode
    for line in lines:
        line_name = line.loc_name
        line_grid = line.cpGrid.loc_name
        if line_grid in grids_to_write:
            if line.HasResults() == 1:
                #Ce je in service
                line_loading = round(line.GetAttribute('c:loading'), 0)
            else:
                # Če je element izklopljen
                line_loading = int(0)
            df_results_loading_hourly.at[line_name, hour] = line_loading
            
    #Nato se rezultate za transformatorje
    for transformer in transformers:
        transformer_name = transformer.loc_name
        transformer_grid = transformer.cpGrid.loc_name
        if transformer_grid in grids_to_write:
            if transformer.HasResults() == 1:
                #Ce je in service
                transformer_loading = round(transformer.GetAttribute('c:loading'), 0)
            else:
                # Če je element izklopljen
                transformer_loading = int(0)
            df_results_loading_hourly.at[transformer_name, hour] = transformer_loading
            
    for terminal in terminals:
        terminal_name = terminal.loc_name
        terminal_grid = terminal.cpGrid.loc_name
        if terminal_grid in grids_to_write:
            if terminal.HasResults() == 1:
                #Ce je in service
                #Napetost je per-unit
                terminal_voltage = round(terminal.GetAttribute('m:u'), 2)
            else:
                # Če je element izklopljen
                terminal_voltage = int(0)
            df_results_voltage_hourly.at[terminal_name, hour] = terminal_voltage
        
    app.PrintInfo("Izpisana " + str(hour) + " ura")
    #app.PrintInfo(df_results_lines)
    loading_file_name = 'Loading_ura_' + str(hour)
    voltage_file_name = 'Voltage_ura_' + str(hour)
    loading_file_path = os.getcwd()  + r'/Vmesni rezutati/Loading/' + loading_file_name +'.csv'
    voltage_file_path = os.getcwd()  + r'/Vmesni rezutati/Voltage/' + voltage_file_name +'.csv'
    df_results_loading_hourly.to_csv(loading_file_path, encoding='utf-8', index=True)
    df_results_voltage_hourly.to_csv(voltage_file_path, encoding='utf-8', index=True)
    return


#Mogoce raje uporabimo xlsxwriter ker ima vec kontrole nad datoteko v excel (barve, grafi, formule itd itd. ) Bo pa treba za podatke ročno loopat verjetno 
#Mogoce pustimo dataframe izpis v posebej datoteki kjer je samo raw data po urah?
#Vse v enem - trafoti + daljnovodi

def writeEndResultsXlsxwriter(lines, transformers, busses, grids_to_write, hours):
    
    print_info = False
    
    #Poberemo datoteke in nardimo velik df vseh ur za daljnovode
    loading_hour_file_name = 'Loading_ura_' + str(hours[0])
    loading_hour_file_path = os.getcwd()  + r'/Vmesni rezutati/Loading/' + loading_hour_file_name +'.csv'
    df_results_loading_hour = pd.read_csv(loading_hour_file_path, index_col = 0)
    
    voltage_hour_file_name = 'Voltage_ura_' + str(hours[0])
    voltage_hour_file_path = os.getcwd()  + r'/Vmesni rezutati/Voltage/' + voltage_hour_file_name +'.csv'
    df_results_voltage_hour = pd.read_csv(voltage_hour_file_path, index_col = 0)
    
    firsthour = True
    for hour in hours:
        if firsthour:
            firsthour = False
        else:
            loading_hour_file_name = 'Loading_ura_' + str(hour)
            voltage_hour_file_name = 'Voltage_ura_' + str(hour)
            loading_hour_file_path = os.getcwd()  + r'/Vmesni rezutati/Loading/' + loading_hour_file_name +'.csv'
            voltage_hour_file_path = os.getcwd()  + r'/Vmesni rezutati/Voltage/' + voltage_hour_file_name +'.csv'
            try:
                df_loading_current_hour = pd.read_csv(loading_hour_file_path, index_col = 0)
                df_voltage_current_hour = pd.read_csv(voltage_hour_file_path, index_col = 0)
                app.PrintInfo(df_loading_current_hour)
                df_results_loading_hour = pd.concat([df_results_loading_hour, df_loading_current_hour], axis=1, join='inner')
                df_results_voltage_hour = pd.concat([df_results_voltage_hour, df_voltage_current_hour], axis=1, join='inner')
            except:
                app.PrintInfo("DATOTEKA NE OBSTAJA, VERJETNO BILA NEKONVERGENCA")
    #Menjamo "out of service" z cifro 0.1
    
    #Shranimo samo podatke
    writerawdftoexcel = True
    if writerawdftoexcel:
        raw_data_loading = os.getcwd()  + r'/Rezultati/Raw data loading.xlsx'
        raw_data_voltage = os.getcwd()  + r'/Rezultati/Raw data voltage.xlsx'
        df_results_loading_hour.to_excel(raw_data_loading, sheet_name = 'RAW DATA')
        df_results_voltage_hour.to_excel(raw_data_voltage, sheet_name = 'RAW DATA')
    
    results_hourly_elementlist = df_results_loading_hour.index
    results_hourly_elementlist_voltage = df_results_voltage_hour.index
    
    #Dobimo 2 najvisji obremenitvi v letu za vse daljnovode in transformatorje
    #mogoce uporabno https://stackoverflow.com/questions/59574855/python-return-second-highest-value-column-name-in-row
    dftemp1 = df_results_loading_hour.mask(df_results_loading_hour == 0)
    #app.PrintInfo("DFTEMP")
    #app.PrintInfo(dftemp1)
    max_min_values = pd.DataFrame()
    max_min_values['max_no1_loading'] = dftemp1.idxmax(axis=1)
    dftemp1_mask = dftemp1.columns.to_numpy() == max_min_values['max_no1_loading'].to_numpy()[:, None]
    dftemp1 = dftemp1.mask(dftemp1_mask)
    max_min_values['max_no2_loading'] = dftemp1.idxmax(axis=1)
    
    #Dobimo se min 2 in max 2 napetosti
    dftemp2 = df_results_voltage_hour.mask(df_results_voltage_hour == 0)
    #app.PrintInfo("DFTEMP")
    #app.PrintInfo(dftemp1)
    max_no = pd.DataFrame()
    max_min_values['max_no1_voltage'] = dftemp2.idxmax(axis=1)
    dftemp2_mask = dftemp2.columns.to_numpy() == max_min_values['max_no1_voltage'].to_numpy()[:, None]
    dftemp2 = dftemp2.mask(dftemp2_mask)
    max_min_values['max_no2_voltage'] = dftemp2.idxmax(axis=1)
    #shranjeni max 2 napetosti zbiralk, zdaj se min 2 napetosti
    max_min_values['min_no1_voltage'] = dftemp2.idxmin(axis=1)
    dftemp2_mask = dftemp2.columns.to_numpy() == max_min_values['min_no1_voltage'].to_numpy()[:, None]
    dftemp2 = dftemp2.mask(dftemp2_mask)
    max_min_values['min_no2_voltage'] = dftemp2.idxmin(axis=1)
    
    #Prazne (nan) vrednosti menjamo z 0   
    max_min_values.fillna('0', inplace = True)
    #m = dftemp1.columns.to_numpy() == df_results_lines_hourly['max_no2'].to_numpy()[:, None]
    #dftemp1 = dftemp1.mask(m)
    #df_results_lines_hourly['max_no3'] = dftemp1.idxmax(axis=1)
    
    #Za obremenitve nardimo matriko z narascanjem po 10%
    #
    percentload = pd.DataFrame()
    for percent10 in range(0,21):
        percent = percent10*10
        dataframe_name = 'nad' + str(percent)
        percentload[dataframe_name]=df_results_loading_hour.iloc[:,1:].ge(float(percent)).sum(axis=1)
    
    app.PrintInfo(percentload)
    
    # Naredimo workbook
    workbook = xlsxwriter.Workbook(os.getcwd() + r'/Rezultati/Rezultati.xlsx')
    
    # Zapisemo sheet 1 dveh najhujsih obremenitev
    worksheet1 = workbook.add_worksheet('LoadFlowNajhujse')
    
    format_header1 = workbook.add_format({'bold': True})
    format_header1.set_align('vcenter')
    format_header1.set_align('center')
    format_header1.set_bg_color('#b9faad')
    format_header1.set_border()
    
    format_header2 = workbook.add_format({'bold': True})
    format_header2.set_align('vcenter')
    format_header2.set_align('center')
    format_header2.set_bg_color('#fcc5c2')
    format_header2.set_border()
    
    format_header3 = workbook.add_format({'bold': True})
    format_header3.set_align('vcenter')
    format_header3.set_align('center')
    format_header3.set_bg_color('#d5fcb8')
    format_header3.set_border()
    
    format_header4 = workbook.add_format({'bold': True})
    format_header4.set_align('vcenter')
    format_header4.set_align('center')
    format_header4.set_bg_color('#90bbf8')
    format_header4.set_border()
    
    format_header5 = workbook.add_format({'bold': True})
    format_header5.set_align('vcenter')
    format_header5.set_align('center')
    format_header5.set_bg_color('#9cf05a')
    format_header5.set_border()
    
    format_header6 = workbook.add_format({'bold': True})
    format_header6.set_align('vcenter')
    format_header6.set_align('center')
    format_header6.set_bg_color('#e6f13b')
    format_header6.set_border()
    
    format_header7 = workbook.add_format({'bold': True})
    format_header7.set_align('vcenter')
    format_header7.set_align('center')
    format_header7.set_bg_color('#f3634f')
    format_header7.set_border()
    
    format_data_lighter = workbook.add_format()
    format_data_lighter.set_bg_color('#f1f1f1')
    
    format_data_darker = workbook.add_format()
    format_data_darker.set_bg_color('#dcdcdc')
    
    worksheet1.set_row(0, 30)
    worksheet1.set_column(0, 2, 20)
    worksheet1.set_column(3, 3, 15)
    worksheet1.set_column(4, 5, 10)
    worksheet1.set_column(6, 7, 13)
    worksheet1.set_column(8, 8, 16)
    worksheet1.set_column(9, 10, 10)
    worksheet1.set_column(11, 12, 13)
    worksheet1.set_column(13, 15, 7)
    worksheet1.set_column(16, 16, 13)
    worksheet1.set_column(17, 18, 7)
    worksheet1.set_column(19, 19, 20)
    worksheet1.write(0, 0, 'PRAVO IME', format_header1)
    worksheet1.write(0, 1, 'ADVANCED IME', format_header1)
    worksheet1.write(0, 2, 'POWERFACTORY IME', format_header1)
    worksheet1.write(0, 3, 'PRVO NAJHUJSE', format_header2)
    worksheet1.write(0, 4, 'URA V LETU', format_header2)
    worksheet1.write(0, 5, 'DATUM', format_header2)
    worksheet1.write(0, 6, 'DAN V TEDNU', format_header2)
    worksheet1.write(0, 7, 'URA V DNEVU', format_header2)
    worksheet1.write(0, 8, 'DRUGO NAJHUJSE', format_header3)
    worksheet1.write(0, 9, 'URA V LETU', format_header3)
    worksheet1.write(0, 10, 'DATUM', format_header3)
    worksheet1.write(0, 11, 'DAN V TEDNU', format_header3)
    worksheet1.write(0, 12, 'URA V DNEVU', format_header3)
    worksheet1.write(0, 13, 'Un', format_header4)
    worksheet1.write(0, 14, 'In', format_header4)
    worksheet1.write(0, 15, 'Pn', format_header4)
    worksheet1.write(0, 16, 'TIP ELEMENTA', format_header4)
    worksheet1.write(0, 17, 'GRID', format_header4)
    worksheet1.write(0, 18, 'AREA', format_header4)
    worksheet1.write(0, 19, 'ZONE', format_header4)
    
    #Zapis headerja sheet 2 - obremenitve nad xx%
    worksheet2 = workbook.add_worksheet('LoadFlowNad%')
    worksheet2.set_column(0, 2, 15)
    worksheet2.write(0, 0, 'PRAVO IME', format_header1)
    worksheet2.write(0, 1, 'ADVANCED IME', format_header1)
    worksheet2.write(0, 2, 'PF IME', format_header1)
    worksheet2.set_column(3, 19, 12)
    worksheet2.set_row(0, 30)
    #Zapisovanje naslovov stolpcev
    for percent10 in range(0,16):
        column_name = 'Nad ' + str(percent10*10) + '% [h]'
        column_number = percent10 + 3
        if percent10 < 7:
            #Za procente pod 7 pustimo obarvan zeleno
            worksheet2.write(0, column_number, column_name, format_header5)
        
        if percent10 >= 7 and percent10 < 10:
            #Za procente med 7 in 10 obarvan rumeno
            worksheet2.write(0, column_number, column_name, format_header6)
        
        if percent10 >= 10:
            #Za procente nad 10 pobarvamo header rdece
            worksheet2.write(0, column_number, column_name, format_header7)
    worksheet2.set_column(19, 21, 7)
    worksheet2.set_column(22, 22, 16)
    worksheet2.set_column(23, 25, 7)
    worksheet2.write(0, 19, 'Un')
    worksheet2.write(0, 20, 'In')
    worksheet2.write(0, 21, 'Pn')
    worksheet2.write(0, 22, 'TIP ELEMENTA')
    worksheet2.write(0, 23, 'GRID')
    worksheet2.write(0, 24, 'AREA')
    worksheet2.write(0, 35, 'ZONE')
    
    #Zapis sheeta 3
    
    worksheet3 = workbook.add_worksheet('NapetostiMax')
    worksheet3.set_row(0, 30)
    worksheet3.set_column(0, 2, 20)
    worksheet3.set_column(3, 3, 20)
    worksheet3.set_column(4, 5, 10)
    worksheet3.set_column(6, 7, 13)
    worksheet3.set_column(8, 8, 24)
    worksheet3.set_column(9, 10, 10)
    worksheet3.set_column(11, 12, 13)
    worksheet3.set_column(13, 15, 7)
    worksheet3.set_column(16, 16, 13)
    worksheet3.set_column(17, 18, 7)
    worksheet3.set_column(19, 19, 20)
    worksheet3.write(0, 0, 'PRAVO IME', format_header1)
    worksheet3.write(0, 1, 'ADVANCED IME', format_header1)
    worksheet3.write(0, 2, 'POWERFACTORY IME', format_header1)
    worksheet3.write(0, 3, 'NAJVISJA NAPETOST', format_header2)
    worksheet3.write(0, 4, 'URA V LETU', format_header2)
    worksheet3.write(0, 5, 'DATUM', format_header2)
    worksheet3.write(0, 6, 'DAN V TEDNU', format_header2)
    worksheet3.write(0, 7, 'URA V DNEVU', format_header2)
    worksheet3.write(0, 8, 'DRUGA NAJVISJA NAPETOST', format_header3)
    worksheet3.write(0, 9, 'URA V LETU', format_header3)
    worksheet3.write(0, 10, 'DATUM', format_header3)
    worksheet3.write(0, 11, 'DAN V TEDNU', format_header3)
    worksheet3.write(0, 12, 'URA V DNEVU', format_header3)
    worksheet3.write(0, 13, 'Un', format_header4)
    worksheet3.write(0, 14, 'In', format_header4)
    worksheet3.write(0, 15, 'Pn', format_header4)
    worksheet3.write(0, 16, 'TIP ELEMENTA', format_header4)
    worksheet3.write(0, 17, 'GRID', format_header4)
    worksheet3.write(0, 18, 'AREA', format_header4)
    worksheet3.write(0, 19, 'ZONE', format_header4)
    #Zapis sheeta 4
    
    worksheet4 = workbook.add_worksheet('NapetostiMaxCasovno')
    
    #Zapis sheeta 5
    
    worksheet5 = workbook.add_worksheet('NapetostiMin')
    worksheet5.set_row(0, 30)
    worksheet5.set_column(0, 2, 20)
    worksheet5.set_column(3, 3, 20)
    worksheet5.set_column(4, 5, 10)
    worksheet5.set_column(6, 7, 13)
    worksheet5.set_column(8, 8, 24)
    worksheet5.set_column(9, 10, 10)
    worksheet5.set_column(11, 12, 13)
    worksheet5.set_column(13, 15, 7)
    worksheet5.set_column(16, 16, 13)
    worksheet5.set_column(17, 18, 7)
    worksheet5.set_column(19, 19, 20)
    worksheet5.write(0, 0, 'PRAVO IME', format_header1)
    worksheet5.write(0, 1, 'ADVANCED IME', format_header1)
    worksheet5.write(0, 2, 'POWERFACTORY IME', format_header1)
    worksheet5.write(0, 3, 'NAJNIZJA NAPETOST', format_header2)
    worksheet5.write(0, 4, 'URA V LETU', format_header2)
    worksheet5.write(0, 5, 'DATUM', format_header2)
    worksheet5.write(0, 6, 'DAN V TEDNU', format_header2)
    worksheet5.write(0, 7, 'URA V DNEVU', format_header2)
    worksheet5.write(0, 8, 'DRUGA NAJNIZJA NAPETOST', format_header3)
    worksheet5.write(0, 9, 'URA V LETU', format_header3)
    worksheet5.write(0, 10, 'DATUM', format_header3)
    worksheet4.write(0, 11, 'DAN V TEDNU', format_header3)
    worksheet5.write(0, 12, 'URA V DNEVU', format_header3)
    worksheet5.write(0, 13, 'Un', format_header4)
    worksheet5.write(0, 14, 'In', format_header4)
    worksheet5.write(0, 15, 'Pn', format_header4)
    worksheet5.write(0, 16, 'TIP ELEMENTA', format_header4)
    worksheet5.write(0, 17, 'GRID', format_header4)
    worksheet5.write(0, 18, 'AREA', format_header4)
    worksheet5.write(0, 19, 'ZONE', format_header4)
    
    #Zapis sheeta 6
    
    worksheet6 = workbook.add_worksheet('NapetostiMinCasovno')
    
    #Zapis sheeta 7
    
    worksheet7 = workbook.add_worksheet('ContingencyNajhujsi')
    
    #Zapis sheeta 8
    
    worksheet8 = workbook.add_worksheet('ContingencyCasovno')
    
    #Zapis sheeta 9 - porocilo kateri izracuni so konvergirali
    
    worksheet9 = workbook.add_worksheet('Porocilo')
    #Beri datoteko nekonvergence.xlsx in potem zapisi z malo lepsim formatiranjem v to datoteko
    #data_konvergence = os.getcwd()  + r'/Rezultati/Raw data voltage.xlsx'
    #df_results_loading_hour.to_excel(data_konvergence, sheet_name = 'DATA', index_col=0)
    
    current_row = 1
    for line in lines:
        if (current_row % 2) == 0:
            row_format = format_data_lighter
        else:
            row_format = format_data_darker
        line_name = line.loc_name
        if line_name in results_hourly_elementlist:
            #Zapisemo podatke objektov
            line_rated_voltage = line.typ_id.uline #Nazivna napetost v kV
            line_rated_current = round(line.typ_id.sline * 1000) #Nazivni tok v A
            line_rated_power = round(line_rated_voltage * line_rated_current * 1.73205 / 1000) # Nazivna moč MW
            line_grid = line.cpGrid.loc_name
            try:
                line_area = line.cpArea.loc_name
            except:
                line_area = "NOAREA"
            try:
                line_zone = line.cpZone.loc_name
            except:
                line_zone = "NOZONE"
            #app.PrintInfo(current_row)
            #Osnovni podatki za worksheet 1
            worksheet1.write(current_row, 0, "", row_format)
            worksheet1.write(current_row, 1, "", row_format)
            worksheet1.write(current_row, 2, line_name, row_format)
            #worksheet1.write(current_row, 3, "/", row_format)
            #worksheet1.write(current_row, 4, "/", row_format)
            
            max_no1_hour = max_min_values.at[line_name, 'max_no1_loading']
            #app.PrintInfo(max_no1_hour)
            try:
                worksheet1.write(current_row, 3, df_results_loading_hour.at[line_name, max_no1_hour], row_format)
            except:
                worksheet1.write(current_row, 3, 0, row_format)
                
            # Za dobit column z max vrednostjo https://www.skytowner.com/explore/getting_column_label_of_max_value_in_each_row_in_pandas_datafrme
            
            worksheet1.write(current_row, 4, int(max_no1_hour), row_format)
            year = int(2030)
            if print_info: app.PrintInfo(max_no1_hour)
            if print_info: app.PrintInfo(type(max_no1_hour))
            yearhour = int(max_no1_hour) + 1
            year, month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
            dayofweek = dobiDanVtednu(year, month, day)
            #dayofweek = dobiDanVtednu(year, month, day)
            worksheet1.write(current_row, 5, str(day) + '.' + str(month) + '.' + str(year), row_format)
            worksheet1.write(current_row, 6, str(dayofweek), row_format)
            worksheet1.write(current_row, 7, int(hour), row_format)
            
            max_no2_hour = max_min_values.at[line_name, 'max_no2_loading']
            try:
                worksheet1.write(current_row, 8, df_results_loading_hour.at[line_name, max_no2_hour], row_format)
            except:
                worksheet1.write(current_row, 8, 0, row_format)
                
            worksheet1.write(current_row, 9, int(max_no2_hour), row_format)
            year = int(2030)
            yearhour = int(max_no2_hour) + 1
            year, month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
            dayofweek = dobiDanVtednu(year, month, day)
            worksheet1.write(current_row, 10, str(day) + '.' + str(month) + '.' + str(year), row_format)
            worksheet1.write(current_row, 11, str(dayofweek), row_format)
            worksheet1.write(current_row, 12, int(hour), row_format)
            worksheet1.write(current_row, 13, line_rated_voltage, row_format)
            worksheet1.write(current_row, 14, line_rated_current, row_format)
            worksheet1.write(current_row, 15, line_rated_power, row_format)
            worksheet1.write(current_row, 16, 'Daljnovod', row_format)
            worksheet1.write(current_row, 17, line_grid, row_format)
            worksheet1.write(current_row, 18, line_area, row_format)
            worksheet1.write(current_row, 19, line_zone, row_format)
            
            #Worksheet 2 z rezultati za % obremenitve in ure
            worksheet2.write(current_row, 2, line_name)
            for percent10 in range(0,21):
                dataframe_name = 'nad' + str(percent10*10)
                column_number = percent10 + 3
                worksheet2.write(current_row, column_number, percentload.at[line_name, dataframe_name])
            worksheet2.write(current_row, 24, line_rated_voltage)
            worksheet2.write(current_row, 25, line_rated_current)
            worksheet2.write(current_row, 26, line_rated_power)
            worksheet2.write(current_row, 27, 'Daljnovod')
            worksheet2.write(current_row, 28, line_grid)
            worksheet2.write(current_row, 29, line_area)
            worksheet2.write(current_row, 30, line_zone)
            current_row += 1
            
    for transformer in transformers:
        transformer_name = transformer.loc_name
        if transformer_name in results_hourly_elementlist:
            transformer_grid = transformer.cpGrid.loc_name
            try:
                transformer_area = transformer.cpArea.loc_name
            except:
                transformer_area = "NOAREA"
            try:
                transformer_zone = transformer.cpZone.loc_name
            except:
                transformer_zone = "NOZONE"
            #Worksheet 2 z rezultati za % obremenitve in ure
            worksheet2.write(current_row, 2, transformer_name)
            for percent10 in range(0,21):
                dataframe_name = 'nad' + str(percent10*10)
                column_number = percent10 + 3
                #worksheet2.write(current_row, column_number, percentload.at[transformer_name, dataframe_name])
                worksheet2.write(current_row, column_number, percentload.at[transformer_name, dataframe_name])
                #worksheet2.write(current_row, 24, line_rated_voltage)
                #worksheet2.write(current_row, 25, line_rated_current)
                #worksheet2.write(current_row, 26, line_rated_power)
            worksheet2.write(current_row, 27, 'Transformator')
            worksheet2.write(current_row, 28, transformer_grid)
            worksheet2.write(current_row, 29, transformer_area)
            worksheet2.write(current_row, 30, transformer_zone)
            current_row += 1
        
    #Izpis za zbiralke (max in min napetosti)
    for terminal in terminals:
        if (current_row % 2) == 0:
            row_format = format_data_lighter
        else:
            row_format = format_data_darker

        terminal_name = terminal.loc_name
        if terminal_name in results_hourly_elementlist_voltage:
            terminal_grid = terminal.cpGrid.loc_name
            try:
                terminal_area = terminal.cpArea.loc_name
            except:
                terminal_area = "NOAREA"
            try:
                terminal_zone = terminal.cpZone.loc_name
            except:
                terminal_zone = "NOZONE"
            #Osnovni podatki za worksheet 1
            worksheet3.write(current_row, 0, "", row_format)
            worksheet3.write(current_row, 1, "", row_format)
            worksheet3.write(current_row, 2, line_name, row_format)
            #worksheet1.write(current_row, 3, "/", row_format)
            #worksheet1.write(current_row, 4, "/", row_format)
            
            max_no1_hour = max_min_values.at[line_name, 'max_no1_voltage']
            #app.PrintInfo(max_no1_hour)
            try:
                worksheet3.write(current_row, 3, df_results_voltage_hour.at[line_name, max_no1_hour], row_format)
            except:
                worksheet3.write(current_row, 3, 0, row_format)
                
            # Za dobit column z max vrednostjo https://www.skytowner.com/explore/getting_column_label_of_max_value_in_each_row_in_pandas_datafrme
            
            worksheet1.write(current_row, 4, int(max_no1_hour), row_format)
            year = int(2030)
            if print_info: app.PrintInfo(max_no1_hour)
            if print_info: app.PrintInfo(type(max_no1_hour))
            yearhour = int(max_no1_hour) + 1
            year, month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
            dayofweek = dobiDanVtednu(year, month, day)
            #dayofweek = dobiDanVtednu(year, month, day)
            worksheet3.write(current_row, 5, str(day) + '.' + str(month) + '.' + str(year), row_format)
            worksheet3.write(current_row, 6, str(dayofweek), row_format)
            worksheet3.write(current_row, 7, int(hour), row_format)
            
            max_no2_hour = max_min_values.at[line_name, 'max_no2_voltage']
            try:
                worksheet3.write(current_row, 8, df_results_voltage_hour.at[line_name, max_no2_hour], row_format)
            except:
                worksheet3.write(current_row, 8, 0, row_format)
                
            worksheet1.write(current_row, 9, int(max_no2_hour), row_format)
            year = int(2030)
            yearhour = int(max_no2_hour) + 1
            year, month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
            dayofweek = dobiDanVtednu(year, month, day)
            worksheet3.write(current_row, 10, str(day) + '.' + str(month) + '.' + str(year), row_format)
            worksheet3.write(current_row, 11, str(dayofweek), row_format)
            worksheet3.write(current_row, 12, int(hour), row_format)

    workbook.close()
    return
    
def writeEndResultsToExcel(lines, grids_to_write, hours):
    df_results_lines_hourly = pd.DataFrame(data=None)
    df_results_lines_condensed = pd.DataFrame(data=None)
    for line in lines:
        line_name = line.loc_name
        line_grid = line.cpGrid.loc_name
        if line_grid in grids_to_write:
            line_rated_voltage = line.typ_id.uline
            line_rated_current = line.typ_id.sline
            line_area = line.cpArea.loc_name
            line_zone = line.cpZone.loc_name
            df_results_lines_condensed.at[line_name, 'PRAVO IME'] = 'neki'
            df_results_lines_condensed.at[line_name, 'ADVANCED IME'] = 'neki'
            df_results_lines_condensed.at[line_name, 'IME PF'] = line_name
            df_results_lines_condensed.at[line_name, 'GRID'] = line_grid
            df_results_lines_condensed.at[line_name, 'AREA'] = line_area
            df_results_lines_condensed.at[line_name, 'ZONE'] = line_zone
            df_results_lines_condensed.at[line_name, 'RATED VOLTAGE [kV]'] = line_rated_voltage
            df_results_lines_condensed.at[line_name, 'RATED CURRENT [A]'] = round(line_rated_current * 1000)
            df_results_lines_condensed.at[line_name, 'RATED S [MW]'] = round(line_rated_voltage * line_rated_current * 1.73205)
            df_results_lines_condensed.at[line_name, 'Najhujsa obremenitev [%]'] = 'stevlka'
            #Line je out of service (izpad)
            
    #Potem zankamo da izpisujemo rezultate za vsako uro
    for hour in hours:
        current_file_name = 'Rezultati_ura_' + str(hour)
        hour_file_path = os.getcwd()  + r'/Vmesni rezutati/' + current_file_name +'.csv'
        current_hour_df = pd.read_csv(hour_file_path, index_col = 0)
        app.PrintInfo(current_hour_df)
        df_results_lines_hourly = pd.concat([df_results_lines_hourly, current_hour_df], axis=1, join='inner')
    #Tu bi zdej še gledali kolk ur je nad 80, nad 90 in nas 100% obremenitve in lahko izpisal worst case
    #df_results_lines['Nad 20%']=df_results_lines.iloc[:,1:].ge(df_results_lines.columns[hours]).sum(axis=1)
    outputfilename = 'Obremenitve daljnovodov'
    outputfilepath = os.getcwd()  + r'/Rezultati/' + outputfilename +'.xlsx'
    df_results_lines_condensed.to_excel(outputfilepath, sheet_name = 'Condensed data')
    df_results_lines_hourly.to_excel(outputfilepath, sheet_name = 'Hourly data')
    return     

def pisiNekonvergencoVCsv(df_nekonvergenca):
    app.PrintInfo("Zapis nekonvergence")
    #outputfilepath = r"S:\SlapnikL_Mag\Vecscenarijska_Analiza\Rezultati\Daljnovodi_Obremenitve.xlsx"
    #imeIzhodneDatoteke = "Rezultati vecscenarijske LoadFlow analize.xlsx"
    # print(antares_file_excel[0])
    #outputfilepath = os.getcwd() + "Rezultati" + "\\" + imeIzhodneDatoteke
    outputfilename = 'Nekonvergence'
    outputfilepath = os.getcwd()  + r'/Rezultati/' + outputfilename +'.xlsx'
    df_nekonvergenca.to_excel(outputfilepath)
    
def dobiDatumIzUreVLetu(leto, uravletu):
    #leto = 2022
    #uravletu = 1
    # leto, mesec, dan, ura = dobiDatumIzUreVLetu(leto,uravletu)
    # danVtednu = dobiDanVtednu(leto, mesec, dan)
    # print ("Datum: " + str(dan) + '.' + str(mesec) + '.' + str(leto) + ' ob ' + str(ura) + ':00 (' + danVtednu + ').')
    
    prestopno = 0
    # preverimo ce je leto prestopno
    if(leto%4==0 and leto%100!=0 or leto%400==0):
        prestopno = 24

    # Create the starting date as a `datetime` object.
    start = dt(leto, 1, 1, 0, 0, 0)
    # List initialiser.
    result = [start]
    
    # Build a list of datetime objects for each hour of the year.
    for i in range(1, 8760 + prestopno):
        start += td(seconds=3600)
        result.append(start)
    
    # Initialise a DataFrame data structure.
    df = pd.DataFrame({'dates': result})
    # Add each column by extracting the object of interest from the datetime.
    df['8760'] = df.index+1
    df['month'] = df['dates'].dt.month
    df['day'] = df['dates'].dt.day
    df['hour'] = df['dates'].dt.hour
    # Remove the datetime object column.
    df.drop(['dates'], inplace=True, axis=1)
    #df['8760'][ura-1]
    mesec = df['month'][uravletu-1]
    dan = df['day'][uravletu-1]
    ura = df['hour'][uravletu-1]
    #df['month'][ura-1]
    return leto,mesec,dan,ura

def dobiDanVtednu(year, month, day):
    # leto = 2022
    # uravletu = 1
    # leto, mesec, dan, ura = dobiDatumIzUreVLetu(leto,uravletu)
    # danVtednu = dobiDanVtednu(leto, mesec, dan)
    # print ("Datum: " + str(dan) + '.' + str(mesec) + '.' + str(leto) + ' ob ' + str(ura) + ':00 (' + danVtednu + ').')
    
    offset = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    week   = ['Nedelja', 
              'Ponedeljek', 
              'Torek', 
              'Sreda', 
              'Četrtek',  
              'Petek', 
              'Sobota']
    afterFeb = 1
    if month > 2: afterFeb = 0
    aux = year - 1700 - afterFeb
    # dayOfWeek for 1700/1/1 = 5, Friday
    dayOfWeek  = 5
    # partial sum of days betweem current date and 1700/1/1
    dayOfWeek += (aux + afterFeb) * 365                  
    # leap year correction    
    dayOfWeek += aux // 4 - aux // 100 + (aux + 100) // 400     
    # sum monthly and day offsets
    dayOfWeek += offset[month - 1] + (day - 1)               
    dayOfWeek %= 7
    return week[dayOfWeek]

def ContingencyAnalysis(hour):
    #contingencies = app.GetCalcRelevantObjects("*.")
    #app.PrintInfo(contingencies)
    
    #test = app.GetFromStudyCase("ComOutage")
    #test.GetObject()
    
    #app.PrintInfo(test_return)
    
    #contingency_file_name = 'Contingency_ura_' + str(hour)
    #contingency_file_path = os.getcwd()  + r'/Vmesni rezutati/' + contingency_file_name +'.csv'
    
    #Najprej je potrebno izvesti contingency analysis
    ctg = app.GetFromStudyCase("ComSimoutage")
    ctg.Execute()
    
    #Potem mormo zalaufat contingency report
    comres = app.GetFromStudyCase("ComRes")
    comres.Execute()
    
    #Report shranmo z DPL skripto ker se se tega ne da z Pythonom
    outputfilename = "Contingency_report_" + hour
    contingency_export_DPL_script = os.getcwd()  + r'/Rezultati/' + outputfilename +'.xlsx'

#######################################################################################################################################

def gridsInterchange(grids, grids_to_check):
    for grid1 in grids:
        #app.PrintInfo(grid1.loc_name)
        if grid1.loc_name in grids_to_check:
            for grid2 in grids:
                if grid2.loc_name != grid1.loc_name and grid2.loc_name in grids_to_check:
                    interchange_value = grid1.CalculateInterchangeTo(grid2)
                    if interchange_value > 0:
                        app.PrintInfo("Izmenjava delovne moči med grid1: " + str(grid1.loc_name) + " in grid2: " + str(grid2.loc_name) + " je: " + str(grid1.GetAttribute('c:InterP')))
                        app.PrintInfo("Izmenjava jalove moči med grid1: " + str(grid1.loc_name) + " in grid2: " + str(grid2.loc_name) + " je: " + str(grid1.GetAttribute('c:InterQ')))
    
def testing():
    #hours = [1,3,5,12,25,65,111,112,113,1254,3431,6900]
    hours = range(1,10)
    #hours = 3682
    grids_to_set, grids_to_write = importParameters()
    grids_to_check = grids_to_set
    df_resultldf = pd.DataFrame(data=None)
    #grids_to_set_pf, grids_to_write_results = importParameters()
    generators = app.GetCalcRelevantObjects("*.ElmSym")
    loads = app.GetCalcRelevantObjects("*.ElmLod")
    voltsrcs = app.GetCalcRelevantObjects("*.ElmVac")
    lines = app.GetCalcRelevantObjects("*.ElmLne")
    transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    grids = app.GetCalcRelevantObjects("*.ElmNet")
    terminals = app.GetCalcRelevantObjects("*.ElmTerm")
    
    df_hourly_market_data, df_crossborder_exchanges, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, df_robna_vozlisca = importCSVFiles()
    
    generator_ratios = calcGenRatios(generators, df_izbrana_vozlisca_p)
    load_ratios = calcLoadRatios(loads, df_izbrana_vozlisca_p)
    
    for hour in hours:
        #calcAndSetGenPower(generators, df_hourly_market_data, generator_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour)
        #calcAndSetLoadPower(loads, df_hourly_market_data, load_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour)    
        calcAndSetGenLoadPower(generators, loads, df_hourly_market_data, generator_ratios, load_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour, grids_to_set)
        nastaviRobnavozlisca(voltsrcs, df_crossborder_exchanges, df_robna_vozlisca, hour)
        #cekiraj doloceno moc v slo
        #app.PrintInfo("DOLOCENA MOC PROIZVODNJE V SI00: " + str(checkAllGenSlo(generators, grids_to_check))) 
        #app.PrintInfo("DOLOCENA MOC BREMEN V SI00: " + str(checkAllLoadSlo(loads, grids_to_check))) 
        
        #Izvedi load flow
        nekonvergenca = ldf.Execute()
        app.PrintInfo("ZAPIS REZULTATOV")
        if nekonvergenca == 0:
            #df_results_lines = writeResultsLinesToDF(df_results_lines, lines, countries_to_write_results, hour)
            shraniVmesneRezultateCsv(lines, transformers, terminals, grids_to_write, hour)
            df_resultldf.at[hour,"Konvergenca"] = "DA"
            #gridsInterchange(grids, grids_to_check)
        else:
            #df_nekonvergenca = shraniNekonvergenco(df_nekonvergenca, nekonvergenca, hour)
            df_resultldf.at[hour,"Konvergenca"] = "NE"
            app.PrintWarn("NEKONVERGENCA, REZULTAT NI PISAN")
    
    #df_results_lines = writeEndResultsToExcel(lines, grids_to_write, hours)
    #app.PrintInfo(df_results_lines)
    return

def testGridExchange():
    hour = 3682
    grids_to_set, grids_to_write = importParameters()
    grids_to_check = grids_to_set
    generators = app.GetCalcRelevantObjects("*.ElmSym")
    loads = app.GetCalcRelevantObjects("*.ElmLod")
    voltsrcs = app.GetCalcRelevantObjects("*.ElmVac")
    grids = app.GetCalcRelevantObjects("*.ElmNet")
    
    df_hourly_market_data, df_crossborder_exchanges, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, df_robna_vozlisca = importCSVFiles()
    
    generator_ratios = calcGenRatios(generators, df_izbrana_vozlisca_p)
    load_ratios = calcLoadRatios(loads, df_izbrana_vozlisca_p)
    
    calcAndSetGenLoadPower(generators, loads, df_hourly_market_data, generator_ratios, load_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour, grids_to_set)
    nastaviRobnavozlisca(voltsrcs, df_crossborder_exchanges, df_robna_vozlisca, hour)
        
    #Izvedi load flow
    ldf.Execute()
    gridsInterchange(grids, grids_to_check)
    return

def testGridExchange1():
    grids_to_set, grids_to_write = importParameters()
    grids_to_check = grids_to_set
    grids = app.GetCalcRelevantObjects("*.ElmNet")
    
    df_hourly_market_data, df_crossborder_exchanges, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, df_robna_vozlisca = importCSVFiles()
    
    gridsInterchange(grids, grids_to_check)
    return

def rezultatiVse():
    hours = range(1,5)
    transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    lines = app.GetCalcRelevantObjects("*.ElmLne")
    terminals = app.GetCalcRelevantObjects("*.ElmTerm")
    grids_to_write = ['SI00', 'ELES Interconnectios']
    writeEndResultsXlsxwriter(lines, transformers, terminals, grids_to_write, hours)
    
    return

def setSingleHour():
    #hour = 1
    hour = 3682
    grids_to_set, grids_to_write = importParameters()
    grids_to_check = grids_to_set
    generators = app.GetCalcRelevantObjects("*.ElmSym")
    loads = app.GetCalcRelevantObjects("*.ElmLod")
    voltsrcs = app.GetCalcRelevantObjects("*.ElmVac")
    df_hourly_market_data, df_crossborder_exchanges, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, df_robna_vozlisca = importCSVFiles()
    generator_ratios = calcGenRatios(generators, df_izbrana_vozlisca_p)
    load_ratios = calcLoadRatios(loads, df_izbrana_vozlisca_p)
    calcAndSetGenLoadPower(generators, loads, df_hourly_market_data, generator_ratios, load_ratios, PFoldcosfi, df_izbrana_vozlisca_p, df_izbrana_vozlisca_q, hour, grids_to_set)
    nastaviRobnavozlisca(voltsrcs, df_crossborder_exchanges, df_robna_vozlisca, hour)
    
    return
#################################################################################### MAIN #######################################################

#vecUrIzbranaVozlisca()
#compare()

#testing()
#testGridExchange1()
rezultatiVse()
#setSingleHour()
#ContingencyAnalysis(3682)
######### KONC

end_time = datetime.datetime.now().time().strftime('%H:%M:%S')
total_time=(datetime.datetime.strptime(end_time,'%H:%M:%S') - datetime.datetime.strptime(start_time,'%H:%M:%S'))
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
app.PrintInfo("Konec izvajanja programa ob " + str(current_time) + ". Potreben čas: " + str(total_time) + '.')