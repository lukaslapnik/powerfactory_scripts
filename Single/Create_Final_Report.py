# -*- coding: utf-8 -*-
"""
Created on Wed Apr 26 13:29:56 2023

@author: Lukc
"""
import pandas as pd
import datetime
import sys
import math
#import numpy
import os
from os import listdir
import glob
import time
from datetime import datetime as dt
from datetime import timedelta as td
from tkinter import Tk
from tkinter import filedialog
import xlsxwriter

##################### PARAMETRI ######################

year = int(2030)
    
#####################################################

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
    return mesec,dan,ura

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

# line_loading_file_path = os.path.join(f_output_data_directory,("Line_loading.csv"))
# transformer_loading_file_path = os.path.join(f_output_data_directory,("Transformer_loading.csv"))
# voltage_file_path = os.path.join(f_output_data_directory,("Terminal_voltage.csv"))
# generator_P_set_file_path = os.path.join(f_output_data_directory,("Generator_P.csv"))
# generator_Q_set_file_path = os.path.join(f_output_data_directory,("Generator_Q.csv"))
# load_P_set_file_path = os.path.join(f_output_data_directory,("Load_P.csv"))
# load_Q_set_file_path = os.path.join(f_output_data_directory,("Load_Q.csv"))

# df_results_load_P_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Load_P.csv"), index_col = 0)
# df_results_load_Q_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Load_Q.csv"), index_col = 0)
# # Bremena...
# df_results_generator_P_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Generator_P.csv"), index_col = 0)
# df_results_generator_Q_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Generator_Q.csv"), index_col = 0)
# # Daljnovodi...
# df_results_line_loading_final = pd.read_csv(os.path.join(f_input_data_directory,"Line_loading.csv"), index_col = 0)
# # Transformatorji...
# df_results_transformer_loading_final = pd.read_csv(os.path.join(f_input_data_directory,"Transformer_loading.csv"), index_col = 0)
# # Zbiralke...
# df_results_terminal_voltage_final = pd.read_csv(os.path.join(f_input_data_directory,"Terminal_voltage.csv"), index_col = 0)

# IMPORT PODATKOV
print("Izberi mapo s podatki (lahko je skrita za oknom v ozadju)!")
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
f_input_data_directory = filedialog.askdirectory()
print("Vhodna mapa izbrna, uvoz podatkov!")

# Daljnovodi...
df_line_info = pd.read_csv(os.path.join(f_input_data_directory,"Line_Info.csv"), index_col = 0)
df_results_line_loading_final = pd.read_csv(os.path.join(f_input_data_directory,"Line_loading.csv"), index_col = 0)
# Transformatorji...
df_transformer_info = pd.read_csv(os.path.join(f_input_data_directory,"Transformer_Info.csv"), index_col = 0)
df_results_transformer_loading_final = pd.read_csv(os.path.join(f_input_data_directory,"Transformer_loading.csv"), index_col = 0)
# Zbiralke...
df_terminal_info = pd.read_csv(os.path.join(f_input_data_directory,"Terminal_Info.csv"), index_col = 0)
df_results_terminal_voltage_final = pd.read_csv(os.path.join(f_input_data_directory,"Terminal_voltage.csv"), index_col = 0)
# Bremena...
df_load_info = pd.read_csv(os.path.join(f_input_data_directory,"Load_Info.csv"), index_col = 0)
df_results_load_P_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Load_P.csv"), index_col = 0)
df_results_load_Q_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Load_Q.csv"), index_col = 0)
# Generatorji...
df_generator_info = pd.read_csv(os.path.join(f_input_data_directory,"Gen_Info.csv"), index_col = 0)
df_results_generator_P_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Generator_P.csv"), index_col = 0)
df_results_generator_Q_set_final = pd.read_csv(os.path.join(f_input_data_directory,"Generator_Q.csv"), index_col = 0)

print("Podatki uvoženi, zacnenjam obdelavo!")

hours = df_results_line_loading_final.columns.to_list()
hours = list(map(int, hours))

t1 = time.time()
# dt0 = t1 - t0
# print("Cas uvoza in izpisa zdruzenih podatkov: " + str (dt0))

#Liste elementov, ker se dela izpis samo za te
results_element_list_generators = df_results_generator_P_set_final.index
results_element_list_loads = df_results_load_P_set_final.index
results_element_list_lines = df_results_line_loading_final.index
results_element_list_transformers = df_results_transformer_loading_final.index
results_element_list_terminals = df_results_terminal_voltage_final.index

# print("Generatorji za izpis: " + str(results_element_list_generators))
# print("Loadi za izpis: " + str(results_element_list_loads))
# print("DV za izpis: " + str(results_element_list_lines))
# print("TR za izpis: " + str(results_element_list_transformers))
# print("Zbiralke za izpis: " + str(results_element_list_terminals))

#Najvisje 3 obremenitve daljnovodov v obdobju
df_temp1 = df_results_line_loading_final
max_no1 = pd.DataFrame()
max_no1['max_no1_loading'] = df_temp1.idxmax(axis=1)
m = df_temp1.columns.to_numpy() == max_no1['max_no1_loading'].to_numpy()[:, None]
df_temp1 = df_temp1.mask(m)
max_no1['max_no2_loading'] = df_temp1.idxmax(axis=1)
m = df_temp1.columns.to_numpy() == max_no1['max_no2_loading'].to_numpy()[:, None]
df_temp1 = df_temp1.mask(m)
max_no1['max_no3_loading'] = df_temp1.idxmax(axis=1)
#NAN menjamo z 0 da ne pride do napak. Mogoce boljse poskrbet da do tega v prvem ne pride, filanje je taktak...
max_no1.fillna(int(0), inplace = True)

#Najvisje 3 obremenitve transformatorjev v obdobju
df_temp2 = df_results_transformer_loading_final
max_no2 = pd.DataFrame()
max_no2['max_no1_loading'] = df_temp2.idxmax(axis=1)
m = df_temp2.columns.to_numpy() == max_no2['max_no1_loading'].to_numpy()[:, None]
df_temp2 = df_temp2.mask(m)
max_no2['max_no2_loading'] = df_temp2.idxmax(axis=1)
m = df_temp2.columns.to_numpy() == max_no2['max_no2_loading'].to_numpy()[:, None]
df_temp2 = df_temp2.mask(m)
max_no2['max_no3_loading'] = df_temp2.idxmax(axis=1)
#NAN menjamo z 0 da ne pride do napak. Mogoce boljse poskrbet da do tega v prvem ne pride, filanje je taktak...
max_no2.fillna(int(0), inplace = True)

max_no3 = pd.DataFrame()
#Dobimo se najvisje 3 in najnizje 3 napetosti v obdobju.
#Problem z NAN "vrednostmi". Naceloma se lahko nardi adaptivno da to dela samo ce so vsaj 3 ure "simulirane"
df_temp3 = df_results_terminal_voltage_final.mask(df_results_terminal_voltage_final == 0)
#Iskanje najvisjih 3 napetost
max_no3['max_no1_voltage'] = df_temp3.idxmax(axis=1)
m = df_temp3.columns.to_numpy() == max_no3['max_no1_voltage'].to_numpy()[:, None]
df_temp3 = df_temp3.mask(m)
max_no3['max_no2_voltage'] = df_temp3.idxmax(axis=1)
m = df_temp3.columns.to_numpy() == max_no3['max_no2_voltage'].to_numpy()[:, None]
df_temp3 = df_temp3.mask(m)
max_no3['max_no3_voltage'] = df_temp3.idxmax(axis=1)

#shranjene najvisje 3 napetosti zbiralk, zdaj se najnizje 3 napetosti
max_no3['min_no1_voltage'] = df_temp3.idxmin(axis=1)
m = df_temp3.columns.to_numpy() == max_no3['min_no1_voltage'].to_numpy()[:, None]
df_temp3 = df_temp3.mask(m)
max_no3['min_no2_voltage'] = df_temp3.idxmin(axis=1)
m = df_temp3.columns.to_numpy() == max_no3['min_no2_voltage'].to_numpy()[:, None]
df_temp3 = df_temp3.mask(m)
max_no3['min_no3_voltage'] = df_temp3.idxmin(axis=1)

max_no3.fillna(int(1), inplace = True)

t2 = time.time()
dt1 = t2 - t1
print("Cas maxmin vrednosti: " + str (dt1))

# Obremenitve nad dolocenim % za daljnovode v korakih po 10%
percent_loading_lines = pd.DataFrame()
for i in range(0,16):
    #dataframe_name = 'nad' + str(i*10)
    percent_loading_lines['nad' + str(i*10)] = df_results_line_loading_final.iloc[:,0:].ge(float(i*10)).sum(axis=1)

# Se obremenitve za transformatorje
percent_loading_transformers = pd.DataFrame()
for i in range(0,16):
    #dataframe_name = 'nad' + str(i*10)
    percent_loading_transformers['nad' + str(i*10)] = df_results_transformer_loading_final.iloc[:,0:].ge(float(i*10)).sum(axis=1)

# Napetosti pri daljnovodh v per-unit ker so razlicni napetostni nivoji in sicer v korakih po 0.005 oz 0,5%
voltage_high_terminals = pd.DataFrame()
for i in range(0,14):
    voltage_high_terminals['nad' + str(1 + i * 0.01)] = df_results_terminal_voltage_final.iloc[:,0:].ge(float(1 + i * 0.01)).sum(axis=1)
    #df['Count']=df.iloc[:,1:].ge(df.iloc [:,0],axis=0).sum(axis=1)

voltage_low_terminals = pd.DataFrame()
for i in range(0,14):
    voltage_low_terminals['pod' + str(1 - i * 0.01)] = df_results_terminal_voltage_final.iloc[:,0:].le(float(1 - i * 0.01)).sum(axis=1)
    
# https://stackoverflow.com/questions/65802624/how-to-find-the-number-of-consecutive-values-greater-than-n-looking-back-from-t
line_list = df_results_line_loading_final.index
above_list = []
loading_start = 0
loading_stop = 110
loading_inc = 10
filter_limit = 60 #Dataframe za kriticne daljnovode z obremenitvami nad to vrednostjo
for i in range(loading_start,loading_stop,loading_inc):
    above_list.append(i)
df_consecutive_above = pd.DataFrame(columns = above_list)
df_consecutive_above_start = pd.DataFrame(columns = above_list)
#print(above_list)
for line in line_list:
    #print(line)
    #loading_list =  df_results_line_loading_final.loc[line]
    for limit in df_consecutive_above.columns:
        #max_consecutive = 0
        #hiter nacin brez da dobimo vrednosti kdaj se zacne
        #consecutive_list = [len(list(g)) for k, g in groupby(loading_list>limit) if k==True]
        #hiter nacin brez da dobimo vrednosti kdaj se zacne
        consecutive_nr = 0
        longest_consecutive = 0
        consecutive_start = 0
        longest_consecutive_start = 0
        for hour in df_results_line_loading_final.columns:
            if df_results_line_loading_final.at[line,hour] > limit:
                if consecutive_nr == 0:
                    consecutive_start = hour
                consecutive_nr += 1
            else:
                if consecutive_nr >= longest_consecutive:
                    longest_consecutive = consecutive_nr
                    longest_consecutive_start = consecutive_start
                consecutive_nr = 0
            #Ce je vse ure nad limit vrednostjo zapise to
            if hour == df_results_line_loading_final.columns[-1]:
                longest_consecutive = consecutive_nr
                longest_consecutive_start = consecutive_start
        #hiter nacin brez da dobimo vrednosti kdaj se zacne
        #if consecutive_list: max_consecutive = max(consecutive_list)
        #df_consecutive_above.at[line,limit] = max_consecutive
        #hiter nacin brez da dobimo vrednosti kdaj se zacne
        df_consecutive_above.at[line,limit] = longest_consecutive
        df_consecutive_above_start.at[line,limit] = longest_consecutive_start
    #print(max_consecutive)
#print(df_consecutive_above)
#print(df_consecutive_above_start)
# df_consecutive_above.to_excel(path + r'/Rezultati/Results_consecutive_above.xlsx', sheet_name = 'data')
#Poberemo samo krtične elemente z obremenitvami nad neko mejo
df_consecutive_above_critical = df_consecutive_above.drop(df_consecutive_above[df_consecutive_above[filter_limit] <= 0].index)
df_consecutive_above_start_critical = df_consecutive_above_start[df_consecutive_above_start.index.isin(df_consecutive_above_critical.index)]
#print(df_consecutive_above_critical)
#print(df_consecutive_above_start_critical)
#Potem naredimo se za dan, datum in uro dneva
df_consecutive_above_start_date = pd.DataFrame(columns = df_consecutive_above_critical.columns)
for line in df_consecutive_above_critical.index:
    for limit in df_consecutive_above_critical.columns:
        hour_of_year = int(df_consecutive_above_start_critical.at[line,limit]) 
        if hour_of_year > 0:
            month,day,hour = dobiDatumIzUreVLetu(year, hour_of_year)
            dayofweek = dobiDanVtednu(year, month, day)
            daydatetime = dayofweek + ", " + str(day) + "." + str(month) + "." + str(year) + ", " + str(hour) + ":00"
            df_consecutive_above_start_date.at[line,limit] = daydatetime
        else: 
            df_consecutive_above_start_date.at[line,limit] = ""
#print(df_consecutive_above_start_date)
    

# Naredimo workbook
workbook = xlsxwriter.Workbook(os.path.join(f_input_data_directory, "Rezultati.xlsx"))

format_header1 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header1.set_align('vcenter')
format_header1.set_align('center')
format_header1.set_bg_color('#b9faad')
format_header1.set_border()

format_header2 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header2.set_align('vcenter')
format_header2.set_align('center')
format_header2.set_bg_color('#fcc5c2')
format_header2.set_border()

format_header3 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header3.set_align('vcenter')
format_header3.set_align('center')
format_header3.set_bg_color('#d5fcb8')
format_header3.set_border()

format_header4 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header4.set_align('vcenter')
format_header4.set_align('center')
format_header4.set_bg_color('#90bbf8')
format_header4.set_border()

format_header5 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header5.set_align('vcenter')
format_header5.set_align('center')
format_header5.set_bg_color('#9cf05a')
format_header5.set_border()

format_header6 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header6.set_align('vcenter')
format_header6.set_align('center')
format_header6.set_bg_color('#e6f13b')
format_header6.set_border()

format_header7 = workbook.add_format({'bold': True,
                                      'text_wrap': True})
format_header7.set_align('vcenter')
format_header7.set_align('center')
format_header7.set_bg_color('#f3634f')
format_header7.set_border()

format_data_lighter = workbook.add_format({'text_wrap': True})
format_data_lighter.set_align('vcenter')
format_data_lighter.set_align('center')
format_data_lighter.set_bg_color('#f1f1f1')

format_data_darker = workbook.add_format({'text_wrap': True})
format_data_darker.set_align('vcenter')
format_data_darker.set_align('center')
format_data_darker.set_bg_color('#dcdcdc')

# Zapisemo sheet 1 dveh najhujsih obremenitev
worksheet1 = workbook.add_worksheet('Max obr.')

worksheet1.set_row(0, 34)
worksheet1_column_width = [20,20,20,17,8,25,17,8,25,17,8,25,11,6,6,6,8,8,20]
[ worksheet1.set_column(column, column, worksheet1_column_width[column]) for column in range(len(worksheet1_column_width)) ]
worksheet1.write(0, 0, 'Pravo ime', format_header1)
worksheet1.write(0, 1, 'Advanced ime', format_header1)
worksheet1.write(0, 2, 'Powerfactory ime', format_header1)
worksheet1.write(0, 3, 'Prva najhujsa obrem. v obd. [%]', format_header2)
worksheet1.write(0, 4, 'Ura v letu', format_header2)
worksheet1.write(0, 5, 'Dan v tednu, datum, ura', format_header2)
worksheet1.write(0, 6, 'Druga najhujsa obrem. v obd. [%]', format_header3)
worksheet1.write(0, 7, 'Ura v letu', format_header3)
worksheet1.write(0, 8, 'Dan v tednu, datum, ura', format_header3)
worksheet1.write(0, 9, 'Tretja najhujsa obrem. v obd. [%]', format_header4)
worksheet1.write(0, 10, 'Ura v letu', format_header4)
worksheet1.write(0, 11, 'Dan v tednu, datum, ura', format_header4)
worksheet1.write(0, 12, 'Tip elementa', format_header5)
worksheet1.write(0, 13, 'Un [kV]', format_header5)
worksheet1.write(0, 14, 'In [A]', format_header5)
worksheet1.write(0, 15, 'Pn [MW]', format_header5)
worksheet1.write(0, 16, 'Grid', format_header5)
worksheet1.write(0, 17, 'Area', format_header5)
worksheet1.write(0, 18, 'Zone', format_header5)


#Zapis headerja sheet 2 - obremenitve nad xx%
worksheet2 = workbook.add_worksheet('Obr nad %')
worksheet2.set_row(0, 34)
worksheet2_column_width = [20,20,20,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,12,11,6,6,6,8,8,20]
[ worksheet2.set_column(column, column, worksheet2_column_width[column]) for column in range(len(worksheet2_column_width)) ]
worksheet2.write(0, 0, 'Pravo ime', format_header1)
worksheet2.write(0, 1, 'Advanced ime', format_header1)
worksheet2.write(0, 2, 'Powerfactory ime', format_header1)
[ worksheet2.write(0, i + 3, 'St. ur nad ' + str(i*10) + ' % obr.', format_header2) for i in range(0,16) ]
#Zapisovanje naslovov stolpcev
# for percent10 in range(0,16):
#     column_name = 'Nad ' + str(percent10*10) + '% [h]'
#     column_number = percent10 + 3
#     if percent10 < 7:
#         #Za procente pod 7 pustimo obarvan zeleno
#         worksheet2.write(0, column_number, column_name, format_header5)
    
#     if percent10 >= 7 and percent10 < 10:
#         #Za procente med 7 in 10 obarvan rumeno
#         worksheet2.write(0, column_number, column_name, format_header6)
    
#     if percent10 >= 10:
#         #Za procente nad 10 pobarvamo header rdece
#         worksheet2.write(0, column_number, column_name, format_header7)
worksheet2.write(0, 19, 'Tip elementa', format_header4)
worksheet2.write(0, 20, 'Un [kV]', format_header4)
worksheet2.write(0, 21, 'In [A]', format_header4)
worksheet2.write(0, 22, 'Pn [MW]', format_header4)
worksheet2.write(0, 23, 'Grid', format_header4)
worksheet2.write(0, 24, 'Area', format_header4)
worksheet2.write(0, 25, 'Zone', format_header4)

worksheet3 = workbook.add_worksheet('Obr zaporedno')
worksheet3.set_row(0, 34)
worksheet3_column_width = [20,20,20,12,8,25,12,8,25,12,8,25,11,6,6,6,8,8,20]
[ worksheet3.set_column(column, column, worksheet3_column_width[column]) for column in range(len(worksheet3_column_width)) ]
worksheet3.write(0, 0, 'Pravo ime', format_header1)
worksheet3.write(0, 1, 'Advanced ime', format_header1)
worksheet3.write(0, 2, 'Powerfactory ime', format_header1)
[ worksheet3.write(0, i*3+3, 'Zap. ur nad ' + str(60+i*20) + '% obr.', format_header2) for i in range(0,3) ]
[ worksheet3.write(0, i*3+4, 'Ura v letu', format_header2) for i in range(0,3) ]
[ worksheet3.write(0, i*3+5, 'Dan v tednu, datum, ura', format_header2) for i in range(0,3) ]
#Zapisovanje naslovov stolpcev
# for percent10 in range(0,16):
#     column_name = 'Nad ' + str(percent10*10) + '% [h]'
#     column_number = percent10 + 3
#     if percent10 < 7:
#         #Za procente pod 7 pustimo obarvan zeleno
#         worksheet3.write(0, column_number, column_name, format_header5)
    
#     if percent10 >= 7 and percent10 < 10:
#         #Za procente med 7 in 10 obarvan rumeno
#         worksheet3.write(0, column_number, column_name, format_header6)
    
#     if percent10 >= 10:
#         #Za procente nad 10 pobarvamo header rdece
#         worksheet3.write(0, column_number, column_name, format_header7)
worksheet3.write(0, 12, 'Tip elementa', format_header4)
worksheet3.write(0, 13, 'Un [kV]', format_header4)
worksheet3.write(0, 14, 'In [A]', format_header4)
worksheet3.write(0, 15, 'Pn [MW]', format_header4)
worksheet3.write(0, 16, 'Grid', format_header4)
worksheet3.write(0, 17, 'Area', format_header4)
worksheet3.write(0, 18, 'Zone', format_header4)

#Zapis sheeta 4 - napetosti max

worksheet4 = workbook.add_worksheet('Napetosti Max')
worksheet4.set_row(0, 34)
worksheet4_column_width = [20,20,20,15,7,25,15,7,25,15,7,25,11,6,8,8,20]
[ worksheet4.set_column(column, column, worksheet4_column_width[column]) for column in range(len(worksheet4_column_width)) ]
worksheet4.write(0, 0, 'Pravo ime', format_header1)
worksheet4.write(0, 1, 'Advanced ime', format_header1)
worksheet4.write(0, 2, 'Powerfactory ime', format_header1)
worksheet4.write(0, 3, 'Prva najvišja nap. v obd. [kV]', format_header2)
worksheet4.write(0, 4, 'Ura v letu', format_header2)
worksheet4.write(0, 5, 'Dan v tednu, datum, ura', format_header2)
worksheet4.write(0, 6, 'Druga najvišja nap. v obd. [kV]', format_header3)
worksheet4.write(0, 7, 'Ura v letu', format_header3)
worksheet4.write(0, 8, 'Dan v tednu, datum, ura', format_header3)
worksheet4.write(0, 9, 'Druga najvišja nap. v obd. [kV]', format_header4)
worksheet4.write(0, 10, 'Ura v letu', format_header4)
worksheet4.write(0, 11, 'Dan v tednu, datum, ura', format_header4)
worksheet4.write(0, 12, 'Tip elementa', format_header4)
worksheet4.write(0, 13, 'Un [kV]', format_header4)
worksheet4.write(0, 14, 'Grid', format_header4)
worksheet4.write(0, 15, 'Area', format_header4)
worksheet4.write(0, 16, 'Zone', format_header4)

#Zapis sheeta 4 - napetostimaxcasovno

worksheet5 = workbook.add_worksheet('Nap. Max Cas')
worksheet5.set_row(0, 34)
worksheet5_column_width = [20,20,20,10,10,10,10,10,10,10,10,10,10,10,10,10,10,11,6,8,8,20]
[ worksheet5.set_column(column, column, worksheet5_column_width[column]) for column in range(len(worksheet5_column_width)) ]
worksheet5.write(0, 0, 'Pravo ime', format_header1)
worksheet5.write(0, 1, 'Advanced ime', format_header1)
worksheet5.write(0, 2, 'Powerfactory ime', format_header1)
#Zapisovanje naslovov stolpcev
#Ostopanje do 2% je z zeleno, odstopanje do 4% z rumeno, nad 4% z rdečo.
[ worksheet5.write(0, i + 3, 'St. ur nad ' + str(round(float(1 + i * 0.01),2)) + ' p.u.', format_header2) for i in range(0,14) ]
worksheet5.write(0, 17, 'Tip elementa', format_header4)
worksheet5.write(0, 18, 'Un', format_header4)
worksheet5.write(0, 19, 'Grid', format_header4)
worksheet5.write(0, 20, 'Area', format_header4)
worksheet5.write(0, 21, 'Zone', format_header4)

#Zapis sheeta 5

worksheet6 = workbook.add_worksheet('Napetosti Min')
worksheet6.set_row(0, 34)
worksheet6_column_width = [20,20,20,15,7,25,15,7,25,15,7,25,11,6,8,8,20]
[ worksheet6.set_column(column, column, worksheet6_column_width[column]) for column in range(len(worksheet6_column_width)) ]
worksheet6.write(0, 0, 'Pravo ime', format_header1)
worksheet6.write(0, 1, 'Advanced ime', format_header1)
worksheet6.write(0, 2, 'Powerfactory ime', format_header1)
worksheet6.write(0, 3, 'Prva najnizja nap. v obd. [kV]', format_header2)
worksheet6.write(0, 4, 'Ura v letu', format_header2)
worksheet6.write(0, 5, 'Dan v tednu, datum, ura', format_header2)
worksheet6.write(0, 6, 'Druga najnizja nap. v obd. [kV]', format_header3)
worksheet6.write(0, 7, 'Ura v letu', format_header3)
worksheet6.write(0, 8, 'Dan v tednu, datum, ura', format_header3)
worksheet6.write(0, 9, 'Druga najnizja nap. v obd. [kV]', format_header4)
worksheet6.write(0, 10, 'Ura v letu', format_header4)
worksheet6.write(0, 11, 'Dan v tednu, datum, ura', format_header4)
worksheet6.write(0, 12, 'Tip elementa', format_header4)
worksheet6.write(0, 13, 'Un [kV]', format_header4)
worksheet6.write(0, 14, 'Grid', format_header4)
worksheet6.write(0, 15, 'Area', format_header4)
worksheet6.write(0, 16, 'Zone', format_header4)

#Zapis sheeta 6

worksheet7 = workbook.add_worksheet('Nap Min Cas')
worksheet7.set_row(0, 34)
worksheet7_column_width = [20,20,20,10,10,10,10,10,10,10,10,10,10,10,10,10,10,11,6,8,8,20]
[ worksheet7.set_column(column, column, worksheet7_column_width[column]) for column in range(len(worksheet7_column_width)) ]
worksheet7.write(0, 0, 'Pravo ime', format_header1)
worksheet7.write(0, 1, 'Advanced ime', format_header1)
worksheet7.write(0, 2, 'Powerfactory ime', format_header1)
[ worksheet7.write(0, i + 3, 'St. ur pod ' + str(round(float(1 - i * 0.01),2)) + ' p.u.', format_header2) for i in range(0,14) ]
worksheet7.write(0, 17, 'Tip elementa', format_header4)
worksheet7.write(0, 18, 'Un', format_header4)
worksheet7.write(0, 19, 'Grid', format_header4)
worksheet7.write(0, 20, 'Area', format_header4)
worksheet7.write(0, 21, 'Zone', format_header4)

# Sheet 7 podatkih generatorjev po urah

worksheet8 = workbook.add_worksheet('Gen P/Q')

worksheet8.set_row(0, 34)
worksheet8_column_width = [20,20,20]
[ worksheet8.set_column(column, column, worksheet8_column_width[column]) for column in range(len(worksheet8_column_width)) ]
worksheet8.merge_range(0, 0, 1, 0, 'Pravo ime', format_header1)
worksheet8.merge_range(0, 1, 1, 1, 'Advanced ime', format_header1)
worksheet8.merge_range(0, 2, 1, 2, 'Powerfactory ime', format_header1)

for i in hours:
    # Hocemo v formatu Ura 1, Torek, 1.1.2030
    month, day, hour = dobiDatumIzUreVLetu(year, i)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet8.merge_range(0, 3+2*i, 0, 4+2*i, 'Ura ' + str(i) + ', ' + str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', format_header1)
    #worksheet8.write(0,3+2*i, 'Ura ' + str(hours[i]) + ', ' + str(dayofweek) + str(day) + '.' + str(month) + '.' + str(year), format_header1)
    worksheet8.write(1, 3+2*i, 'P [MW]', format_header1)
    worksheet8.write(1, 4+2*i, 'Q [Mvar]', format_header1)

#Zapis sheeta 10 - nastavljene P in Q generatorjev v določenih gridih
worksheet9 = workbook.add_worksheet('Load P/Q')

worksheet9.set_row(0, 34)
worksheet9_column_width = [20,20,20]
[ worksheet9.set_column(column, column, worksheet9_column_width[column]) for column in range(len(worksheet9_column_width)) ]
worksheet9.merge_range(0, 0, 1, 0, 'Pravo ime', format_header1)
worksheet9.merge_range(0, 1, 1, 1, 'Advanced ime', format_header1)
worksheet9.merge_range(0, 2, 1, 2, 'Powerfactory ime', format_header1)

for i in hours:
    # Hocemo v formatu Ura 1, Torek, 1.1.2030
    month, day, hour = dobiDatumIzUreVLetu(year, i)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet9.merge_range(0, 3+2*i, 0, 4+2*i, 'Ura ' + str(i) + ', ' + str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', format_header1)
    #worksheet9.write(0,3+2*i, 'Ura ' + str(hours[i]) + ', ' + str(dayofweek) + str(day) + '.' + str(month) + '.' + str(year), format_header1)
    worksheet9.write(1, 3+2*i, 'P [MW]', format_header1)
    worksheet9.write(1, 4+2*i, 'Q [Mvar]', format_header1)

#Zapis sheeta 9 - porocilo kateri izracuni so konvergirali

worksheet10 = workbook.add_worksheet('Porocilo')
worksheet10.set_row(0, 34)
worksheet10_column_width = [15,20,20]
[ worksheet10.set_column(column, column, worksheet10_column_width[column]) for column in range(len(worksheet10_column_width)) ]
#Beri datoteko nekonvergence.xlsx in potem zapisi z malo lepsim formatiranjem v to datoteko
#data_konvergence = os.getcwd()  + r'/Rezultati/Raw data voltage.xlsx'
#df_results_loading_hour.to_excel(data_konvergence, sheet_name = 'DATA', index_col=0)
worksheet10.write(0, 0, 'Ura v letu', format_header1)
worksheet10.write(0, 1, 'Konvergenca?', format_header1)
worksheet10.write(0, 2, 'Potreben cas [s]', format_header1)

t3 = time.time()
dt2 = t3 - t2
print("Narejena excel osnova, cas: " + str(dt2))
print("Zacetek pisanja vrednost za daljnovode/transformatorje v excel datoteke...")

current_row = 1
for line in results_element_list_lines:
    #Pisemo samo za daljnovode z razultati
    if (current_row % 2) == 0: row_format = format_data_lighter
    else: row_format = format_data_darker
    worksheet1.set_row(current_row, 18)
    #Osnovni podatki za worksheet 1
    worksheet1.write(current_row, 0, "", row_format)
    worksheet1.write(current_row, 1, "", row_format)
    worksheet1.write(current_row, 2, line, row_format)
    
    max_no1_hour = max_no1.at[line, 'max_no1_loading']
    try: worksheet1.write(current_row, 3, df_results_line_loading_final.at[line, max_no1_hour], row_format)
    except: worksheet1.write(current_row, 3, 0, row_format)
    worksheet1.write(current_row, 4, int(max_no1_hour), row_format)
    yearhour = int(max_no1_hour) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet1.write(current_row, 5, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    max_no2_hour = max_no1.at[line, 'max_no2_loading']
    try: worksheet1.write(current_row, 6, df_results_line_loading_final.at[line, max_no2_hour], row_format)
    except: worksheet1.write(current_row, 6, 0, row_format)
    worksheet1.write(current_row, 7, int(max_no2_hour), row_format)
    yearhour = int(max_no2_hour) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet1.write(current_row, 8, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    max_no3_hour = max_no1.at[line, 'max_no3_loading']
    try: worksheet1.write(current_row, 9, df_results_line_loading_final.at[line, max_no3_hour], row_format)
    except: worksheet1.write(current_row, 9, 0, row_format)
    worksheet1.write(current_row, 10, int(max_no3_hour), row_format)
    yearhour = int(max_no3_hour) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet1.write(current_row, 11, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    worksheet1.write(current_row, 12, 'Daljnovod', row_format)
    worksheet1.write(current_row, 13, df_line_info.at[line, 'rated_voltage'], row_format)
    worksheet1.write(current_row, 14, df_line_info.at[line, 'rated_current'], row_format)
    worksheet1.write(current_row, 15, df_line_info.at[line, 'rated_power'], row_format)
    worksheet1.write(current_row, 16, df_line_info.at[line, 'grid'], row_format)
    worksheet1.write(current_row, 17, df_line_info.at[line, 'area'], row_format)
    worksheet1.write(current_row, 18, df_line_info.at[line, 'zone'], row_format)
    
    #Worksheet 2 z rezultati za % obremenitve in ure
    worksheet2.set_row(current_row, 18)
    worksheet2.write(current_row, 0, "", row_format)
    worksheet2.write(current_row, 1, "", row_format)
    worksheet2.write(current_row, 2, line, row_format)
    for i in range(0,16):
        dataframe_name = 'nad' + str(i*10)
        column_number = i + 3
        worksheet2.write(current_row, column_number, percent_loading_lines.at[line, dataframe_name], row_format)
    worksheet2.write(current_row, 19, 'Daljnovod', row_format)
    worksheet2.write(current_row, 20, df_line_info.at[line, 'rated_voltage'], row_format)
    worksheet2.write(current_row, 21, df_line_info.at[line, 'rated_current'], row_format)
    worksheet2.write(current_row, 22, df_line_info.at[line, 'rated_power'], row_format)
    worksheet2.write(current_row, 23, df_line_info.at[line, 'grid'], row_format)
    worksheet2.write(current_row, 24, df_line_info.at[line, 'area'], row_format)
    worksheet2.write(current_row, 25, df_line_info.at[line, 'zone'], row_format)
    current_row += 1

current_row += 1

for transformer in results_element_list_transformers:
    if (current_row % 2) == 0: row_format = format_data_lighter
    else: row_format = format_data_darker
    worksheet1.set_row(current_row, 18)
    worksheet1.write(current_row, 0, "", row_format)
    worksheet1.write(current_row, 1, "", row_format)
    worksheet1.write(current_row, 2, transformer, row_format)
    max_no1_hour = max_no2.at[transformer, 'max_no1_loading']
    try: worksheet1.write(current_row, 3, df_results_transformer_loading_final.at[transformer, max_no1_hour], row_format)
    except: worksheet1.write(current_row, 3, 0, row_format)
    worksheet1.write(current_row, 4, int(max_no1_hour), row_format)
    yearhour = int(max_no1_hour) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet1.write(current_row, 5, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    max_no2_hour = max_no2.at[transformer, 'max_no2_loading']
    try: worksheet1.write(current_row, 6, df_results_transformer_loading_final.at[transformer, max_no2_hour], row_format)
    except: worksheet1.write(current_row, 6, 0, row_format)
    worksheet1.write(current_row, 7, int(max_no2_hour), row_format)
    yearhour = int(max_no2_hour) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet1.write(current_row, 8, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    max_no3_hour = max_no2.at[transformer, 'max_no3_loading']
    try: worksheet1.write(current_row, 9, df_results_transformer_loading_final.at[transformer, max_no3_hour], row_format)
    except: worksheet1.write(current_row, 9, 0, row_format)
    worksheet1.write(current_row, 10, int(max_no3_hour), row_format)
    yearhour = int(max_no3_hour) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet1.write(current_row, 11, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    worksheet1.write(current_row, 12, 'Transformator', row_format)
    worksheet1.write(current_row, 13, "/", row_format)
    worksheet1.write(current_row, 14, "/", row_format)
    worksheet1.write(current_row, 15, "/", row_format)
    worksheet1.write(current_row, 16, df_transformer_info.at[transformer, 'grid'], row_format)
    worksheet1.write(current_row, 17, df_transformer_info.at[transformer, 'area'], row_format)
    worksheet1.write(current_row, 18, df_transformer_info.at[transformer, 'zone'], row_format)
    
    worksheet2.set_row(current_row, 18)
    worksheet2.write(current_row, 0, "", row_format)
    worksheet2.write(current_row, 1, "", row_format)
    worksheet2.write(current_row, 2, transformer, row_format)
    #Worksheet 2 z rezultati za % obremenitve in ure
    for i in range(0,16):
        dataframe_name = 'nad' + str(i*10)
        column_number = i + 3
        #worksheet2.write(current_row, column_number, percentload.at[transformer_name, dataframe_name])
        worksheet2.write(current_row, column_number, percent_loading_transformers.at[transformer, dataframe_name], row_format)
        #worksheet2.write(current_row, 24, line_rated_voltage)
        #worksheet2.write(current_row, 25, line_rated_current)
        #worksheet2.write(current_row, 26, line_rated_power)
    worksheet2.write(current_row, 19, 'Transformator', row_format)
    worksheet2.write(current_row, 20, "/", row_format)
    worksheet2.write(current_row, 21, "/", row_format)
    worksheet2.write(current_row, 22, "/", row_format)
    worksheet2.write(current_row, 23, df_transformer_info.at[transformer, 'grid'], row_format)
    worksheet2.write(current_row, 24, df_transformer_info.at[transformer, 'area'], row_format)
    worksheet2.write(current_row, 25, df_transformer_info.at[transformer, 'zone'], row_format)
    current_row += 1

current_row = 1
for line in results_element_list_lines:
    #Pisemo samo za daljnovode z razultati
    if (current_row % 2) == 0: row_format = format_data_lighter
    else: row_format = format_data_darker
    
    if line in df_consecutive_above_critical.index:
        for i in range(0,3):
            #Worksheet 3 z rezultati za % obremenitve in ure
            worksheet3.set_row(current_row, 18)
            worksheet3.write(current_row, 0, "", row_format)
            worksheet3.write(current_row, 1, "", row_format)
            worksheet3.write(current_row, 2, line, row_format)
            loading_value = 60+i*20
            consecituve_above_critical = df_consecutive_above_critical.at[line, loading_value]
            worksheet3.write(current_row, i*3+3, consecituve_above_critical, row_format)
            loading_hour = df_consecutive_above_start_critical.at[line, loading_value]
            if consecituve_above_critical > 0:
                worksheet3.write(current_row, i*3+4, loading_hour, row_format)
                yearhour = int(loading_hour) + 1
                month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
                dayofweek = dobiDanVtednu(year, month, day)
                worksheet3.write(current_row, i*3+5, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
            else: 
                worksheet3.write(current_row, i*3+4, '', row_format)
                worksheet3.write(current_row, i*3+5, '', row_format)
                
            worksheet3.write(current_row, 12, 'Daljnovod', row_format)
            # worksheet3.write(current_row, 13, df_line_info.at[line, 'rated_voltage'], row_format)
            # worksheet3.write(current_row, 14, df_line_info.at[line, 'rated_current'], row_format)
            # worksheet3.write(current_row, 15, df_line_info.at[line, 'rated_power'], row_format)
            # worksheet3.write(current_row, 16, df_line_info.at[line, 'grid'], row_format)
            # worksheet3.write(current_row, 17, df_line_info.at[line, 'area'], row_format)
            # worksheet3.write(current_row, 18, df_line_info.at[line, 'zone'], row_format)
        current_row += 1
    

t4 = time.time()
dt3 = t4 - t3
print("Izpisane vrednosti daljnovodov/transformatorjev v excel, cas: " + str(dt3))

current_row = 1
#Izpis za zbiralke (max in min napetosti)
#for terminal in terminals:
#for terminal in results_hourly_elementlist_voltage:
print("Zacetek pisanja rezultatov napetosti zbiralk v excel...")
for terminal in results_element_list_terminals:
    if (current_row % 2) == 0: row_format = format_data_lighter
    else: row_format = format_data_darker
    worksheet4.set_row(current_row, 18)
    terminal_rated_voltage = 0
    terminal_grid = df_terminal_info.at[terminal, 'grid']
    terminal_area = df_terminal_info.at[terminal, 'area']
    terminal_zone = df_terminal_info.at[terminal, 'zone']
        #Osnovni podatki za worksheet 3
    worksheet4.write(current_row, 0, "", row_format)
    worksheet4.write(current_row, 1, "", row_format)
    worksheet4.write(current_row, 2, terminal, row_format)
    
    # Za dobit column z max vrednostjo https://www.skytowner.com/explore/getting_column_label_of_max_value_in_each_row_in_pandas_datafrme
    max_no1_volt = max_no3.at[terminal, 'max_no1_voltage']
    #print(max_no1_hour)
    try: worksheet4.write(current_row, 3, round((df_results_terminal_voltage_final.at[terminal, max_no1_volt] * terminal_rated_voltage), 1), row_format)
    except: worksheet4.write(current_row, 3, 0, row_format)
    worksheet4.write(current_row, 4, int(max_no1_volt), row_format)
    yearhour = int(max_no1_volt) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet4.write(current_row, 5, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
        
    max_no2_volt = max_no3.at[terminal, 'max_no2_voltage']
    try: worksheet4.write(current_row, 6, round((df_results_terminal_voltage_final.at[terminal, max_no2_volt] * terminal_rated_voltage), 1), row_format)
    except: worksheet4.write(current_row, 6, 0, row_format)
    worksheet4.write(current_row, 7, int(max_no2_volt), row_format)
    yearhour = int(max_no2_volt) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet4.write(current_row, 8, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    max_no3_volt = max_no3.at[terminal, 'max_no3_voltage']
    try: worksheet4.write(current_row, 9, round((df_results_terminal_voltage_final.at[terminal, max_no3_volt] * terminal_rated_voltage), 1), row_format)
    except: worksheet4.write(current_row, 9, 0, row_format)
    worksheet4.write(current_row, 10, int(max_no3_volt), row_format)
    yearhour = int(max_no3_volt) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet4.write(current_row, 11, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    worksheet4.write(current_row, 12, "Zbiralka", row_format)
    worksheet4.write(current_row, 13, terminal_rated_voltage, row_format)
    worksheet4.write(current_row, 14, df_terminal_info.at[terminal, 'grid'], row_format)
    worksheet4.write(current_row, 15, df_terminal_info.at[terminal, 'area'], row_format)
    worksheet4.write(current_row, 16, df_terminal_info.at[terminal, 'zone'], row_format)
    
    #Worksheet 2 z rezultati za % obremenitve in ure
    worksheet5.set_row(current_row, 18)
    worksheet5.write(current_row, 0, "", row_format)
    worksheet5.write(current_row, 1, "", row_format)
    worksheet5.write(current_row, 2, terminal, row_format)
    for i in range(0,14):
        overvoltagepercent = i * 0.01
        endvoltage = float(1 + overvoltagepercent)
        dataframe_name = 'nad' + str(endvoltage)
        column_number = i + 3
        worksheet5.write(current_row, column_number, voltage_high_terminals.at[terminal, dataframe_name], row_format)
    worksheet5.write(current_row, 17, "Zbiralka", row_format)
    worksheet5.write(current_row, 18, terminal_rated_voltage, row_format)
    worksheet5.write(current_row, 19, df_terminal_info.at[terminal, 'grid'], row_format)
    worksheet5.write(current_row, 20, df_terminal_info.at[terminal, 'area'], row_format)
    worksheet5.write(current_row, 21, df_terminal_info.at[terminal, 'zone'], row_format)
    
    #Osnovni podatki za worksheet 1
    worksheet6.set_row(current_row, 18)
    worksheet6.write(current_row, 0, "", row_format)
    worksheet6.write(current_row, 1, "", row_format)
    worksheet6.write(current_row, 2, terminal, row_format)
    min_no1_volt = max_no3.at[terminal, 'min_no1_voltage']
    #print(max_no1_hour)
    try: worksheet6.write(current_row, 3, round((df_results_terminal_voltage_final.at[terminal, min_no1_volt] * terminal_rated_voltage), 1), row_format)
    except: worksheet6.write(current_row, 3, 0, row_format)
    worksheet6.write(current_row, 4, int(min_no1_volt), row_format)
    yearhour = int(min_no1_volt) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet6.write(current_row, 5, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
        
    # Za dobit column z max vrednostjo https://www.skytowner.com/explore/getting_column_label_of_max_value_in_each_row_in_pandas_datafrme
    min_no2_volt = max_no3.at[terminal, 'min_no2_voltage']
    try: worksheet6.write(current_row, 6, round((df_results_terminal_voltage_final.at[terminal, min_no2_volt] * terminal_rated_voltage), 1), row_format)
    except: worksheet6.write(current_row, 6, 0, row_format)
    worksheet6.write(current_row, 7, int(min_no2_volt), row_format)
    yearhour = int(min_no2_volt) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet6.write(current_row, 8, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    
    min_no3_volt = max_no3.at[terminal, 'min_no3_voltage']
    try: worksheet6.write(current_row, 9, round((df_results_terminal_voltage_final.at[terminal, min_no3_volt] * terminal_rated_voltage), 1), row_format)
    except: worksheet6.write(current_row, 9, 0, row_format)
    worksheet6.write(current_row, 10, int(min_no3_volt), row_format)
    yearhour = int(min_no3_volt) + 1
    month, day, hour = dobiDatumIzUreVLetu(year, yearhour)
    dayofweek = dobiDanVtednu(year, month, day)
    worksheet6.write(current_row, 11, str(dayofweek) + ', ' + str(day) + '.' + str(month) + '.' + str(year) + ', ' + str(hour) + ':00', row_format)
    worksheet6.write(current_row, 12, "Zbiralka", row_format)
    worksheet6.write(current_row, 13, terminal_rated_voltage, row_format)
    worksheet6.write(current_row, 14, terminal_grid, row_format)
    worksheet6.write(current_row, 15, terminal_area, row_format)
    worksheet6.write(current_row, 16, terminal_zone, row_format)
    
    worksheet7.set_row(current_row, 18)
    worksheet7.write(current_row, 0, "", row_format)
    worksheet7.write(current_row, 1, "", row_format)
    worksheet7.write(current_row, 2, terminal, row_format)
    for i in range(0,14):
        undervoltagepercent = i * 0.01
        endvoltage = float(1 - undervoltagepercent)
        dataframe_name = 'pod' + str(endvoltage)
        column_number = i + 3
        worksheet7.write(current_row, column_number, voltage_low_terminals.at[terminal, dataframe_name], row_format)
    worksheet7.write(current_row, 17, "Zbiralka", row_format)
    worksheet7.write(current_row, 18, terminal_rated_voltage, row_format)
    worksheet7.write(current_row, 19, terminal_grid, row_format)
    worksheet7.write(current_row, 20, terminal_area, row_format)
    worksheet7.write(current_row, 21, terminal_zone, row_format)
    
    current_row += 1
        
t5 = time.time()
dt4 = t5 - t4
print("Izpisane vrednosti zbiralk v excel, cas: " + str(dt4))
################################OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
current_row = 2

        #Izpis za generatorje, delovne in jalove nastavljene moči
print("Zacetek pisanaja dodeljenih moci generatorjev v excel...")
for generator in results_element_list_generators:
    if (current_row % 2) == 0: row_format = format_data_lighter
    else: row_format = format_data_darker
    worksheet8.set_row(current_row, 18)
    worksheet8.write(current_row, 0, "", row_format)
    worksheet8.write(current_row, 1, "", row_format)
    worksheet8.write(current_row, 2, generator, row_format)
    for i in range(len(hours)):
        try:
            worksheet8.write(current_row, 3+2*i, round(df_results_generator_P_set_final.at[generator, str(hours[i])], 2), row_format)
            worksheet8.write(current_row, 4+2*i, round(df_results_generator_Q_set_final.at[generator, str(hours[i])], 2), row_format)
            # print("Zapisal")
        except:
            worksheet8.write(current_row, 3+2*i, round(0, 2), row_format)
            worksheet8.write(current_row, 4+2*i, round(0, 2), row_format)
            # print("Ni slo...")
            
    current_row += 1
        
# Zzpisemo delovne in jalove moči po urah za bremena
        
print("Konc pisanja vrednosti generatorjev")
print("Zacetek pisanja nastavljenih vrednosti P in Q generatorjev pu urah")
current_row = 2
for load in results_element_list_loads:
    if (current_row % 2) == 0: row_format = format_data_lighter
    else: row_format = format_data_darker
    worksheet9.set_row(current_row, 18)
    worksheet9.write(current_row, 0, "", row_format)
    worksheet9.write(current_row, 1, "", row_format)
    worksheet9.write(current_row, 2, load, row_format)
    for i in range(len(hours)):
        try:
            worksheet9.write(current_row, 3+2*i, round(df_results_load_P_set_final.at[load, str(hours[i])], 2), row_format)
            worksheet9.write(current_row, 4+2*i, round(df_results_load_Q_set_final.at[load, str(hours[i])], 2), row_format)
        except:
            worksheet9.write(current_row, 3+2*i, round(0, 2), row_format)
            worksheet9.write(current_row, 4+2*i, round(0, 2), row_format)
    current_row += 1
        
print("Konec pisanja vrednosti generatorjev po urah")
print("Izpisovanje statusov in casa izracunov")

# # Izpisemo rezultate loadflov kalkulacij na sheet 9
# current_row = 1
# for hour in df_results_calculation_status_joined.index:
#     if (current_row % 2) == 0:
#         row_format = format_data_lighter
#     else:
#         row_format = format_data_darker
#     worksheet10.set_row(current_row, 18)
#     worksheet10.write(current_row, 0, hour, row_format)
#     convergence_status = df_results_calculation_status_joined.at[hour,'convergence']
#     if convergence_status == 0:
#         worksheet10.write(current_row, 1, 'DA', row_format)
#     else:
#         worksheet10.write(current_row, 1, 'NE', row_format)
#     worksheet10.write(current_row, 2, df_results_calculation_status_joined.at[hour,'calculation_time'], row_format)
#     current_row += 1
#     # df_results_calcstatus.at[hour,'convergence'] = int(status)
#     # df_results_calcstatus.at[hour,'calculation_time'] = int(t_calc)
print("Izpisovanje datoteke z rezultati končano")
workbook.close()
