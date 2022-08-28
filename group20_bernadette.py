# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
#for bar charts
import matplotlib as mlp
import matplotlib.pyplot as plt


# import BarChart class from openpyxl.chart sub_module

from datetime import datetime

# Importing the data from file
# if to specify the any sheet of excel

#EUROPEAN COUNTRIES
from openpyxl import Workbook, load_workbook
wb = load_workbook('avia_tf_cm_spreadsheet.xlsx')
wb.active = wb["Sheet 1"]
ws = wb.active

values = []
number_of_flights = list(ws.rows)
for cell in number_of_flights[8]:
    values.append(cell.value)

values = values[2:]

while '' in values:
    values.remove('')

#WESTERN & EASTERN COUNTRIES
#import dataset
df = pd.read_excel("avia_tf_cm_spreadsheet.xlsx", sheet_name=2) #, header= None
rows, columns = df.shape

#define time, but actually not used 
time = df.loc[5,:]
time2=time.dropna(axis='rows')
time2=time2[2:]

df.replace('Germany (until 1990 former territory of the FRG)', 'Germany', inplace=True)
df.rename(columns={'Data extracted on 06/07/2022 20:44:02 from [ESTAT]': 'Countries', 'Unnamed: 1': 'Type'}, inplace=True)
#df.set_index('Countries', inplace=True) #wenn ich das tue, dann kann ich den Filter nicht mehr anwenden

#list with western countries
western_countries = ['Austria', 'Belgium', 'France', 'Germany', 'Luxembourg', 'Netherlands', 'Switzerland']
eastern_countries = ['Bulgaria', 'Czechia', 'Hungary', 'Poland', 'Romania', 'Slovakia']

#percentage change
filt_western_pc = (df['Countries'].isin(western_countries)) & (df['Type'] == 'Percentage change compared to same month in 2019')
df_westerncountries_pc = df.loc[filt_western_pc]

filt_eastern_pc = (df['Countries'].isin(eastern_countries)) & (df['Type'] == 'Percentage change compared to same month in 2019')
df_easterncountries_pc = df.loc[filt_eastern_pc]

#delete nan values 
df_westerncountries_pc = df_westerncountries_pc.dropna(axis ='columns',)
df_easterncountries_pc = df_easterncountries_pc.dropna(axis ='columns',)

#percentage change from dec2019 to dec2021
pc_to_2019_western = list( df_westerncountries_pc['Unnamed: 72']) #unnamed 72 is december 2021
countries_western = list( df_westerncountries_pc['Countries'])

pc_to_2019_eastern = list( df_easterncountries_pc['Unnamed: 72']) #unnamed 72 is december 2021
countries_eastern = list( df_easterncountries_pc['Countries'])

#creating figure western countries
plt.rcParams["figure.figsize"] = [10, 8]
plt.rcParams["figure.autolayout"] = True
fig1 = plt.bar(countries_western, pc_to_2019_western, color=('blue'), label = 'Western European countries')
fig2 = plt.bar(countries_eastern, pc_to_2019_eastern, color=('red'), label = 'Eastern European countries')
plt.title("Percentage Change of flights in Western & Eastern Europe between Dec2019 & Dec2021", fontsize=13)
plt.xlabel("Countries", fontsize=13)

plt.tick_params(axis='both', which='major', labelsize=7)
plt.xticks(fontsize=13, rotation=45)
plt.yticks(fontsize=13)
plt.ylabel("Percentage Change of number of flights", fontsize=13)
plt.legend(fontsize=13)

plt.draw()
plt.savefig('barplot_percentage_change__eastern_western_EU_2019_to_2021.png', dpi = 700)
plt.show()

#absolute number of flights
#western countries herausfiltern
filt = (df['Countries'].isin(western_countries)) & (df['Type'] == 'Number')
df_westerncountries = df.loc[filt]

#delete nan values 
df_westerncountries = df_westerncountries.dropna(axis ='columns',)

#sum of all western countries flights per month
number_of_flights_western = list(df_westerncountries.sum(axis=0))
number_of_flights_western = number_of_flights_western[2:]

#Western flights
flights_western = list(map(int, number_of_flights_western))
#European flights
fligths = list(map(int, values))

# month and year
times = []
month_year_of_flight = list(ws.rows)
for cell in month_year_of_flight[6]:
    times.append(cell.value)

times= times[1:]

while None in times:
    times.remove(None)
    
new_list = []
list_2 = []
for i in times:
    new_list.append(datetime.strptime(i, '%Y-%m').date())
    
# putting in the correct format of month and year   
for i in new_list:
    list_2.append(datetime.strftime(i, '%b-%y'))


#creating figure
#plt.rcParams["figure.figsize"] = [6, 4]
#plt.rcParams["figure.autolayout"] = True
#fig1 = plt.bar(list_2, fligths, color=('#7700EE'))
#fig2 = plt.bar(list_2, flights_western, color=('red'))
#plt.title("Commercial Flights in Western EU (Jan 2019- May 2022)")
#plt.xlabel("Month & Year")

#plt.tick_params(axis='both', which='major', labelsize=7)
#plt.xticks(fontsize=7, rotation=90)
#plt.yticks(fontsize=9)
#plt.ylabel("Number of Flights")

#plt(fig1)
#plt(fig2)

#Saving plot with high resolution
#plt.draw()
#plt.savefig('barplot_EU_Flights_2019_to_2022.png', dpi = 700)
#plt.show()


