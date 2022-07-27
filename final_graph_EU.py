# -*- coding: utf-8 -*-
"""
Created on Wed Jul 27 19:02:08 2022

@author: Muhammad Affan Qamar
"""

import pandas as pd
#for bar chats
import matplotlip as mlp
import matplotlib.pyplot as plt

# import BarChart class from openpyxl.chart sub_module

from datetime import datetime

# Importing the data from file
# if to specify the any sheet of excel
#df = pd.read_excel (r'avia_tf_cm_spreadsheet.xlsx', sheet_name='Sheet 1')
from openpyxl import Workbook, load_workbook
# workbook = wb
wb = load_workbook('avia_tf_cm_spreadsheet.xlsx')
# ws = worksheet
wb.active = wb["Sheet 1"]
ws = wb.active

# number of flights
values = []
number_of_flights = list(ws.rows)
for cell in number_of_flights[8]:
    values.append(cell.value)
    
values = values[2:]

while '' in values:
    values.remove('')
    
flights = list(map(int, values))


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
plt.rcParams["figure.figsize"] = [6, 4]
plt.rcParams["figure.autolayout"] = True
fig1 = plt.bar(list_2, flights, color=('#7700EE'))
plt.title("Commercial Flights in EU (Jan 2019- May 2022)")
plt.xlabel("Month & Year")

#plt.tick_params(axis='both', which='major', labelsize=7)
plt.xticks(fontsize=7, rotation=90)
plt.yticks(fontsize=9)
plt.ylabel("Number of Flights")

#Saving plot with high resolution
plt.draw()
#for .png format
plt.savefig('barplot_EU_Flights_2019_to_2022.png', dpi = 700)
#for pdf
#plt.savefig('barplot_EU_Flights_2019_to_2022.pdf')
plt.show()