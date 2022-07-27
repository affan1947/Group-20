# -*- coding: utf-8 -*-
"""
Created on Thu Jun 30 19:00:42 2022

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


    
print(dir(wb))