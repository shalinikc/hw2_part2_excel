# -*- coding: utf-8 -*-
"""
Created on Sun Jun  5 15:27:35 2022

@author: Shalini
"""

import openpyxl
from openpyxl import Workbook
import csv
import os
from glob import glob
from pathlib import Path
import pandas as pd

file_path = './output/extra credit 1/'

csvs_path = Path.cwd() / 'data' / 'logs'

# Loop over all the csv files 
for csv_fn in csvs_path.glob('*.csv'):
    # Split the filename off from csv extension. We'll use the filename
    # (without the extension) as the key in the dfs dict.
    print(csv_fn)
    fstem = csv_fn.stem
    stream_name = fstem.split('-')[0]

    # Read the next csv file into a pandas DataFrame and add it to the dfs dictionary.
    df = pd.read_csv(csv_fn)
    df.columns =['Datetime', 'Temperature scale', 'Temperature']
        
    #print(df)
    
    file_name = file_path+stream_name+'.xlsx'
    print(file_name)
    writer = pd.ExcelWriter(file_name, engine='openpyxl',mode='a', if_sheet_exists="replace")
    
    # Write your DataFrame to a file   
    df.to_excel(writer, fstem,index=False)
    
    
    # if os.path.isfile(file_name):
    #     print('existing file')
    #     #book=open_workbook(file_name)
    #     book = openpyxl.load_workbook(filename = file_name)
    # else:
    #     print('new file')
    #     workbook2=xlwt.Workbook(file_name)
    #     ws = workbook2.add_sheet('Tested')
    #     workbook2.save(fname2)
    #     book = open_workbook(fname2)
    
    
# Save the result
writer.save()