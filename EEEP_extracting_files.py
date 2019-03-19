# -*- coding: utf-8 -*-
import tempfile
import pandas as pd
import numpy as np
import re
import os
import sys
import win32com.client
import time
from tqdm import tqdm
import xlwings as xw
import openpyxl
from openpyxl import load_workbook
from fuzzywuzzy import fuzz
import msoffcrypto
import xlrd

from BEW_extracting_files import *
import BEW_extracting_files as bew

path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

def execute_each_file_Other(new_path,file):
    input_file = new_path+file
    excel_file = pd.ExcelFile(input_file)
    new_path = '/'.join(new_path.split('/')[:-2]) + '/'
    # First tab
    df_project = pd.read_excel(excel_file,excel_file.sheet_names[0],header = None,keep_default_na=False,skiprows=1)
    bew.write_file(new_path, '', df_project, 'SEEEP Project and Technology Summary January 2010_Project')
    # Second tab
    df_energy = pd.read_excel(excel_file,excel_file.sheet_names[1],header = None,keep_default_na=False,skiprows=2,usecols=[0,11])
    bew.write_file(new_path, '', df_energy, 'SEEEP Project and Technology Summary January 2010_Energy')

def execute_each_file_Stats(new_path,file):
    input_folder = new_path + file
    excel_file = pd.ExcelFile(input_folder)
    if 'Admin' in excel_file.sheet_names[0]:
        lst_col_admin = ['Reference No.','Cat. ','Cat. No.','Submitted By','Project Title','County','Approved Funding']
        df= pd.read_excel(excel_file,keep_default_na=False,header=None,skiprows=1)
        series = df.iloc[0]
        col_list = series[series.isin(lst_col_admin)].index.tolist()
        return df[col_list]

def execute_each_folder(eep_path,folder_name,project_year):
    new_path = eep_path+folder_name+'/'
    file_lst = os.listdir(new_path)
    for file in file_lst:
        if re.search(r'Statistical',file):
            df = execute_each_file_Stats(new_path,file)
            new_path = '/'.join(new_path.split('/')[:-2])+'/'
            df.insert(0, '', project_year)
            df.iloc[0, 0] = 'Year'
            bew.write_file(new_path,'',df,'Admin')
        # else:
        #     execute_each_file_Other(new_path,file)


def main():
    #start_time = time.time()
    path_lst = os.listdir(path)
    if 'SEEEP' in path_lst:
        seeep_path = path+'SEEEP/'
        folder = os.listdir(seeep_path)
        for folder_name in folder:
            if re.search(r'EE',folder_name):
                project_year = re.search(r'\d+', folder_name).group()
                execute_each_folder(seeep_path,folder_name,project_year)
    #print('Done! from ', time.asctime(time.localtime(start_time)), ' to ',time.asctime(time.localtime(time.time())))

if __name__ == '__main__':
    main()