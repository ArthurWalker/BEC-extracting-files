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


def execute_each_file(new_path,file):
    input_folder = new_path + file + '/'
    project_name = re.search(r'\w+\s+\w+', file).group()
    project_year = re.search(r'\d+', file).group()
    excel_file = pd.ExcelFile(input_folder + file)
    if 'Admin' in excel_file.sheet_names[0]:
        lst_col_admin = ['Reference No.','Cat. ','Submitted By','Project Title','County','Approved Funding']
        df= pd.read_excel(excel_file,keep_default_na=False,header=None,skiprows=1)
        series = df.iloc[0]
        col_list = series[series.isin(lst_col_admin)].index.tolist()
        return df[col_list]

def execute_each_folder(eep_path,folder_name):
    new_path = eep_path+folder_name+'/'
    file_lst = os.listdir(new_path)
    for file in file_lst:
        if re.search(r'Summary',file):
            df = execute_each_file(new_path,file)
    return

def main():
    start_time = time.time()
    path_lst = os.listdir(path)
    if 'SEEEP' in path_lst:
        seeep_path = path+'SEEEP/'
        folder = os.listdir(seeep_path)
        for folder_name in folder:
            if re.search(r'EE',folder_name):
                execute_each_folder(seeep_path,folder_name)
    print('Done! from ', time.asctime(time.localtime(start_time)), ' to ',time.asctime(time.localtime(time.time())))

if __name__ == '__main__':
    main()