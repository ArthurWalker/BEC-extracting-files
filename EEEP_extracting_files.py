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


path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

def execute_each_folder(eep_path,folder_name)
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