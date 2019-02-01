import pandas as pd
#import numpy as np
import os
import BEC00760_Non_Domestic
import sys
import xlwings as xw
import win32com.client


def main():

    path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/BEC 2018/')
    file_name = 'BEC 00760_ EXAMPLE EXTRACT FIELDS.xlsm'
    BEC_file = pd.ExcelFile(path + file_name)

    BEC_sheet = {}
    for sheet in BEC_file.sheet_names:
        if ('Non Domestic' in sheet):
            BEC_sheet[sheet] = BEC_file.parse(sheet)
    print 'Done!'

    print BEC_sheet['Non Domestic 3']
if __name__=='__main__':
    main()