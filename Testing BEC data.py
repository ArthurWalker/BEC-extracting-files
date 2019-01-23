import pandas as pd
#import numpy as np
import os
import BEC00760_Non_Domestic
import sys
import xlwings as xw
import win32com.client

def unprotect_xlsx(path,filename):
    xcl = win32com.client.Dispatch('Excel.Application')
    pw_str = 'Bec2018dec2017'
    wb = xcl.Workbooks.Open(path+filename,False,True,None,pw_str)
    xcl.DisplayAlerts=False
    wb.SaveAs(filename,None,'','')
    xcl.Quit()

def main():

    path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')
    file_name = 'BEC 00760_ EXAMPLE EXTRACT FIELDS.xlsm'
    unprotect_xlsx(path,file_name)
    BEC_file = pd.ExcelFile(path + file_name)

    BEC_sheet = {}
    for sheet in BEC_file.sheet_names:
        if ('Non Domestic' in sheet):
            BEC_sheet[sheet] = BEC_file.parse(sheet)
    print 'Done!'

    print BEC_sheet['Non Domestic 3']
if __name__=='__main__':
    main()