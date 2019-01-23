import pandas as pd
#import numpy as np
import os
import BEC00760_Non_Domestic
import sys
import win32com.client

path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

class BEC00760(object):
    def __init__(self,file):
        self.bec00760_file = pd.ExcelFile(path+file)
        self.BEC00760_worksheet={}
        for sheetName in self.bec00760_file.sheet_names:
            if ('Non Domestic' in sheetName):
                self.BEC00760_worksheet[sheetName] = BEC00760_Non_Domestic(self.bec00760_file,sheetName)
            elif ('Project Summary' == sheetName or 'Beneficiary' == sheetName):
                self.BEC00760_worksheet[sheetName] = self.bec00760_file.parse(sheetName)

    def print_list_sheet(self):
        print self.bec00760_file.sheet_names

    def print_each_sheet(self):
        for sheetName in self.bec00760_file.sheet_names:
            if ('Non Domestic' in sheetName):
                self.BEC00760_worksheet[sheetName].print_sheet_content()
            elif ('Project Summary' == sheetName or 'Beneficiary' == sheetName):
                print self.BEC00760_worksheet[sheetName]


class BEC00760_Non_Domestic(object):
    def __init__(self,bec00760_file,sheetName):
        self.fileName = 'BEC00760'
        self.sheetName= sheetName
        #self.sheet = bec00760_file.parse(sheetName)
        self.sheet = pd.read_excel(bec00760_file,sheetName,keep_default_na =False,header=None)

    def print_sheet_content(self):
        print 'File name: ',self.fileName
        print 'Sheet name: ',self.sheetName
        print self.sheet

def unprotect_xlsx(path,filename):
    xcl = win32com.client.Dispatch('Excel.Application')
    pw_str = 'Bec2018dec2017'
    wb = xcl.Workbooks.Open(path+filename,False,True,None,pw_str)
    xcl.DisplayAlerts=False
    wb.SaveAs(filename,None,'','')
    xcl.Quit()

def main():
    file_name='BEC 00760_ EXAMPLE EXTRACT FIELDS.xlsm'
    temp_file = BEC00760(file_name)
    unprotect_xlsx(path, file_name)
    #temp_file.print_list_sheet()
    #temp_file.print_each_sheet()
    print temp_file.BEC00760_worksheet['Non Domestic 1'].sheet
    print 'Done!'

if __name__=='__main__':
    main()
