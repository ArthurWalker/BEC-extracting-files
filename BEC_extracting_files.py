import pandas as pd
import numpy as np
import os
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
                self.BEC00760_worksheet[sheetName] = pd.read_excel(self.bec00760_file,sheetName,keep_default_na =False,header=None)

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
        self.sheet = pd.read_excel(bec00760_file,sheetName,keep_default_na =False,header=None).dropna(thresh=1)
        self.data_site_reference = ''
        self.data_site_measures = ''

    def print_input_sheet_content(self):
        print 'File name: ',self.fileName
        print 'Sheet name: ',self.sheetName
        print self.sheet

    def extract_data_from_input_sheet(self):
        self.data_site_reference = self.sheet.iloc[2:14,0:3]
        TEMP_data_site_measures =self.sheet.iloc[14:,0:24]
        self.data_site_measures = TEMP_data_site_measures.loc[~TEMP_data_site_measures[0].isin(['Total','','-'])]
        return self.data_site_measures,self.data_site_reference

    def print_output_sheet_content(self):
        print ''
        print 'Data of site reference: '
        print self.data_site_reference
        print ''
        print 'Data of site measures: '
        print self.data_site_measures

    def write_csv_file(self):
        output_reference_filename= 'BEC_Site_Reference.csv'
        output_measures_filename= 'BEC_Site_Measures.csv'
        if not (os.path.isfile(path+output_reference_filename)):
            self.data_site_reference.to_csv(path_or_buf=output_reference_filename,index=None,header=False)
        if not (os.path.isfile(path+output_measures_filename)):
            self.data_site_measures.to_csv(path_or_buf=output_measures_filename,index=None,header=False)


def unprotect_xlsm_file(path,filename):
    xcl = win32com.client.Dispatch('Excel.Application')
    pw_str = 'Bec2018dec2017'
    wb = xcl.Workbooks.Open(path+filename,False,True,None,pw_str)
    xcl.DisplayAlerts=False
    wb.SaveAs(filename,None,'','')
    xcl.Quit()

def main():
    file_name='BEC 00760_ EXAMPLE EXTRACT FIELDS.xlsm'
    unprotect_xlsm_file(path, file_name)
    temp_file = BEC00760(file_name)

    non_domestic_1 = temp_file.BEC00760_worksheet['Non Domestic 1']
    non_domestic_1.print_input_sheet_content()
    non_domestic_1.extract_data_from_input_sheet()
    non_domestic_1.print_output_sheet_content()
    #non_domestic_1.write_csv_file()

    print 'Done!'

if __name__=='__main__':
    main()
