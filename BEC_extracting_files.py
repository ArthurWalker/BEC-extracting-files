import pandas as pd
import numpy as np
import os
import sys
import win32com.client
from pandas import ExcelWriter
path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

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
        small_df = pd.concat([self.sheet.loc[11:13,3],self.sheet.loc[11:13,2]],axis=1)
        small_df = small_df.transpose().reset_index(drop=True).transpose()
        self.data_site_reference = self.sheet.iloc[2:14,0:2].append(small_df,ignore_index=True)
        TEMP_data_site_measures_proposed_energy_upgrades =self.sheet.iloc[15:25,0:7].reset_index(drop=True)
        TEMP_data_site_measures_unit = self.sheet.iloc[14:25, 7:24].drop(15,axis=0).reset_index(drop=True)
        TEMP_data_site_measures = pd.concat([TEMP_data_site_measures_proposed_energy_upgrades,TEMP_data_site_measures_unit],axis=1)
        self.data_site_measures = TEMP_data_site_measures.loc[~TEMP_data_site_measures[0].isin(['Total','','-'])]
        return self.data_site_measures,self.data_site_reference

    def print_output_sheet_content(self):
        print ''
        print 'Data of site reference: '
        print self.data_site_reference
        print ''
        print 'Data of site measures: '
        print self.data_site_measures

class BEC00760(object):
    def __init__(self,file):
        self.bec00760_file = pd.ExcelFile(path+file)
        self.BEC00760_worksheet={}
        self.project_summary_dataframe = ''
        self.beneficiary_dataframe = ''
        self.site_references = ''
        self.site_measures = ''
        for sheetName in self.bec00760_file.sheet_names:
            if ('Non Domestic' in sheetName):
                self.BEC00760_worksheet[sheetName] = BEC00760_Non_Domestic(self.bec00760_file,sheetName)
            elif ('Project Summary' == sheetName or 'Beneficiary' == sheetName):
                self.BEC00760_worksheet[sheetName] = pd.read_excel(self.bec00760_file,sheetName,keep_default_na =False,header=None)

    def print_list_sheet(self):
        print self.bec00760_file.sheet_names

    def print_original_sheet(self):
        for sheetName in self.bec00760_file.sheet_names:
            if ('Non Domestic' in sheetName):
                self.BEC00760_worksheet[sheetName].print_input_sheet_content()
            elif ('Project Summary' == sheetName or 'Beneficiary' == sheetName):
                print 'File name: ', 'BEC00760'
                print 'Sheet name: ', sheetName
                print self.BEC00760_worksheet[sheetName]

    def extract_summary_data(self):
        TEMP_dataframe = self.BEC00760_worksheet['Project Summary'].iloc[86:,1]
        list_Add_addition_row = TEMP_dataframe[TEMP_dataframe=='Add additional rows as required'].index.tolist()
        if (len(list_Add_addition_row)==1):
            TEMP_data_project_summary1 = self.BEC00760_worksheet['Project Summary'].iloc[86:list_Add_addition_row[0], 1:6].reset_index(drop=True).drop(3,axis=1)
            TEMP_data_project_summary2 = self.BEC00760_worksheet['Project Summary'].iloc[84:list_Add_addition_row[0],18:21].drop([85,86],axis=0).reset_index(drop=True)
            data_project_summary = pd.concat([TEMP_data_project_summary1, TEMP_data_project_summary2], axis=1)
            data_project_summary.insert(0,'1',[i for i in range(data_project_summary.shape[0])])
            data_project_summary.insert(0,'0','BEC00760')
            data_project_summary.iloc[0,0]='Project Code'
            data_project_summary.iloc[0,1]='ID'
            self.project_summary_dataframe=data_project_summary
        else:
            print 'Can not identify as there are more "Add additional rows as required" or no results'

    def extract_beneficiary_data(self):
        TEMP_data_beneficiary = self.BEC00760_worksheet['Beneficiary'].iloc[8:,1]
        data_beneficiary = TEMP_data_beneficiary.loc[~TEMP_data_beneficiary.isin(['Total Project Cost',''])].to_frame().reset_index(drop=True)
        data_beneficiary.insert(0,0,'BEC00760')
        data_beneficiary.iloc[0,0]='Project Code'
        self.beneficiary_dataframe = data_beneficiary

    def extract_non_domestic_data(self):
        non_domestic_list = [i for i in self.BEC00760_worksheet.keys() if 'Non Domestic' in i]
        list_measures = []
        list_reference = []
        for non_domestic_sheet in non_domestic_list:
        # Non Domestic Measures
            non_domestic_measures = self.BEC00760_worksheet[non_domestic_sheet].extract_data_from_input_sheet()[0]
            non_domestic_measures.insert(0, '2', [i for i in range(non_domestic_measures.shape[0])])
            if len(list_measures)>0:
                non_domestic_measures=non_domestic_measures.drop(0,axis=0)
            non_domestic_measures.insert(0, '1', non_domestic_sheet)
            list_measures.append(non_domestic_measures)
        # Non Domestic Reference
            non_domestic_reference= self.BEC00760_worksheet[non_domestic_sheet].extract_data_from_input_sheet()[1].transpose()
            non_domestic_reference.insert(0, '2', [i for i in range(non_domestic_reference.shape[0])])
            if len(list_reference)>0:
                non_domestic_reference=non_domestic_reference.drop(0,axis=0)
            non_domestic_reference.insert(0, '1', non_domestic_sheet)
            list_reference.append(non_domestic_reference)
    #Non Domestic Measures
        self.site_measures = pd.concat(list_measures,ignore_index=True)
        self.site_measures.insert(0, '0', 'BEC00760')
        self.site_measures.iloc[0,0]='Project Code'
        self.site_measures.iloc[0,1]='Tab'
        self.site_measures.iloc[0,2]='ID Measure'
    #Non Domestic Reference
        self.site_references = pd.concat(list_reference,ignore_index=True)
        self.site_references.insert(0, '0', 'BEC00760')
        self.site_references.iloc[0,0]='Project Code'
        self.site_references.iloc[0,1]='Tab'
        self.site_references.iloc[0,2]='ID Reference'

    def extract_data(self):
        self.extract_summary_data()
        self.extract_beneficiary_data()
        self.extract_non_domestic_data()
        print 'Data outputs are available'

    def check_available_result(self):
        if (self.project_summary_dataframe.shape[0]>0 and self.beneficiary_dataframe.shape[0]>0 and self.site_references.shape[0]>0 and self.site_measures.shape[0]>0):
            return True
        else:
            return False

    def print_output_sheets(self):
        self.project_summary_dataframe = ''
        self.beneficiary_dataframe = ''
        self.site_references = ''
        self.site_measures = ''
        if self.check_available_result():
            print 'Project summary',self.project_summary_dataframe
            print 'Beneficiary', self.beneficiary_dataframe
            print 'Site references', self.site_references
            print 'Site measures', self.site_measures
        else:
            print 'Need to run extract_data() to execute input file to have results'

    def write_csv_file(self):
        self.project_summary_dataframe.to_excel(path+'BEC00760_Project Summary.xlsx','Project Summary',header=False,index=False)
        self.beneficiary_dataframe.to_excel(path+'BEC00760_Beneficiary.xlsx','Beneficiary',header=False,index=False)
        self.site_references.to_excel(path+'BEC00760_References.xlsx','References',header=False,index=False)
        self.site_measures.to_excel(path+'BEC00760_Measures.xlsx','Measures',header=False,index=False)

def unprotect_xlsm_file(path,filename):
    xcl = win32com.client.Dispatch('Excel.Application')
    pw_str = 'Bec2018dec2017'
    wb = xcl.Workbooks.Open(path+filename,False,True,None,pw_str)
    xcl.DisplayAlerts=False
    wb.SaveAs(filename,None,'','')
    xcl.Quit()

def main():
    file_name='BEC 00760_ EXAMPLE EXTRACT FIELDS.xlsm'
    #unprotect_xlsm_file(path, file_name)
    temp_file = BEC00760(file_name)
    temp_file.extract_data()
    if (temp_file.check_available_result()):
        temp_file.write_csv_file()
    else:
        print 'Output data is not available'
    print 'Done!'

if __name__=='__main__':
    main()
