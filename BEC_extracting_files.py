import pandas as pd
import numpy as np
import re
import os
import sys
import win32com.client
import time
from tqdm import tqdm

path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

class BEC_Non_Domestic(object):
    def __init__(self,bec_file,sheetName,project_name,file_name):
        self.fileName = file_name
        self.project_name=project_name
        self.sheetName= sheetName
        self.sheet = pd.read_excel(bec_file,sheetName,keep_default_na =False,header=None).dropna(thresh=1)
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
        TEMP_data_site_measures_proposed_energy_upgrades =self.sheet.iloc[15:25,0:7].reset_index(drop=True).drop([1,3,5],axis=1)
        TEMP_data_site_measures_unit = self.sheet.iloc[14:25, 7:24].drop(15,axis=0).reset_index(drop=True).drop([20,21],axis=1)
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

class BEC_project(object):
    def __init__(self,folder,file):
        self.file_name = file
        self.input_folder= path+folder+'/'
        self.out_put_folder = ''
        self.project_name = re.search(r'BEC(\s?)\d+',file).group()
        self.project_year = re.search(r'\d+',self.input_folder).group()
        self.bec_file = pd.ExcelFile(self.input_folder+file)
        self.BEC_worksheet={}
        self.empty_line = []
        self.beneficiary_dataframe = None
        self.project_summary_dataframe = None
        self.site_references = None
        self.site_measures = None
        for sheetName in self.bec_file.sheet_names:
            if ('Project Summary' == sheetName):
                self.BEC_worksheet[sheetName] = pd.read_excel(self.bec_file, sheetName, keep_default_na=False,header=None)
            if ('Non Domestic' in sheetName):
                self.BEC_worksheet[sheetName] = BEC_Non_Domestic(self.bec_file,sheetName,self.project_name,self.file_name)
            if ('Beneficiary' == sheetName):
                self.BEC_worksheet[sheetName] = pd.read_excel(self.bec_file, sheetName, keep_default_na=False,header=None)


    def print_list_sheet(self):
        print self.bec_file.sheet_names

    def print_original_sheet(self):
        for sheetName in self.bec_file.sheet_names:
            if ('Non Domestic' in sheetName):
                self.BEC_worksheet[sheetName].print_input_sheet_content()
            elif ('Project Summary' == sheetName or 'Beneficiary' == sheetName):
                print 'File name: ', self.project_name
                print 'Sheet name: ', sheetName
                print self.BEC_worksheet[sheetName]

    def extract_summary_data(self):
        TEMP_dataframe = self.BEC_worksheet['Project Summary'].iloc[:,1]
        list_Values_Automatically_brought = TEMP_dataframe[TEMP_dataframe=='Values automatically brought in from "Non Domestic " sheets'].index.tolist()
        list_Add_addition_row = TEMP_dataframe[TEMP_dataframe=='Add additional rows as required'].index.tolist()
        if (len(list_Add_addition_row)==1):
            TEMP_data_project_summary1 = self.BEC_worksheet['Project Summary'].iloc[list_Values_Automatically_brought[-1]+1:list_Add_addition_row[0], 0:6].reset_index(drop=True).drop(3,axis=1)
            list_0 = TEMP_data_project_summary1[TEMP_data_project_summary1.iloc[:,1] == 0].index.tolist()
            list_empty = TEMP_data_project_summary1[TEMP_data_project_summary1[2]==''].index.tolist()
            self.empty_line=list_0+list_empty
            TEMP_data_project_summary1 = TEMP_data_project_summary1.drop(self.empty_line,axis=0).reset_index(drop=True)
            if (len(TEMP_data_project_summary1.iloc[:,3].unique())==1 and TEMP_data_project_summary1.iloc[:,3].unique()[0]==u' '):
                TEMP_data_project_summary1.drop(4,axis=1,inplace=True)
            else:
                TEMP_data_project_summary1.iloc[0,3]+=' (%)'
            TEMP_data_project_summary1.iloc[0, 3] += ' (%)'
            TEMP_data_project_summary1.update((TEMP_data_project_summary1.iloc[1:, 3:] * 100).astype(int))
            TEMP_data_project_summary2 = self.BEC_worksheet['Project Summary'].iloc[list_Values_Automatically_brought[-1]-1:list_Add_addition_row[0],18:21].drop([list_Values_Automatically_brought[-1],list_Values_Automatically_brought[-1]+1],axis=0).reset_index(drop=True)
            TEMP_data_project_summary2 = TEMP_data_project_summary2.drop(self.empty_line,axis=0).reset_index(drop=True)
            data_project_summary = pd.concat([TEMP_data_project_summary1, TEMP_data_project_summary2], axis=1)
            data_project_summary.insert(0,'-1',self.project_name)
            data_project_summary.iloc[0,0]='Project Code'
            data_project_summary.iloc[0,1]='Tab'
            data_project_summary.insert(0, '-2', self.project_year)
            data_project_summary.iloc[0, 0] = 'Year'
            self.project_summary_dataframe=data_project_summary
        else:
            print 'Can not identify as there are more "Add additional rows as required" or no results'

    def extract_beneficiary_data(self):
        TEMP_data_beneficiary = self.BEC_worksheet['Beneficiary'].iloc[8:,1]
        data_beneficiary = TEMP_data_beneficiary.loc[~TEMP_data_beneficiary.isin(['Total Project Cost','','Enter Name of Beneficiary',0])].to_frame().reset_index(drop=True)
        data_beneficiary.insert(0,0,self.project_name)
        data_beneficiary.iloc[0,0]='Project Code'
        data_beneficiary.insert(0, '-1', self.project_year)
        data_beneficiary.iloc[0, 0] = 'Year'
        self.beneficiary_dataframe = data_beneficiary

    def extract_non_domestic_data(self):
        non_domestic_list = [i for i in self.BEC_worksheet.keys() if 'Non Domestic' in i and int(re.search(r'\b\d+\b',i).group()) in self.project_summary_dataframe[0].tolist()]
        list_measures = []
        list_reference = []
        for non_domestic_sheet in non_domestic_list:
        # Non Domestic Measures
            non_domestic_measures = self.BEC_worksheet[non_domestic_sheet].extract_data_from_input_sheet()[0]
            non_domestic_measures.insert(0, '2', [i for i in range(non_domestic_measures.shape[0])])
            if len(list_measures)>0:
                non_domestic_measures=non_domestic_measures.drop(0,axis=0)
            non_domestic_measures.insert(0, '1', non_domestic_sheet)
            list_measures.append(non_domestic_measures)
        # Non Domestic Reference
            non_domestic_reference= self.BEC_worksheet[non_domestic_sheet].extract_data_from_input_sheet()[1].transpose()
            non_domestic_reference.insert(0, '2', int(re.search(r'\b\d+\b',non_domestic_sheet).group()))
            if len(list_reference)>0:
                non_domestic_reference=non_domestic_reference.drop(0,axis=0)
            non_domestic_reference.insert(0, '1', non_domestic_sheet)
            list_reference.append(non_domestic_reference)
    #Non Domestic Measures
        self.site_measures = pd.concat(list_measures,ignore_index=True)
        self.site_measures.insert(0, '0', self.project_name)
        self.site_measures.insert(0, '-1', self.project_year)
        self.site_measures.iloc[0,0]='Year'
        self.site_measures.iloc[0,1]='Project Code'
        self.site_measures.iloc[0,2]='Tab'
        self.site_measures.iloc[0,3]='ID Measures'
    #Non Domestic Reference
        TEMP_site_reference_df = pd.concat(list_reference,ignore_index=True)
        TEMP_site_reference_df.insert(0, '0', self.project_name)
        TEMP_site_reference_df.insert(0, '-1', self.project_year)
        TEMP_site_reference_df.iloc[0,0]='Year'
        TEMP_site_reference_df.iloc[0,1]='Project Code'
        TEMP_site_reference_df.iloc[0,2]='Tab'
        TEMP_site_reference_df.iloc[0,3]='ID References'
        TEMP_site_reference_df.iloc[0,11]+=' (number)'
        TEMP_site_reference_df.insert(12, 'Unit', 'Unit')
        TEMP_site_reference_df.iloc[1:,12]=TEMP_site_reference_df.iloc[1:,11].str.replace(r'\d+(\.?)\d+','',regex=True)
        TEMP_site_reference_df.iloc[1:, 11] = TEMP_site_reference_df.iloc[1:, 11].str.extract(r'(\d+(\.?)\d+)',expand=False)[0]
        self.site_references=TEMP_site_reference_df

    def extract_data(self):
        self.extract_summary_data()
        if 'Beneficiary' in self.bec_file.sheet_names:
            self.extract_beneficiary_data()
        self.extract_non_domestic_data()

    def check_available_result(self):
        #if (self.project_summary_dataframe is not None and self.beneficiary_dataframe is not None  and self.site_references is not None  and self.site_measures.shape[0] is not None ):
        if (self.project_summary_dataframe is not None and self.site_references is not None and self.site_measures.shape[0] is not None):
            return True
        else:
            return False

    def print_output_sheets(self):
        if self.check_available_result():
            print 'Project summary',self.project_summary_dataframe
            print 'Beneficiary', self.beneficiary_dataframe
            print 'Site references', self.site_references
            print 'Site measures', self.site_measures
        else:
            print 'Need to run extract_data() to execute input file to have results'

    def write_csv_file(self,folder_name):
        if not os.path.exists(path+folder_name+' Extracted Data/'):
            os.makedirs(path+folder_name+' Extracted Data/')
        new_path = path + folder_name + ' Extracted Data/'
        if not os.path.exists(new_path+self.project_name+'/'):
            os.makedirs(new_path+self.project_name+'/')
        new_path +=self.project_name+'/'
        self.out_put_folder = new_path
        if (self.project_summary_dataframe is not None):
            self.project_summary_dataframe.to_excel(self.out_put_folder+self.project_name+'_Project Summary.xlsx','Project Summary',header=False,index=False)
        if (self.beneficiary_dataframe is not None):
            self.beneficiary_dataframe.to_excel(self.out_put_folder+self.project_name+'_Beneficiary.xlsx','Beneficiary',header=False,index=False)
        if (self.site_references is not None and self.site_measures is not None):
            self.site_references.to_excel(self.out_put_folder+self.project_name+'_References.xlsx','References',header=False,index=False)
            self.site_measures.to_excel(self.out_put_folder+self.project_name+'_Measures.xlsx','Measures',header=False,index=False)

def unprotect_xlsm_file(path,filename):
    xcl = win32com.client.Dispatch('Excel.Application')
    pw_str = 'Bec2018dec2017'
    wb = xcl.Workbooks.Open(path+filename,False,True,None,pw_str)
    xcl.DisplayAlerts=False
    wb.SaveAs(filename,None,'','')
    xcl.Quit()

def access_to_working_file(folder_name):
    files = os.listdir(path+folder_name)
    return files

def execute_each_project(folder_name):
    file_list =access_to_working_file(folder_name)
    errors = []
    if (len(file_list) > 0):
        for file_name in tqdm(file_list):
            if ('772' in file_name):
                #try:
                    temp_file = BEC_project(folder_name,file_name)
                    temp_file.extract_data()
                    if (temp_file.check_available_result()):
                        temp_file.write_csv_file(folder_name)
                #except Exception:
                #   errors.append(temp_file.project_name + ' from ' + temp_file.file_name )
    else:
        print 'Folder '+folder_name+' is empty'
    #if (len(errors)>0):
    #    print ''
    #    print 'Errors: ',errors

def main():
    start_time = time.time()
    folder_name = 'BEC 2018'
    execute_each_project(folder_name)
    print 'Done! from ', time.asctime( time.localtime(start_time)),' to ',time.asctime( time.localtime(time.time()))

if __name__=='__main__':
    main()


