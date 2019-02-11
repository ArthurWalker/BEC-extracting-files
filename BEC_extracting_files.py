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
import openpyxl
from openpyxl import load_workbook
from fuzzywuzzy import fuzz
import msoffcrypto
import xlrd

path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

class BEC_Non_Domestic(object):
    def __init__(self,bec_file,sheetName,project_name,file_name,tab):
        self.fileName = file_name
        self.project_name=project_name
        self.sheetName= sheetName
        self.tab=tab
        self.sheet = pd.read_excel(bec_file,sheetName,keep_default_na =False,header=None).dropna(thresh=1)
        self.data_site_reference = ''
        self.data_site_measures = ''

    # Extract input data from non domestic tab
    def extract_data_from_input_sheet(self):
        #print (self.tab)
        try:
            proposed_engergy_upgrade_index = self.sheet.iloc[:,0][self.sheet.iloc[:,0]=='Proposed Energy Upgrades'].index.tolist()[0]
        except IndexError:
            proposed_engergy_upgrade_index=14
        project_category_index = self.sheet.iloc[:,0][self.sheet.iloc[:,0]=='Project Category'].index.tolist()[0]
        small_df = pd.concat([self.sheet.loc[proposed_engergy_upgrade_index-3:proposed_engergy_upgrade_index-1,3],self.sheet.loc[proposed_engergy_upgrade_index-3:proposed_engergy_upgrade_index-1,2]],axis=1,sort=False)
        small_df = small_df.transpose().reset_index(drop=True).transpose()
        TEMP_df = self.sheet.iloc[project_category_index:proposed_engergy_upgrade_index,0:2].reset_index(drop=True)
        list_empty = TEMP_df[TEMP_df[0] == ''].index.tolist()
        if (len(list_empty)>0):
            TEMP_df=(TEMP_df.drop(list_empty, axis=0).reset_index(drop=True))
        self.data_site_reference = TEMP_df.append(small_df,ignore_index=True,sort=False)
        extracted_column_index_for_site_measures_energy_upgrades = self.sheet.iloc[proposed_engergy_upgrade_index,0:][self.sheet.iloc[proposed_engergy_upgrade_index,0:]=='Electrical Savings kWh'].index.tolist()[0]
        columns_to_drop_energy = self.sheet.iloc[proposed_engergy_upgrade_index+1,0:][self.sheet.iloc[proposed_engergy_upgrade_index+1,0:].isin(['Description of Minimum Data Required for Existing Specification','Description of Minimum Data Required for Proposed Specification','Additional Information'])].index.tolist()
        TEMP_data_site_measures_proposed_energy_upgrades =self.sheet.iloc[proposed_engergy_upgrade_index+1:25,0:extracted_column_index_for_site_measures_energy_upgrades].reset_index(drop=True).drop(columns_to_drop_energy,axis=1)
        last_column_unit = self.sheet.iloc[proposed_engergy_upgrade_index,0:][self.sheet.iloc[proposed_engergy_upgrade_index,0:]=='Energy Credits'].index.tolist()[-1]
        columns_to_drop_unit = self.sheet.iloc[proposed_engergy_upgrade_index, extracted_column_index_for_site_measures_energy_upgrades:last_column_unit+1][self.sheet.iloc[proposed_engergy_upgrade_index, extracted_column_index_for_site_measures_energy_upgrades:last_column_unit+1].isin(['Milestone', 'Invoice', '','Milestone Claim','Amount'])].index.tolist()
        TEMP_data_site_measures_unit = self.sheet.iloc[proposed_engergy_upgrade_index:25, extracted_column_index_for_site_measures_energy_upgrades:last_column_unit+1].drop(proposed_engergy_upgrade_index+1,axis=0).reset_index(drop=True).drop(columns_to_drop_unit,axis=1)
        TEMP_data_site_measures = pd.concat([TEMP_data_site_measures_proposed_energy_upgrades,TEMP_data_site_measures_unit],axis=1,sort=False)
        TEMP_data_site_measures.columns = [i for i in range(TEMP_data_site_measures.shape[1])]
        self.data_site_measures = TEMP_data_site_measures.loc[~TEMP_data_site_measures[0].astype(str).isin(['Total','','-',' '])]
        return [self.data_site_measures,self.data_site_reference]

    def print_input_sheet_content(self):
        print ('File name: ',self.fileName)
        print ('Sheet name: ',self.sheetName)
        print (self.sheet)

    def print_output_sheet_content(self):
        print ('')
        print ('Data of site reference: ')
        print (self.data_site_reference)
        print ('')
        print ('Data of site measures: ')
        print (self.data_site_measures)

class BEC_project(object):
    def __init__(self,folder,file):
        self.file_name = file
        self.input_folder= path+folder+'/'
        self.out_put_folder = ''
        self.project_name = re.search(r'BEC(\s|\_?)\d+(\s?)\d+',file).group()
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
                self.BEC_worksheet[sheetName] = BEC_Non_Domestic(self.bec_file,sheetName,self.project_name,self.file_name,sheetName)
            if ('Beneficiary' == sheetName):
                self.BEC_worksheet[sheetName] = pd.read_excel(self.bec_file, sheetName, keep_default_na=False,header=None)

    #Extract project summary data
    def extract_summary_data(self):
        TEMP_dataframe = self.BEC_worksheet['Project Summary'].iloc[:,1].astype(str)
        TEMP_dataframe2 = self.BEC_worksheet['Project Summary'].iloc[:,0].astype(str)
        # If you want to remove unnecessary data
        # TEMP_dataframe2.replace('', np.nan, inplace=True)
        # TEMP_dataframe2.replace(r'^\d+$', np.nan, inplace=True,regex=True)
        # TEMP_dataframe2.dropna(inplace=True)
        list_Values_Better_Energy_Communities_Programes_Non_Domestic_costs =TEMP_dataframe2[TEMP_dataframe2.str.contains('Better Energy Communities Programme - Non Domestic Costs',na=False)].index.tolist()
        list_Add_addition_row = TEMP_dataframe[TEMP_dataframe=='Add additional rows as required'].index.tolist()
        if len(list_Add_addition_row)==0:
            list_Add_addition_row = TEMP_dataframe2[TEMP_dataframe2.str.contains('Better Energy Communities Programme - Domestic Costs',na=False)].index.tolist()
        if (len(list_Add_addition_row)==1):
            if int(self.project_year) >=2017:
                column_to_collect = 6
            else:
                column_to_collect = 4
            # Get data of the first half of requested table
            TEMP_data_project_summary1 = self.BEC_worksheet['Project Summary'].iloc[list_Values_Better_Energy_Communities_Programes_Non_Domestic_costs[-1]+2:list_Add_addition_row[0], 0:column_to_collect].reset_index(drop=True).drop(3,axis=1)
            list_0 = TEMP_data_project_summary1[TEMP_data_project_summary1.iloc[:,1] == 0].index.tolist()
            list_empty = TEMP_data_project_summary1[TEMP_data_project_summary1[2]==''].index.tolist()
            self.empty_line=list_0+list_empty
            TEMP_data_project_summary1 = TEMP_data_project_summary1.drop(self.empty_line,axis=0).reset_index(drop=True)
            if int(self.project_year) >=2017:
                # Convert % num to numemric
                if (len(TEMP_data_project_summary1.iloc[:,3].unique())==1 and TEMP_data_project_summary1.iloc[:,3].unique()[0]==u' '):
                    TEMP_data_project_summary1.drop(4,axis=1,inplace=True)
                TEMP_data_project_summary1.iloc[1:,3:]=TEMP_data_project_summary1.iloc[1:,3:].fillna(0)
                TEMP_data_project_summary1.update((TEMP_data_project_summary1.iloc[1:, 3:] * 100).astype(int))
            # Get data of the second half of requested table
            header_line_boolean = self.BEC_worksheet['Project Summary'].iloc[list_Values_Better_Energy_Communities_Programes_Non_Domestic_costs[-1]].astype(str).isin(['Total Project Cost','SEAI funding','Eligible VAT','SEAI Funding'])
            header_line_index = header_line_boolean[header_line_boolean==True].index.tolist()
            if int(self.project_year)==2016:
                header_line_index = header_line_index[:3]
            TEMP_data_project_summary2 = self.BEC_worksheet['Project Summary'].iloc[list_Values_Better_Energy_Communities_Programes_Non_Domestic_costs[-1]:list_Add_addition_row[0],header_line_index].drop([list_Values_Better_Energy_Communities_Programes_Non_Domestic_costs[-1]+1,list_Values_Better_Energy_Communities_Programes_Non_Domestic_costs[-1]+2],axis=0).reset_index(drop=True)
            TEMP_data_project_summary2 = TEMP_data_project_summary2.drop(self.empty_line,axis=0).reset_index(drop=True)
            # Merge 2 tables into 1
            data_project_summary = pd.concat([TEMP_data_project_summary1, TEMP_data_project_summary2], axis=1,sort=False)
            data_project_summary.insert(0,'-1',self.project_name)
            data_project_summary.iloc[0,0]='Project Code'
            data_project_summary.iloc[0,1]='Tab'
            data_project_summary.insert(0, '-2', self.project_year)
            data_project_summary.iloc[0, 0] = 'Year'
            self.project_summary_dataframe=data_project_summary
        else:
            print ('Can not identify as there are more "Add additional rows as required" or no results')

    # Extract beneficiary data
    def extract_beneficiary_data(self):
        TEMP_data_beneficiary = self.BEC_worksheet['Beneficiary'].iloc[8:,1]
        data_beneficiary = TEMP_data_beneficiary.loc[~TEMP_data_beneficiary.isin(['Total Project Cost','','Enter Name of Beneficiary','Enter Name of Beneficiary ','Name Of Beneficiary',0])].to_frame().reset_index(drop=True)
        data_beneficiary.insert(0,0,self.project_name)
        data_beneficiary.iloc[0,0]='Project Code'
        data_beneficiary.insert(0, '-1', self.project_year)
        data_beneficiary.iloc[0, 0] = 'Year'
        self.beneficiary_dataframe = data_beneficiary

    # Collect data from each non domestic data tab in each project
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
        TEMP_site_measures_df = pd.concat(list_measures,ignore_index=True,sort=False)
        TEMP_site_measures_df.insert(0, '0', self.project_name)
        TEMP_site_measures_df.insert(0, '-1', self.project_year)
        TEMP_site_measures_df.iloc[0,0]='Year'
        TEMP_site_measures_df.iloc[0,1]='Project Code'
        TEMP_site_measures_df.iloc[0,2]='Tab'
        TEMP_site_measures_df.iloc[0,3]='ID Measures'
        #TEMP_site_measures_df.iloc[0,16:20]=(deal_with_strange_characters(TEMP_site_measures_df.iloc[0, 16:20]))
        self.site_measures = TEMP_site_measures_df
    #Non Domestic Reference
        TEMP_site_reference_df = pd.concat(list_reference,ignore_index=True,sort=False)
        TEMP_site_reference_df.insert(0, '0', self.project_name)
        TEMP_site_reference_df.insert(0, '-1', self.project_year)
        TEMP_site_reference_df.iloc[0,0]='Year'
        TEMP_site_reference_df.iloc[0,1]='Project Code'
        TEMP_site_reference_df.iloc[0,2]='Tab'
        TEMP_site_reference_df.iloc[0,3]='ID References'
        columns = TEMP_site_reference_df.iloc[0,:].reset_index(drop=True)
        floor_area = columns[columns=='Floor Area of building'].index[0]
        TEMP_site_reference_df.insert(int(floor_area+1), 'Unit', 'Unit')
        TEMP_site_reference_df.insert(int(floor_area+1), 'Number', 'Num')
        TEMP_site_reference_df.loc[1:,'Unit']=TEMP_site_reference_df.iloc[1:,int(floor_area)].astype(str).str.replace(r'\d+(\.?)\d+','',regex=True)
        TEMP_site_reference_df.loc[1:, 'Number'] = TEMP_site_reference_df.iloc[1:, int(floor_area)].astype(str).str.extract(r'(\d+(\.?)\d+)',expand=False)[0]
        #TEMP_site_reference_df.iloc[0,16:20]=(deal_with_strange_characters(TEMP_site_reference_df.iloc[0, 16:20]))
        self.site_references=TEMP_site_reference_df

    # Function that controls extracting functions
    def extract_data(self):
        self.extract_summary_data()
        if 'Beneficiary' in self.bec_file.sheet_names:
            self.extract_beneficiary_data()
        self.extract_non_domestic_data()

    # Checking if attributes are available or not
    def check_available_result(self):
        if (self.project_summary_dataframe is not None and self.site_references is not None and self.site_measures.shape[0] is not None):
            return True
        else:
            return False

    # Write individual project into seperate files
    def write_seperate_excel_file(self,folder_name):
        if not os.path.exists(path+folder_name+' Shared Data/'):
            os.makedirs(path+folder_name+' Shared Data/')
        new_path = path + folder_name + ' Shared Data/'
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
            self.site_references.to_excel(self.out_put_folder+self.project_name+'_References.xlsx','References',header=False,index=False)
            self.site_measures.to_excel(self.out_put_folder+self.project_name+'_Measures.xlsx','Measures',header=False,index=False)

    # Write to files
    def write_files(self,dataframe,file_name):
        if not (os.path.isfile(self.out_put_folder+file_name+'.xlsx')):
            dataframe.to_excel(self.out_put_folder +file_name+'.xlsx',file_name, header=False, index=False)
        else:
        # Tab
            dataframe.columns=[i for i in range(len(dataframe.columns))]
            current_df = dataframe
            extracted_df = pd.read_excel(self.out_put_folder + file_name + '.xlsx', file_name, keep_default_na=False,header=None, index=False,nrows=1)
        # to see what makes different
            if current_df.iloc[0, :].astype(str).tolist() != extracted_df.iloc[0,:].astype(str).tolist():
                # Checking for missing headers in both dataframe
                if check_different(current_df.iloc[0, :].astype(str).tolist(),extracted_df.iloc[0,:].astype(str).tolist()):
                    extracted_df_index_missing = find_difference(current_df.iloc[0, :].astype(str).tolist(),extracted_df.iloc[0,:].astype(str).tolist(),'missing')
                    if extracted_df_index_missing is not None and len(extracted_df_index_missing)>0:
                        extracted_df = pd.read_excel(self.out_put_folder + file_name + '.xlsx', file_name,keep_default_na=False, header=None, index=False)
                        for new_column in extracted_df_index_missing:
                            #new_column[1] += extracted_df_index_missing.index(new_column)
                            extracted_df = fill_empty_value_into_blank_columns(new_column, extracted_df)
                        extracted_df.to_excel(self.out_put_folder + file_name + '.xlsx', file_name, header=False,index=False)
                if check_different(extracted_df.iloc[0,:].astype(str).tolist(),current_df.iloc[0, :].astype(str).tolist()):
                    current_df_index_missing = find_difference(extracted_df.iloc[0, :].astype(str).tolist(),current_df.iloc[0, :].astype(str).tolist(), 'missing')
                    if current_df_index_missing is not None and len(current_df_index_missing) > 0:
                        for new_column in current_df_index_missing:
                            #new_column[1] += current_df_index_missing.index(new_column)
                            current_df = fill_empty_value_into_blank_columns(new_column, current_df)
                # Checking for different headers in both dataframe
                if check_different(current_df.iloc[0, :].astype(str).tolist(),extracted_df.iloc[0, :].astype(str).tolist()):
                    extracted_df_index_different = find_difference(current_df.iloc[0, :].astype(str).tolist(),extracted_df.iloc[0, :].astype(str).tolist(),'different')
                    if extracted_df_index_different is not None and len(extracted_df_index_different)>0:
                        if (int(self.project_year)>2018):
                            for column in extracted_df_index_different:
                                extracted_df.iloc[0,column[1]]=column[0]
                        else:
                            for column in extracted_df_index_different:
                                current_df.iloc[0,column[1]]=extracted_df.iloc[0,column[1]]
                if check_different(extracted_df.iloc[0, :].astype(str).tolist(),current_df.iloc[0, :].astype(str).tolist()):
                    current_df_index_different = find_difference(extracted_df.iloc[0,:].astype(str).tolist(),current_df.iloc[0, :].astype(str).tolist(),'different')
                    if current_df_index_different is not None and len(current_df_index_different) >0:
                        if (int(self.project_year)>2018):
                            for column in current_df_index_different:
                                extracted_df.iloc[0,column[1]]=column[0]
                        else:
                            for column in current_df_index_different:
                                current_df.iloc[0,column[1]]=extracted_df.iloc[0,column[1]]
            if  current_df.iloc[0, :].astype(str).tolist() == extracted_df.iloc[0,:].astype(str).tolist():
                book = load_workbook(self.out_put_folder + file_name + '.xlsx')
                writer = pd.ExcelWriter(self.out_put_folder + file_name + '.xlsx', engine='openpyxl')
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                current_df.iloc[1:, :].to_excel(writer, file_name, index=False, header=False,startrow=writer.sheets[file_name].max_row)
                writer.save()
            else:
                print (self.project_name,'hasnt printed',file_name)

            # Append to the whole df which is not recommended
            # extracted_df = pd.read_excel(self.out_put_folder + file_name + '.xlsx', file_name,
            #                              keep_default_na=False, header=None, index=False)
            # extracted_df = extracted_df.drop(extracted_df.index[0])
            # lastest_update_df = extracted_df.append(current_df, sort=False, ignore_index=True)
            # lastest_update_df.to_excel(self.out_put_folder + file_name + '.xlsx', file_name, header=False,index=False)

    # Add data into an excel file
    def add_project(self):
        # if not os.path.exists(path + 'BEC ' + self.project_year + ' Shared Data/'):
        #     os.makedirs(path + 'BEC ' + self.project_year + ' Shared Data/')
        # self.out_put_folder = path + 'BEC ' + self.project_year + ' Shared Data/'
        if not os.path.exists(path+'BEC Shared Data/'):
            os.makedirs(path+'BEC Shared Data/')
        self.out_put_folder= path+'BEC Shared Data/'
        self.write_files(self.project_summary_dataframe,'Project Summary')
        if (self.beneficiary_dataframe is not None):
           self.write_files(self.beneficiary_dataframe,'Beneficiary')
        self.write_files(self.site_measures,'Site Measures')
        self.write_files(self.site_references,'Site References')

    def print_list_sheet(self):
        print (self.bec_file.sheet_names)

    def print_original_sheet(self):
        for sheetName in self.bec_file.sheet_names:
            if ('Non Domestic' in sheetName):
                self.BEC_worksheet[sheetName].print_input_sheet_content()
            elif ('Project Summary' == sheetName or 'Beneficiary' == sheetName):
                print ('File name: ', self.project_name)
                print ('Sheet name: ', sheetName)
                print (self.BEC_worksheet[sheetName])

    def print_output_sheets(self):
        if self.check_available_result():
            print ('Project summary',self.project_summary_dataframe)
            print ('Beneficiary', self.beneficiary_dataframe)
            print ('Site references', self.site_references)
            print ('Site measures', self.site_measures)
        else:
            print ('Need to run extract_data() to execute input file to have results')

def check_different(list1,list2):
    if len(list(set(list1)-set(list2)))>0:
        return True
    return False

def check_header(text,list_text):
    for i in list_text:
        if fuzz.ratio(i,text)>=92:
            return ['different',list_text.index(i)]
    return ['missing']

def fill_empty_value_into_blank_columns(new_column,current_df):
    current_df.insert(new_column[1], 'Empty Value', new_column[0])
    current_df.columns = [i for i in range(len(current_df.columns))]
    current_df.iloc[1:, new_column[1]] = ''
    return current_df

def find_difference(list1,list2,flag):
    index_list_different, index_list_missing=None,None
    if len(list(set(list1)-set(list2))):
        diff = list(set(list1)-set(list2))
        index_list_different =[[i,list1.index(i),check_header(i,list2)[1]] for i in diff if check_header(i,list2)[0]=='different']
        index_list_missing =[[i,list1.index(i)] for i in diff if check_header(i,list2)[0]=='missing']
        index_list_different.sort(key=lambda x:x[1])
        index_list_missing.sort(key=lambda x:x[1])
    if (flag=='missing'):
        return index_list_missing
    if (flag=='different'):
        return index_list_different
    return

def find_details_of_deferences(list1,list2):
    diff = list(set(list1).symmetric_difference(set(list2)))
    index_list1 = []
    index_list2 = []
    for i in diff:
        if i in list1:
            index_list1.append([i,list1.index(i)])
        elif i in list2:
            index_list2.append([i,list2.index(i)])
    index_list1.sort(key=lambda x: x[1])
    index_list2.sort(key=lambda x: x[1])
    return index_list1,index_list2

def deal_with_strange_characters(series):
    series=series.apply(lambda x:x.encode('utf-8').decode('utf-8'))
    return series

def unprotect_xlsm_file(path,filename,passw):
    xcl = win32com.client.Dispatch('Excel.Application')
    #Pass for files in 2018 'Bec2018dec2017'
    pw_str = passw
    wb = xcl.Workbooks.Open(path+filename,False,True,None,pw_str)
    xcl.DisplayAlerts=False
    wb.SaveAs(path+filename+'x',None,'','')
    xcl.Quit()

def access_to_working_file(folder_name):
    files = os.listdir(path+folder_name)
    return files

def execute_each_project_in_a_year(folder_name):
    file_list =access_to_working_file(folder_name)
    errors = []
    if (len(file_list) > 0):
        for file_name in tqdm(file_list):
            if ('.xlsm' in file_name or '.xlsx' in file_name or '.xls' in file_name):
                #try:
                    temp_file = BEC_project(folder_name,file_name)
                    temp_file.extract_data()
                    if (temp_file.check_available_result()):
                        #temp_file.write_seperate_excel_file(folder_name)
                        temp_file.add_project()
                #except Exception:
                #   errors.append(temp_file.project_name + ' from ' + temp_file.file_name)
    else:
        print ('Folder '+folder_name+' is empty')
    if (len(errors)>0):
       print ('')
       print ('Errors: ',len(errors),errors)

def working_with_folder():
    folder_list = os.listdir(path)
    for folder_name in folder_list[::-1]:
        if re.search(r'^BEC \d+$',folder_name) :
            print ('Checking folder',folder_name)
            execute_each_project_in_a_year(folder_name)

def extract_randomly_data():
    selected_folder = input('Choose folder to select: ')
    selected_file = input('Choose file to select: ')
    extracted_headers = pd.read_excel(path + selected_folder+' Shared Data/' + selected_file + '.xlsx', selected_file,
                                      keep_default_na=False, header=None, index=False, nrows=1)
    extracted_data = pd.read_excel(path + selected_folder+' Shared Data/' + selected_file + '.xlsx', selected_file,
                                   keep_default_na=False, header=None, index=False)
    numb_of_rand_data = int(input('Number of data points: '))
    random_selected_data = extracted_data.iloc[1:, :].sample(numb_of_rand_data)
    result = extracted_headers.append(random_selected_data,sort=False)
    result.to_excel(selected_folder+' '+selected_file + ' ' + str(numb_of_rand_data) + ' searching results.xlsx', header=False, index=False)
    print('Done!')

def main():
    #option = input('Choose your task (1 for executing files or 2 for randomly selecting data points): ')
    option='1'
    if (option == '1'):
        start_time = time.time()
        working_with_folder()
        print('Done! from ', time.asctime(time.localtime(start_time)), ' to ',time.asctime(time.localtime(time.time())))
    else:
        extract_randomly_data()

if __name__=='__main__':
    main()


