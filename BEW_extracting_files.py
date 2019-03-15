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

def write_file(path,folder_name,df,new_file_name):
    # Create a shared folder along side with year
    new_path = path +'Shared Data/'
    if not os.path.exists(new_path):
        os.makedirs(new_path)
        if not (os.path.isfile(new_path + new_file_name+'.xlsx')):
            df.to_excel(new_path + new_file_name+'.xlsx',new_file_name,header=False, index=False)
    else:
        if not (os.path.isfile(new_path + new_file_name+'.xlsx')):
            df.to_excel(new_path + new_file_name+'.xlsx',new_file_name,header=False, index=False)
        else:
            book = load_workbook(new_path +new_file_name+'.xlsx')
            writer = pd.ExcelWriter(new_path +new_file_name+'.xlsx',engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df.iloc[1:, :].to_excel(writer, new_file_name, index=False, header=False,startrow=writer.sheets[new_file_name].max_row)
            writer.save()


def write_to_1_file(path,df):
    empty_list = df[df[1] == ''].index.tolist()
    if (len(empty_list) > 0):
        df = (df.drop(empty_list, axis=0).reset_index(drop=True))
    write_file(path,'',df,'Evaluation')
    # new_path = path + 'Shared Data Evaluation/'
    # if not os.path.exists(path + 'Shared Data Evaluation/'):
    #     os.makedirs(path + 'Shared Data Evaluation/')
    #     if not (os.path.isfile(new_path + 'Evaluation.xlsx')):
    #         df.to_excel(new_path +'Evaluation.xlsx', 'Summary Sheet',header=False, index=False)
    # else:
    #     book = load_workbook(new_path +'Evaluation.xlsx')
    #     writer = pd.ExcelWriter(new_path +'Evaluation.xlsx', engine='openpyxl')
    #     writer.book = book
    #     writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    #     df.iloc[1:, :].to_excel(writer, 'Summary Sheet', index=False, header=False,startrow=writer.sheets['Summary Sheet'].max_row)
    #     writer.save()

def extract_data(excel_file,tab,extracted_lst,skiprow):
    temp_df = pd.read_excel(excel_file, tab, skiprows=skiprow,keep_default_na=False, header=None)
    print('Phase 2')
    df_1_line = temp_df.iloc[0]
    print('Phase 3')
    if tab == 'Technologies':
        print('Phase 4')
        extracted_df = temp_df.iloc[:,2:]
    else:
        col_index_workbook = find_column(df_1_line, extracted_lst)
        col_extended_workbook = find_extended_column(tab, df_1_line, col_index_workbook)
        print('Phase 4')
        extracted_df = temp_df[col_extended_workbook]
    return extracted_df

def find_extended_column(tab,df_l_line,num_list):
    num_list = sorted(num_list)
    new_num_lst = []
    if tab == 'BE Workplaces main workbook':
        start_point = df_l_line[df_l_line=='Select Thermal Fuel'].index.tolist()[0]
        end_point = df_l_line[df_l_line == 'Total Energy Cost Savings €'].index.tolist()[0]
        start_point2 = df_l_line[df_l_line=='Primary Energy Savings kWh'].index.tolist()[0]
        end_point2 = df_l_line[df_l_line == 'Site Energy Reduction %'].index.tolist()[0]
        num_list[num_list.index(end_point2)] = end_point2-1
        end_point2 = end_point2-1
        new_num_lst = num_list[:num_list.index(start_point)]+list(range(start_point,end_point+1))+num_list[num_list.index(end_point)+1:num_list.index(start_point2)+1]+list(range(start_point2,end_point2))+num_list[num_list.index(end_point2):]
    if len(new_num_lst)==0:
        new_num_lst = num_list
    return new_num_lst

def find_column(df_1_line,lst_to_find):
    return df_1_line[df_1_line.astype(str).isin(lst_to_find)].index.tolist()

def assign_task_Evaluation(seeep_path,folder):
    input_folder = seeep_path + folder + '/'
    file_path_lst = os.listdir(input_folder)
    print (file_path_lst)
    for file in file_path_lst:
        if re.search(r'Batch',file):
            excel_file = pd.ExcelFile(input_folder+file)
            col_lst = ['Reference', 'Applicant', 'Description']
            summary_df = extract_data(excel_file,'Summary Sheet',col_lst,0)
            write_to_1_file(seeep_path[:-9],summary_df)


def assign_task_Summary(seeep_path,file,folder):
    input_folder = seeep_path + folder + '/'
    project_name = re.search(r'\w+\s+\w+', file).group()
    project_year = re.search(r'\d+', folder).group()
    excel_file = pd.ExcelFile(input_folder + file)
    if 'Admin' in excel_file.sheet_names:
        lst_col_admin = ['Reference No.','Cat. ','Submitted By','Project Title','County','Approved Funding']
        admin_df = extract_data(excel_file,'Admin',lst_col_admin,1)
        write_file(seeep_path,folder,admin_df,'Admin')

def assign_task_Overview(seeep_path,file,folder):
    input_folder = seeep_path + folder + '/'
    project_name = re.search(r'\w+\s+\w+', file).group()
    project_year = re.search(r'\d+', folder).group()
    print ('Phase 0')
    excel_file = pd.ExcelFile(input_folder + file)
    print ('Phase 1')
    # if 'BE Workplaces main workbook' in excel_file.sheet_names:
    #     lst_col_workbook = ['SEAI Reference', 'Organisation', 'Project Title', 'Total Incl VAT', 'Total Excl VAT',
    #                         'Select Thermal Fuel', 'Total Energy Cost Savings €', 'Grant  /Approved (Proposed)',
    #                         'Grant /Approved (Proposed)', 'Primary Energy Savings kWh', 'Site Energy Reduction %']
    #     df_workplaces = extract_data(excel_file,'BE Workplaces main workbook',lst_col_workbook,3)
    #     write_file(seeep_path,folder, df_workplaces,'Workplaces')
    if 'Technologies' in excel_file.sheet_names:
        lst_col_tech = []
        tech_df = extract_data(excel_file,'Technologies',lst_col_tech,0)
        write_file(seeep_path,folder, tech_df,'Technologies')
    print ('Phase 5')

def execute_each_folder(seeep_path,folder_name):
    file_path = seeep_path+folder_name+'/'
    file_path_lst = os.listdir(file_path)
    for file in file_path_lst:
        if file == 'Evaluations':
            assign_task_Evaluation(seeep_path+'BEW 2012/',file)
        if re.search(r'Better Energy',file) and re.search(r'Summary',file):
            assign_task_Summary(seeep_path,file,folder_name)
        if re.search(r'Better Energy Board',file):
            assign_task_Overview(seeep_path,file,folder_name)


def main():
    start_time = time.time()
    path_lst = os.listdir(path)
    if 'SEEEP' in path_lst:
        seeep_path = path+'SEEEP/'
        folder = os.listdir(seeep_path)
        for folder_name in folder:
            if re.search(r'BEW',folder_name):
                execute_each_folder(seeep_path,folder_name)
    print('Done! from ', time.asctime(time.localtime(start_time)), ' to ',time.asctime(time.localtime(time.time())))

if __name__ == '__main__':
    main()
