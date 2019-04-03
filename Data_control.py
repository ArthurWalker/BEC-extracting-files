import pandas as pd
import numpy as np
import os

path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

def main():
    # Get the data
    excel_file = pd.ExcelFile(path+'BEW_EEEP Technologies.xlsx')
    df = pd.read_excel(excel_file,header=None,sheet_name='Technologies')
    reference = df.iloc[:251,:4]
    reference.columns  = [i for i in range(len(reference.columns))]
    measure = df.iloc[:251,5:]
    measure.columns = [i for i in range(len(measure.columns))]
    new_df = pd.DataFrame(['Year','SEAI Reference','Organisation','ProjectName','EnergyMeasureCategoryName','Yes/No']).transpose()
    for index,row_measure in  measure.iloc[1:,:].iterrows():
        lst = [reference.iloc[index].tolist() for i in range(40)]
        small_df_text= pd.DataFrame(lst)
        small_df_text_2= measure.iloc[0,:].to_frame()
        small_df_digit= row_measure.to_frame()
        new_small_df = pd.concat([small_df_text,small_df_text_2,small_df_digit],axis=1)
        new_small_df.columns = [i for i in range(len(new_small_df.columns))]
        new_df = new_df.append(new_small_df,ignore_index=True)
    #writer = pd.ExcelWriter(path+'BEW_EEEP Technologies.xlsx',engine= 'openpyxl')
    new_df.to_excel(path+'BEW_EEEP Technologies.xlsx','Converted Technologies',header=False, index=False)

if __name__ == '__main__':
    main()