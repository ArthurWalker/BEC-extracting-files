import pandas as pd
# import numpy as np
import os
# Initilize working folder
path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

def main():
    # Get the data
    excel_file = pd.ExcelFile(path+'BEW_EEEP Technologies.xlsx')
    df = pd.read_excel(excel_file,header=None,sheet_name='Technologies')
    # Retreive the reference data
    reference = df.iloc[:251,:4]
    # Re-range column indexies
    reference.columns  = [i for i in range(len(reference.columns))]
    # Retrieve the measure data
    measure = df.iloc[:251,5:]
    # Re-range column indexies
    measure.columns = [i for i in range(len(measure.columns))]
    # Initilize a new dataframe with headings
    new_df = pd.DataFrame(['Year','SEAI Reference','Organisation','ProjectName','EnergyMeasureCategoryName','Yes/No']).transpose()
    # For each data point in the dataframe
    for index,row_measure in  measure.iloc[1:,:].iterrows():
        # Initialize the list with default value
        lst = [reference.iloc[index].tolist() for i in range(40)]
        # Put the list of default values to a temporary dataframe
        small_df_text= pd.DataFrame(lst)
        # Get the measure data of each data point
        small_df_text_2= measure.iloc[0,:].to_frame()
        # Get the Yes/No data
        small_df_digit= row_measure.to_frame()
        # Merge all dataframes into 1 big one
        new_small_df = pd.concat([small_df_text,small_df_text_2,small_df_digit],axis=1)
        # Rearrange the column
        new_small_df.columns = [i for i in range(len(new_small_df.columns))]
        # Add the new dataframe of each old data point to the final dataframe
        new_df = new_df.append(new_small_df,ignore_index=True)
    # Write the dataframe to excel
    new_df.to_excel(path+'BEW_EEEP Technologies.xlsx','Converted Technologies',header=False, index=False)

if __name__ == '__main__':
    main()