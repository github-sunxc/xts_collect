# -*- coding: utf-8 -*-
import os
import sys
import datetime
import pandas


def save_df_to_excel(update_df, filename, sheet_name):
    '''
    "openpyxl.load_workbook" is not work on my environment, so I have to
    using a loop to substitute it, The purpose is to prevent overwriting other sheets.
    '''
    df = pandas.read_excel(filename,sheet_name=None)
    writer = pandas.ExcelWriter(filename)
    
    print('===>start to write data into %s' %(filename))
    for sheet in df.keys():
        if(sheet == sheet_name):
            update_df.to_excel(writer,index=False,sheet_name=sheet_name)
        else:
            df[sheet].to_excel(writer,index=False,sheet_name=sheet)

    writer.save()


def process_result(previous_excel, update_excel, test_type):

    previous_df = pandas.read_excel(previous_excel, sheet_name=test_type, skiprows=0)
    update_df   = pandas.read_excel(update_excel  , sheet_name=test_type, skiprows=0)
    update_copy_df = update_df
    
    for update_index in range(0, len(update_df)):
        for previous_index in range(0, len(previous_df)):
            if (previous_df.iloc[previous_index, 0] == update_df.iloc[update_index, 0] and 
                previous_df.iloc[previous_index, 1] == update_df.iloc[update_index, 1]):
                # print(update_df.iloc[update_index:update_index+1])
                update_df.iloc[update_index:update_index+1] = previous_df[previous_index:previous_index+1]
                
    save_df_to_excel(update_df, update_excel, test_type)
        

if __name__ == '__main__':
    # argvs: previous_excel; update_excel; test_type; 
    process_result(sys.argv[1], sys.argv[2], sys.argv[3])