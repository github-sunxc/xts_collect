# -*- coding: utf-8 -*-
import os
import sys
import pandas
import numpy as np

def process_result(input_excel, test_type):
    check_table_df  = pandas.read_excel(input_excel, sheet_name=test_type, skiprows=0)
    check_table_len = len(check_table_df)
    
    filename = open("cmd_8mm.txt", "a")
    filename.write("run cts ")
    
    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, 8] == 'y' and check_table_df.iloc[line, 5] is np.nan):
            filename.write("--include-filter ")
            filename.write("\"")
            filename.write(check_table_df.iloc[line, 1])
            if(check_table_df.iloc[line, 0] is np.nan):
                filename.write("\" ")
                print("error")
            else:
                filename.write(" ")
                filename.write(check_table_df.iloc[line, 0])
                filename.write("\" ")
    filename.close()
    
    filename = open("cmd_8mp.txt", "a")
    filename.write("run cts ")
    
    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, 9] == 'y' and check_table_df.iloc[line, 5] is np.nan):
            filename.write("--include-filter ")
            filename.write("\"")
            filename.write(check_table_df.iloc[line, 1])
            if(check_table_df.iloc[line, 0] is np.nan):
                filename.write("\" ")
                print("error")
            else:
                filename.write(" ")
                filename.write(check_table_df.iloc[line, 0])
                filename.write("\" ")
    filename.close()
    
    filename = open("cmd_8mq.txt", "a")
    filename.write("run cts ")
    
    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, 10] == 'y' and check_table_df.iloc[line, 5] is np.nan):
            filename.write("--include-filter ")
            filename.write("\"")
            filename.write(check_table_df.iloc[line, 1])
            if(check_table_df.iloc[line, 0] is np.nan):
                filename.write("\" ")
                print("error")
            else:
                filename.write(" ")
                filename.write(check_table_df.iloc[line, 0])
                filename.write("\" ")
    filename.close()
    
    filename = open("cmd_8mn.txt", "a")
    filename.write("run cts ")
    
    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, 11] == 'y' and check_table_df.iloc[line, 5] is np.nan):
            filename.write("--include-filter ")
            filename.write("\"")
            filename.write(check_table_df.iloc[line, 1])
            if(check_table_df.iloc[line, 0] is np.nan):
                filename.write("\" ")
                print("error")
            else:
                filename.write(" ")
                filename.write(check_table_df.iloc[line, 0])
                filename.write("\" ")
    filename.close()
    
    

if __name__ == '__main__':
    # argvs: input_excel; test_type; 
    process_result(sys.argv[1], sys.argv[2])