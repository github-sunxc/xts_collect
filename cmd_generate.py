# -*- coding: utf-8 -*-
import os
import sys
import pandas
import numpy as np


def process_result(input_excel, test_type, board):
    check_table_df  = pandas.read_excel(input_excel, sheet_name=test_type, skiprows=0)
    check_table_len = len(check_table_df)

    if (board == "8QM"):
        board_num = 6
    elif (board == "8QXP"):
        board_num = 7
    elif (board == "8MM"):
        board_num = 8
    elif (board == "8MP"):
        board_num = 9
    elif (board == "8MQ"):
        board_num = 10
    elif (board == "8MN"):
        board_num = 11
    elif (board == "8ULP"):
        board_num = 12
    elif (board == "8ULP9"):
        board_num = 13
    else :
        print("board error")
        sys.exit()

    command_head = 'run '
    if (test_type.lower() == 'cts'):
        command_head = command_head + 'cts '
    elif (test_type.lower() == 'cts-on-gsi'):
        command_head = command_head + 'cts-on-gsi '
    elif (test_type.lower() == 'vts'):
        command_head = command_head + 'vts '
    elif (test_type.lower() == 'gts'):
        command_head = command_head + 'gts '
    elif (test_type.lower() == 'sts'):
        command_head = command_head + 'sts-dynamic-full '
    else :
        print("test type error")
        sys.exit()
        
    filename = "cmd_{0}.txt".format(board)
    print('===>Start write command in %s' %(filename))

    cmd_file = open(filename, "w")
    cmd_file.write(command_head)

    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, board_num] == 'y') and pandas.isnull(check_table_df.iloc[line, 5]):
            cmd_file.write("--include-filter ")
            cmd_file.write("\"")
            cmd_file.write(check_table_df.iloc[line, 1])
            if(check_table_df.iloc[line, 0] is np.nan):
                cmd_file.write("\" ")
                print("Incomplete")
            else:
                cmd_file.write(" ")
                cmd_file.write(check_table_df.iloc[line, 0])
                cmd_file.write("\" ")
    cmd_file.close()


if __name__ == '__main__':
    # argvs: input_excel; test_type; board
    process_result(sys.argv[1], sys.argv[2], sys.argv[3])
