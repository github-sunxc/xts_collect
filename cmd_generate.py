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

    filename = "cmd_{0}.txt".format(board)
    print('===>Start write command in %s' %(filename))

    filename = open("cmd_8qm.txt", "a")
    filename.write("run cts ")

    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, board_num] == 'y' and check_table_df.iloc[line, 5] is np.nan):
            filename.write("--include-filter ")
            filename.write("\"")
            filename.write(check_table_df.iloc[line, 1])
            if(check_table_df.iloc[line, 0] is np.nan):
                filename.write("\" ")
                print("Incomplete")
            else:
                filename.write(" ")
                filename.write(check_table_df.iloc[line, 0])
                filename.write("\" ")
    filename.close()

    filename = open("cmd_8mp.txt", "a")
    filename.write("run cts ")


if __name__ == '__main__':
    # argvs: input_excel; test_type; board
    process_result(sys.argv[1], sys.argv[2], sys.argv[3])
