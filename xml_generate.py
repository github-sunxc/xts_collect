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

    filename = "{0}_{1}.xml".format(board, test_type)
    print('===>Start write command in %s' %(filename))

    xml_head = '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'no\' ?>\n<SubPlan version=\"2.0\">\n'
    xml_tail = '</SubPlan>'
    line_head = '  <Entry include=\"'
    line_tail = '\" />\n'
    
    xml_file = open(filename, "w")
    xml_file.write(xml_head)

    for line in range(0, check_table_len):
        if (check_table_df.iloc[line, board_num] == 'y') and pandas.isnull(check_table_df.iloc[line, 5]):
            xml_file.write(line_head)
            xml_file.write(check_table_df.iloc[line, 1])
            if(pandas.isnull(check_table_df.iloc[line, 0])):
                print("Incomplete")
            else:
                xml_file.write(" ")
                xml_file.write(check_table_df.iloc[line, 0])
            xml_file.write(line_tail)
    xml_file.write(xml_tail)
    xml_file.close()


if __name__ == '__main__':
    # argvs: input_excel; test_type; board
    process_result(sys.argv[1], sys.argv[2], sys.argv[3])
