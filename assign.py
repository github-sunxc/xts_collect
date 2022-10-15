# -*- coding: utf-8 -*-
import os
import sys
import datetime
import pandas
import numpy as np


def process_result(update_excel, owner_excel, test_type):
    owner_df  = pandas.read_excel(owner_excel)
    # print(owner_df.iloc[1, 2])
    update_df = pandas.read_excel(update_excel, sheet_name=test_type, skiprows=0)
    lines_to_update = len(update_df)
    lines_to_assign = len(owner_df)
    for owner_index in range(0: len(update_df)):
        if update_df['Owner'][owner_index] is np.nan:
        


if __name__ == '__main__':
    # argvs: update_excel; owner_excel; test_type; 
    process_result(sys.argv[1], sys.argv[2], sys.argv[3])