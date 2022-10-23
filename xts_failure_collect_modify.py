# -*- coding: utf-8 -*-
import os
import sys
import datetime
import pandas

def save_df_to_excel(new_case_df, sheet_name, board_type):
    '''
    Refactoring the function, all failed case will add into 'new_fail_case.xlsx', 
    will not modify the old table. 
    '''
    filename = 'new_fail_case.xlsx'
    if (os.path.isfile(filename)) :
        print('===>file %s not exist, creat it now' %(filename))
        print('===>start to write data into %s' %(filename))
        new_case_df.to_excel(filename + '.xlsx', sheet_name = sheet_name, index = False)
    else :
        df = pandas.read_excel(filename,sheet_name=None)
        writer = pandas.ExcelWriter(filename)

        print('===>start to write data into %s' %(filename))
        for sheet in df.keys():
            if(sheet == sheet_name):
                write_df = df[sheet]
                for old_item_index in len(df[sheet]) :
                    flag = 0
                    for new_item_index in len(new_case_df) :
                        if (df[sheet].iloc[old_item_index, 0] == new_case_df.iloc[new_item_index, 0] and 
                        df[sheet].iloc[old_item_index, 1] == new_case_df.iloc[new_item_index, 1]) :
                            write_df[board_type][old_item_index] = 'y'
                            flag = 1
                            break
                    if (flag == 0) :
                        write_df = pandas.concat([write_df, new_case_df.iloc[new_item_index: new_item_index+1]])
                write_df.to_excel(writer,index=False,sheet_name=sheet)
            else :
                df[sheet].to_excel(writer,index=False,sheet_name=sheet)

        writer.save()


def write_excel(failure_data, need_check_table, test_type, board_type):
    failure_result = parse_test_failure_html(failure_data)
    failure_result_df = pandas.DataFrame(
        columns=["Module", "Case Name"], data=failure_result)

    # remove duplicate cases for armv7 and armv8
    failure_result_df.drop_duplicates(
        ['Case Name'], keep='first', inplace=True)
    failure_result_df.reset_index(drop=True, inplace=True)

    failure_result_to_add_df = pandas.DataFrame(data=None, index=None, columns=[
                                                "Case Name", "Module", "Owner", "Status", "Detail", "RC version", "8QM", "8QXP", "8MM", "8MP", "8MQ", "8MN", "8ULP", "8ULP9"])
    failure_result_to_add_df["Module"] = failure_result_df["Module"]
    failure_result_to_add_df["Case Name"] = failure_result_df["Case Name"]
    failure_result_to_add_df[board_type] = "y"
    lines_to_check = len(failure_result_to_add_df)

    after_compare_need_to_add_cases = failure_result_to_add_df

    # TODO should add judge if the file do not exist
    need_check_table_df = pandas.read_excel(
        need_check_table, sheet_name=test_type, skiprows=0)
    lines_need_check_table_df = len(need_check_table_df)

    if(board_type == "8QM"):
        board_type_column = 6  # G column
    elif(board_type == "8QXP"):
        board_type_column = 7  # H column
    elif(board_type == "8MM"):
        board_type_column = 8  # I column
    elif(board_type == "8MP"):
        board_type_column = 9  # J column
    elif(board_type == "8MQ"):
        board_type_column = 10  # K column
    elif(board_type == "8MN"):
        board_type_column = 11  # L column
    elif(board_type == "8ULP"):
        board_type_column = 12  # M column
    elif(board_type == "8ULP9"):
        board_type_column = 13  # N column

    new_case_df = pandas.DataFrame(data=None, index=None, columns=[
                                                "Case Name", "Module", "Owner", "Status", "Detail", "RC version", "8QM", "8QXP", "8MM", "8MP", "8MQ", "8MN", "8ULP", "8ULP9"])

    if(lines_need_check_table_df > 0):
    
        for need_update_df_num in range(0, lines_to_check):
            flag = 0
            for to_check_case_num in range(0, lines_need_check_table_df):
                if (need_check_table_df.iloc[to_check_case_num, 0] == failure_result_to_add_df.iloc[need_update_df_num, 0] and 
                need_check_table_df.iloc[to_check_case_num, 1] == failure_result_to_add_df.iloc[need_update_df_num, 1]):
                    # Recorded case had been fixed by owner, it seems not essential to mark it again. 
                    # but stay this step will not influent our output. 
                    # need_check_table_df.iloc[to_check_case_num, board_type_column] = 'y'
                    # new_case_df = pandas.concat([new_case_df, need_check_table_df.iloc[to_check_case_num: to_check_case_num+1]], ignore_index=True)
                    flag = 1
                    break
            if (flag == 0): 
                # Record new happened case only
                new_case_df = pandas.concat([new_case_df,
                failure_result_to_add_df.iloc[need_update_df_num:need_update_df_num+1]], ignore_index=True)

        # Sort to get a more convenience review
        new_case_df = new_case_df.sort_values(by='Module' ,kind='mergesort')

        save_df_to_excel(new_case_df,test_type,board_type)

    else: # empty table, dump fail case into it
        save_df_to_excel(after_compare_need_to_add_cases,test_type,board_type)

    return failure_result


def html2dataframe(html_file):
    '''
    process html to datafram with pandas
    '''
    print("===>start to pandas process: ", html_file)
    with open(html_file, 'r') as fh:
        # content = pandas.read_html(f.read(), encoding='utf-8')
        content = pandas.read_html(fh.read())
    return content


def html2dataframe_by_url(url):
    '''
    process html to datafram with pandas
    '''
    print("===>start to pandas process: ", url)
    try:
        return pandas.read_html(url)
    except:
        return None


def get_failed_div(table_df):
    div_name = get_table_div_name(table_df)
    if not div_name in ["incomplete modules", "module", "summary"]:
        if table_df.columns.size == 3:
            section_keys = ("Test", "Result", "Details")
            if section_keys == tuple(table_df.iloc[1].values.tolist()):
                div_name = "failed modules"
    return div_name


def get_table_div_name(table_df):
    columns = table_df.columns.tolist()
    return str(columns[0]).lower()


def get_cols_from_row(each_table_df, row):
    if row is None:
        first_row = each_table_df.columns.tolist()
    else:
        first_row = each_table_df.iloc[row].values.tolist()
    headers = [str(header).lower() for header in first_row]
    return tuple(headers)


def format_html_table_df(html_file):
    html2df_list = html2dataframe(html_file)
    html_div_data = {}
    for each_table_df in html2df_list:
        div_name = get_failed_div(each_table_df)

        col_row = None
        if div_name == "failed modules":
            col_row, drop_row = 1, 1

        each_table_df.columns = get_cols_from_row(each_table_df, col_row)
        if div_name == "failed modules":
            each_table_df = each_table_df.drop(drop_row)
        html_div_data.setdefault(div_name, [])
        html_div_data[div_name].append(each_table_df)
    # print(html_div_data)
    return html_div_data


def get_latest_test_failure_html(results_home):
    result_dirs = [dir for dir in os.listdir(results_home) if os.path.isdir(
        os.path.join(results_home, dir)) and dir != "latest"]
    result_dirs.sort(key=lambda date: datetime.datetime.strptime(
        date, '%Y.%m.%d_%H.%M.%S'), reverse=True)  # 倒序排列
    if not result_dirs:
        return None

    # find latest test_result_failures_suite.html
    test_result_failure_list = []
    for result_dir in result_dirs:
        latest_result_dir = os.path.join(results_home, result_dir)
        #ziped_latest_result = latest_result_dir + ".zip"
        test_result_failures = os.path.join(
            latest_result_dir, "test_result_failures_suite.html")
        while True:
            # if os.path.isfile(ziped_latest_result) and os.path.isfile(test_result_failures): break
            if os.path.isfile(test_result_failures):
                break
        test_result_failure_list.append(test_result_failures)
    #print("===>get all failure htmls")
    # print(test_result_failure_list)
    if test_result_failure_list:
        # print(test_result_failure_list[len_result_dirs-1])
        print("===>get latest failure htmls")
        # print(test_result_failure_list[0])
        return test_result_failure_list
    return None


def get_session(failure_files):
    return len(failure_files) - 1


def process_result(results_home, need_check_table, test_type, board_type):
    test_failure_html_list = get_latest_test_failure_html(results_home)
    if test_failure_html_list is not None:
        print("===>start to format the latest failure html file")
        failure_data = format_html_table_df(test_failure_html_list[0])

        write_excel(failure_data, need_check_table, test_type, board_type)
    return []


def parse_test_failure_html(failure_data):
    fail_cnt = 0
    test_suite_result = []
    # process_module_result_df(failure_data["module"])
    # process_summary_result_df(failure_data["summary"])
    failure_result = []

    if "failed modules" in failure_data:
        print("===>start to parse failed modules and cases section")
        # result_details = ["test", "result", "details"]
        result_details = ["test"]
        for each_table_df in failure_data["failed modules"]:
            failed_module = str(each_table_df.iloc[0, 0])
            abi_module = failed_module.split()
            #abi = abi_module[0].strip()
            module_name = abi_module[1].strip()

            for indexs in each_table_df.index:
                if indexs == 0:
                    continue
                if indexs > len(each_table_df):
                    break
                failed_case = [module_name, str(
                    each_table_df.loc[indexs, result_details]).strip()]
                if not failed_case in failure_result:
                    failure_result.append(
                        [module_name, str(each_table_df.loc[indexs, "test"]).strip()])
    return failure_result


def process_module_result_df(module_df):
    each_module_df = module_df[0]
    print("===>modules result section")

    module_result = {}
    result_types = ["module", "failed", "passed", "total tests", "done"]
    for indexs in each_module_df.index:
        if indexs > len(each_module_df):
            break
        for result_type in result_types:
            module_result[result_type] = str(
                each_module_df.loc[indexs, result_type]).strip()
        # print(module_result)


if __name__ == '__main__':
    # argvs:test_result_path;need_check_table;test_type;board_type
    process_result(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
