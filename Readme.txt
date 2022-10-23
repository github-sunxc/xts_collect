1、运行xts_failure_collect_modify.py，检查当前路径下是否存在new_fail_case.xlsx文件，如果不存在，创建，如果存在，读取
2、提取出新的fail case，写入new_fail_case.xlsx的对应sheet中，去重
3、与之前的版本对比，替换掉以前出现过的case，标记version
4、分配
5、手动复制到teams里面的表格后面