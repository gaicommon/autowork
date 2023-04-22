# encoding=utf-8


import re
from Util.diropration import make_current_date_dir, make_current_hour_dir
from Util.excel import *
from Action.pageaction import *
import traceback
from Util.log import *


def get_test_case_sheet(test_cases_excel_path):
    test_case_sheet_names = []
    excel_obj = Excel(test_cases_excel_path)
    excel_obj.set_sheet_by_index(1)
    test_case_rows = excel_obj.get_rows_object()[1:]
    for row in test_case_rows:
        if row[3].value == 'y':
            print(row[2].value)
            test_case_sheet_names.append((int(row[0].value) + 1, row[2].value))
    return test_case_sheet_names


# print(get_test_case_sheet(ProjDirPath+"\\TestData\\126邮箱联系人.xlsx"))

def execute(test_cases_excel_path, row_no, test_case_sheet_name):
    excel_obj = Excel(test_cases_excel_path)
    excel_obj.set_sheet_by_name(test_case_sheet_name)
    # 获取除第一行之外的所有行对象
    test_step_rows = excel_obj.get_rows_object()[1:]
    # 拼接开始时间：当前年月日+当前时间
    start_time = get_current_date_and_time()
    # 开始计时
    start_time_stamp = time.time()
    # 设置默认用例时执行成功的，如果抛异常则说明用例执行失败
    test_result_flag = True
    for test_step_row in test_step_rows:
        if test_step_row[6].value == "y":
            test_action = test_step_row[2].value
            locator_method = test_step_row[3].value
            locator_exp = test_step_row[4].value
            test_value = test_step_row[5].value
            print(test_action, locator_method, locator_exp, test_value)
            if locator_method is None:
                if test_value is None:
                    command = test_action + "()"
                else:
                    command = test_action + "('%s')" % test_value
            else:
                if test_value is None:
                    command = test_action + "('%s','%s')" % (locator_method, locator_exp)
                else:
                    command = test_action + "('%s','%s','%s')" % (locator_method, locator_exp, test_value)
            print(command)
            eval(command)
            # 处理异常
            try:
                info(command)
                eval(command)
                excel_obj.write_cell_value(int(test_step_row[0].value) + 1, test_step_result_col_no, "执行成功")
                excel_obj.write_cell_value(int(test_step_row[0].value) + 1, test_step_error_info_col_no, "")
                excel_obj.write_cell_value(int(test_step_row[0].value) + 1, test_step_capture_pic_path_col_no, "")
                info("执行成功")  # 加入日志信息
            except Exception as e:
                test_result_flag = False
                traceback.print_exc()
                error(command + ":" + traceback.format_exc())  # 加入日志信息
                excel_obj.write_cell_value(int(test_step_row[0].value) + 1, test_step_result_col_no, "失败", "red")
                excel_obj.write_cell_value(int(test_step_row[0].value) + 1, test_step_error_info_col_no, \
                                           command + ":" + traceback.format_exc())
                dir_path = make_current_date_dir(ProjDirPath + "\\" + "ScreenCapture\\")
                dir_path = make_current_hour_dir(dir_path + "\\")
                pic_path = os.path.join(dir_path, get_current_time() + ".png")
                capture(pic_path)
                excel_obj.write_cell_value(int(test_step_row[0].value) + 1, test_step_capture_pic_path_col_no, pic_path)
        # 拼接结束时间：年月日+当前时间
        end_time = get_current_date() + " " + get_current_time()
        # 计时结束时间
        end_time_stamp = time.time()
        # 执行用例时间等于结束时间-开始时间；需要转换数据类型int()
        elapsed_time = int(end_time_stamp - start_time_stamp)
        # 将时间转换成分钟；整除60
        elapsed_minutes = int(elapsed_time // 60)
        # 将时间转换成秒；除60取余
        elapsed_seconds = elapsed_time % 60
        # 拼接用例执行时间；分+秒
        elapsed_time = str(elapsed_minutes) + "分" + str(elapsed_seconds) + "秒"
        # 判断用例是否执行成功
        if test_result_flag:
            test_case_result = "测试用例执行成功"
        else:
            test_case_result = "测试用例执行失败"
        # 需要写入的时第一个sheet
        excel_obj.set_sheet_by_index(1)
        # 写入开始时间
        excel_obj.write_cell_value(int(row_no), test_case_start_time_col_no, start_time)
        # 写入结束时间
        excel_obj.write_cell_value(int(row_no), test_case_end_time_col_no, end_time)
        # 写入执行时间
        excel_obj.write_cell_value(int(row_no), test_case_elapsed_time_col_no, elapsed_time)
        # 写入执行结果；
        if test_result_flag:
            # 用例执行成功，写入执行结果
            excel_obj.write_cell_value(int(row_no), test_case_result_col_no, test_case_result)
        else:
            # 用例执行失败，用red字体写入执行结果
            excel_obj.write_cell_value(int(row_no), test_case_result_col_no, test_case_result, "red")

    # 清理excel记录的结果数据
    def clear_test_data_file_info(test_data_excel_file_path):
        excel_obj = Excel(test_data_excel_file_path)
        excel_obj.set_sheet_by_index(1)
        test_case_rows = excel_obj.get_rows_object()[1:]
        for test_step_row in test_case_rows:
            excel_obj.set_sheet_by_index(1)
            if test_step_row[test_case_is_executed_flag_row_no].value == "y":
                excel_obj.write_cell_value(
                    int(test_step_row[test_case_id_col_no].value) + 1, test_case_start_time_col_no, "")
                excel_obj.write_cell_value(
                    int(test_step_row[test_case_id_col_no].value) + 1, test_case_end_time_col_no, "")
                excel_obj.write_cell_value(
                    int(test_step_row[test_case_id_col_no].value) + 1, test_case_elapsed_time_col_no, "")
                excel_obj.write_cell_value(
                    int(test_step_row[test_case_id_col_no].value) + 1, test_case_result_col_no, "")

                excel_obj.set_sheet_by_name(test_step_row[test_case_sheet_name].value)
                test_step_rows = excel_obj.get_rows_object()[1:]
                for test_step_row in test_step_rows:
                    if test_step_row[test_step_id_col_no].value is None:
                        continue
                    excel_obj.write_cell_value(
                        int(test_step_row[test_step_id_col_no].value) + 1, test_step_result_col_no, "")
                    excel_obj.write_cell_value(
                        int(test_step_row[test_step_id_col_no].value) + 1, test_step_error_info_col_no, "")
                    excel_obj.write_cell_value(
                        int(test_step_row[test_step_id_col_no].value) + 1, test_step_capture_pic_path_col_no, "")


if __name__ == "__main__":
    test_data_excel_file_path = ProjDirPath + "\\TestData\\126邮箱联系人.xlsx"
    print(get_test_case_sheet(test_data_excel_file_path))
    execute(test_data_excel_file_path, 2, "登录")
