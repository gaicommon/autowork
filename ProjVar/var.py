import os

# 当前文件所在目录的父目录的绝对路径
ProjDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 浏览器驱动存放的位置
# firefoxDriverFilePath = os.path.join(ProjDirPath, r"driver\geckodriver.exe")
# chromeDriverFilePath = os.path.join(ProjDirPath, r"driver\chromedriver.exe")
IE32DriverFilePath = os.path.join(ProjDirPath, r"driver\IEDriverServer32.exe")
IE64DriverFilePath = os.path.join(ProjDirPath, r"driver\IEDriverServer.exe")
# 配置文件地址


test_step_id_col_no = 0
test_step_result_col_no = 8
test_step_error_info_col_no = 9
test_step_capture_pic_path_col_no = 10

test_case_id_col_no = 0
test_case_sheet_name = 2
test_case_is_executed_flag_col_no = 3
test_case_is_hybrid_test_data_sheet_col_no = 4

test_case_start_time_col_no = 6
test_case_end_time_col_no = 7
test_case_elapsed_time_col_no = 8
test_case_result_col_no = 9

if __name__ == "__main__":
    print(Config_path)
