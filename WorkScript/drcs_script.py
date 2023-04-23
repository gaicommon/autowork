from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium import webdriver

from DataModel.drcs import DRCS
from Util import element


#  根据员工信息获取员工
def get_person(driver, input_person_code_id, person_code, div_person_list_xpath):
    Person_code_attr_name = "dataid"
    driver.find_element(By.ID, input_person_code_id).send_keys(person_code)
    element.find_el_click(driver, div_person_list_xpath, person_code, Person_code_attr_name, is_only=True)


def save_drcs(driver):
    Btn_save_id = 'ctl00_ContentPlaceHolder1_BtnSave'
    driver.find_element('id', Btn_save_id).click()
    Btn_save_confirm_xpath = '//div[@class="dialog-button messager-button"]/a'
    el = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
        lambda x: x.find_element(By.XPATH, Btn_save_confirm_xpath))
    el.click()
    i = 0
    while i < 10:
        els = driver.find_elements(By.XPATH, Btn_save_confirm_xpath)
        if els:
            els[0].click()
            driver.implicitly_wait(1)
        else:
            break


def create_drcs(drcs: DRCS):
    URL = url = r"http://aed.cnpdc.com/AED/Flow/DRCS/AED_DRCS_Edit.aspx?modi=add&ReUrl=http%3a%2f%2faed.cnpdc.com%2fAED%2fCommon%2fAED_WorkFlow%2fAED_NewDraft_List.aspx"
    driver = webdriver.Ie("IEDriverServer.exe")
    driver.implicitly_wait(2)
    driver.maximize_window()
    driver.get(URL)

    # region 填写表单页面

    Select_secret_id: str = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_DDL_Doc_SecClass"  # m密级选择

    try:
        # 等待
        WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
            lambda x: x.find_element('id', Select_secret_id))
    except:
        if driver.find_elements('id', 'otpTitle'):
            return element.Error["Wait_Token_Error"]
        else:
            return element.Error["Wait_Token_Error"]

    # region 密级、公司 语种和信函模板选择
    Select_company_id: str = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_DDL_Bas_COMP_CODE"  # 所属公司选择
    Select_language_id: str = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_DDL_Bas_Language"  # 语种选择
    Select_file_template_id: str = "stFileTemplate"  # 信函模板选择
    element.select_by_visible_text(driver, Select_secret_id, drcs.secret)
    element.select_by_visible_text(driver, Select_company_id, drcs.company)
    element.select_by_visible_text(driver, Select_language_id, drcs.language)
    element.select_by_visible_text(driver, Select_file_template_id, drcs.file_template)
    # endregion

    # region 项目选择
    Btn_project_open_id: str = "selectPrj"
    Wait_project_load_id = 'selecPrj'
    Frame_project_id = 'frmDialogselecPrj'
    Input_project_code_id = "ctl00_ContentPlaceHolder1_TB_Prj__txt1"
    Btn_lookup_project_xpath = '//table//tr//input[@class="btnFnt gOpt_Btn_Ok"]'
    Div_project_list_xpath = '//*[contains(@id,"datagrid-row-r2-2-")]/td[2]/div'
    Btn_project_save_xpath = '//div[@id="selecPrj"]/following-sibling::*/a[1]'
    prj_frame = element.Frame_select_attr()
    prj_frame.btn_click = [By.ID, Btn_project_open_id]
    prj_frame.wait_load_check_id = Wait_project_load_id
    prj_frame.iframe_id = Frame_project_id

    prj_frame.text_key = [By.ID, Input_project_code_id, drcs.project_code]
    prj_frame.el_list_xpath = Div_project_list_xpath
    prj_frame.btn_lookup = [By.XPATH, Btn_lookup_project_xpath]
    prj_frame.btn_confirm = [By.XPATH, Btn_project_save_xpath]
    element.get_frame_select(driver, prj_frame)
    # endregion
    # region 发送方和接收方选择
    # 发送方
    Btn_sender_open_xpath = '//input[@id="ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_fromTo_TB_Bas_FromFax"]/following-sibling::input'  # 发送方选择按钮
    Input_sender_name_id: str = 'ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_fromTo_TB_FromCodeText'  # 发送方名称id
    # 接收方
    Btn_receiver_open_id: str = 'addJieShouFang'  # 接收方选择——打开一个frame
    Input_receiver_name_id: str = 'ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_fromTo_TB_FromCodeText'  # 接收发方名称id

    # 发送方和接收方信框架的信息
    Wait_channel_load_id = 'OpenSelectFromToCode'  # 检查是否加载完成的元素
    Frame_channel_id = 'frmDialogOpenSelectFromToCode'  # 框架iframe的id
    Input_channel_code_id = 'txtKey'
    Btn_channel_lookup_xpath = '//*[@id="aspnetForm"]/div[3]/div[2]/div[1]/input[2]'
    Div_channel_list_xpath = '//*[contains(@id,"datagrid-row-r1-2-")]/td[3]/div'
    Btn_channel_save_xpath = '/html/body/div[13]/div[3]/a[1]'
    channel_frame = element.Frame_select_attr()

    # 框架信息
    channel_frame.wait_load_check_id = Wait_channel_load_id
    channel_frame.iframe_id = Frame_channel_id
    channel_frame.btn_lookup = [By.XPATH, Btn_channel_lookup_xpath]
    channel_frame.el_list_xpath = Div_channel_list_xpath
    channel_frame.btn_confirm = [By.XPATH, Btn_channel_save_xpath]

    # 发送方打开iframe按钮
    channel_frame.btn_click = [By.XPATH, Btn_sender_open_xpath]
    channel_frame.text_key = [By.ID, Input_channel_code_id, drcs.sender_code]
    element.get_frame_select(driver, channel_frame)

    # 接收方选择
    channel_frame.btn_click = [By.ID, Btn_receiver_open_id]
    channel_frame.text_key = [By.ID, Input_channel_code_id, drcs.receiver_code]
    element.get_frame_select(driver, channel_frame)

    # 添加发送方和接收方名称
    driver.find_element(By.ID, Input_sender_name_id).send_keys(drcs.sender_name)
    if drcs.receiver_name:
        driver.find_element(By.ID, Input_receiver_name_id).send_keys(drcs.receiver_name)

    # endregion

    # region 发文号等输入类信息
    # Inpiut_our_fax_code_xpath ='//*[@id="olQdh"]/li/input'                                                  # 我方发函号输入
    Input_replied_file_code_name = 'txtToFwbh0'  # 对方发文号
    Input_replyed_channel_code_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_TB_Otr_TransmitRef"  # 文件通道号输入
    Input_request_answer_date_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_Ref_AskFeedBackDate_dateTextBox"  # 要求答复日期选择——可直接输入
    # Input_attachment_code_id ="ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_TB_Otr_AttachFileNum"               # 所附文号
    Input_receiverd_date_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_TB_Otr_ReceiveDate_dateTextBox"  # 收文日期
    # Input_drcs_page_number_id="ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_TB_Doc_LetterCount__txt1"          # 正文页数
    Input_attachment_page_number_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_TB_Doc_AttCount__txt1"  # 附件页数
    # 对方发文号
    if drcs.replied_file_code:
        driver.find_element(By.NAME, Input_replied_file_code_name).send_keys(drcs.replied_file_code)
    # 文件通道号
    if drcs.replied_file_code:
        driver.find_element(By.ID, Input_replyed_channel_code_id).send_keys(drcs.replied_file_code)
    else:
        driver.find_element(By.ID, Input_replyed_channel_code_id).send_keys(drcs.channel_code)
    # 要求答复日期
    if drcs.request_reply_date:
        driver.find_element(By.ID, Input_request_answer_date_id).send_keys(drcs.request_reply_date)
    # 收文日期
    if drcs.received_date:
        driver.find_element(By.ID, Input_receiverd_date_id).send_keys(drcs.received_date)
    # 附件页数
    if drcs.attachment_page_number:
        driver.find_element(By.ID, Input_attachment_page_number_id).send_keys(drcs.attachment_page_number)

    # 标题和审查意见
    Input_title_id: str = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_Bas_Name__txt1"  # 标题输入
    Input_main_commen_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_Bas_Remark"  # 审核意见输入
    if drcs.drcs_title:
        driver.find_element(By.ID, Input_title_id).send_keys(drcs.drcs_title)
    driver.find_element(By.ID, Input_main_commen_id).send_keys(drcs.main_common)
    # endregion

    # region 审批人员
    Select_review_model_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_ddlSP"  # 审批模式选择
    Input_drafter_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_tbLHSJ_txt1"  # 联合编制人
    Input_checker_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_tbSH_txt1"  # 审核人选择
    Input_approver_id = "ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_tbPZ_txt1"  # 批准人选择

    Div_drafter_list_xpath = '//*[@id="ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_tbLHSJ_frm1"]/div/div'  # 起草人清单xpath
    Div_checker_list_xpath = '//*[@id="ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_tbSH_frm1"]/div/div'  # 审核人清单xpath
    Div_approver_list_xpath = '//*[@id="ctl00_ContentPlaceHolder1_AED_DRCS_Edit1_spMain_tbPZ_frm1"]/div/div'  # 批准人清单xpath

    # 选择审批模式
    element.select_by_visible_text(driver, Select_review_model_id, drcs.review_model)
    # 人员选择
    # 联合编制人
    if drcs.drafter_code:
        get_person(driver, Input_drafter_id, drcs.drafter_code, Div_drafter_list_xpath)

    if drcs.checker_code:
        get_person(driver, Input_checker_id, drcs.checker_code, Div_checker_list_xpath)

    if drcs.approver_code:
        get_person(driver, Input_approver_id, drcs.approver_code, Div_approver_list_xpath)

    # endregion

    # region 内部抄送和外部抄送
    Btn_inner_cc_xpath = '//*[@id="drcs"]//table[@class="formedit"]/tbody/tr[18]/td/div/input[@title="选择内部抄送"]'  # 内部抄送选择——打开一个frame
    Frame_inner_cc_id = "frmDialogOpenSelectCC"
    Input_frame_cc_group_id = "ctl00_ContentPlaceHolder1_TB_FFZ"  # 抄送的Frame id
    Btn_inner_cc_save_xpaht = '//div[@id="OpenSelectCC"]/following-sibling::*/a[1]'
    Wait_inner_cc_load_id = 'OpenSelectCC'
    if drcs.inner_cc:
        cc_frame = element.Frame_select_attr()

        # 框架信息
        cc_frame.wait_load_check_id = Wait_inner_cc_load_id
        cc_frame.iframe_id = Frame_inner_cc_id
        cc_frame.btn_confirm = [By.XPATH, Btn_inner_cc_save_xpaht]

        # 发送方打开iframe按钮
        cc_frame.btn_click = [By.XPATH, Btn_inner_cc_xpath]
        cc_frame.text_key = [By.ID, Input_frame_cc_group_id, drcs.inner_cc]
        element.get_frame_input(driver, cc_frame)

    """外部抄送"""
    Btn_outer_cc_xpath = '//input[@title="选择外部抄送"]'  # 外部抄送
    if drcs.outer_cc:
        # 发送方打开iframe按钮
        channel_frame.btn_click = [By.XPATH, Btn_outer_cc_xpath]
        channel_frame.text_key = [By.ID, Input_channel_code_id, drcs.outer_cc]
        element.get_frame_select(driver, channel_frame)
    # endregion

    # region 表单保存
    save_drcs(driver)
    print("表单保存成功")
    # endregion
    # endregion
    # region 处理文件清单标签页面
    Btn_document_list_tab_xpath = '//div[@id="drcs"]//div[@class="tabs-wrap"]/ul/li[2]/a'
    file_tab = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
        lambda x: x.find_element(By.XPATH, Btn_document_list_tab_xpath))
    file_tab.click()
    # 转至文件添加页面并验证加载完成
    Bnt_add_documents_xpath = '//div[@id="drcs"]//div[@class="datagrid-toolbar"]/table/tbody/tr/td[1]/a'
    add_file_dialog = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
        lambda x: x.find_element(By.XPATH, Bnt_add_documents_xpath))
    if add_file_dialog:
        add_file_dialog.click()
        el_addfile = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
            lambda x: x.find_element(By.ID, "addfile"))
        if not el_addfile:
            return False
        # 转到添加文件Frame
        Frame_select_documents_id = "frmDialogaddfile"
        driver.switch_to.frame(Frame_select_documents_id)
        Input_document_frame_channel_code_id = 'txtHJCode'
        Btn_document_frame_lookup_id = 'btnSearch'
        # 添加文件传递单通道号并点击查询
        driver.find_element(By.ID, Input_document_frame_channel_code_id).send_keys(drcs.channel_code)
        driver.find_element(By.ID, Btn_document_frame_lookup_id).click()

        # 展开通道号的文件：
        Btn_document_frame_show_document_xpath = '//td[@field="bi_oper"]/div/a'
        # 加载文件选择列表并选择
        el_show_flie_bt = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
            lambda x: x.find_element(By.XPATH, Btn_document_frame_show_document_xpath))
        if el_show_flie_bt:
            el_show_flie_bt.click()
            document_code_list = element.get_value_list_by_dickey(drcs.document_list, "code")  # w需审查文件代码列表
            Get_add_documents_xpath = '//tr[contains(@id,"datagrid-row-r1-2-")]'  # 获取文件列表信息
            el_show_flie_list = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
                lambda x: x.find_elements(By.XPATH, Get_add_documents_xpath))
            if el_show_flie_list:
                for tr in el_show_flie_list:
                    try:
                        tr_id = tr.get_attribute("id")
                        code_xpath = f'//tr[@id="{tr_id}"]/td[@field="code"]//a'
                        code_el = driver.find_element(By.XPATH, code_xpath)
                        document_code = code_el.get_attribute("textContent")
                        if document_code in document_code_list:
                            check_el_xpath = f'//tr[@id="{tr_id}"]/td[@field="ckk"]//input'
                            driver.find_element(By.XPATH, check_el_xpath).click()
                    except:
                        print("未添加成功")

        driver.switch_to.default_content()
        Btn_save_file_list_xpath = '//div[@id="addfile"]/following-sibling::*//a[1]'
        driver.find_element(By.XPATH, Btn_save_file_list_xpath).click()
        Get_documents_common_xpath = '//tr[contains(@id,"datagrid-row-r2-2-")]'
        el_flie_list = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
            lambda x: x.find_elements(By.XPATH, Get_documents_common_xpath))
        if el_flie_list:
            for tr in el_flie_list:
                tr_id = tr.get_attribute("id")
                code_xpath = f'//tr[@id="{tr_id}"]/td[@field="code"]//a'
                code_el = driver.find_element(By.XPATH, code_xpath)
                document_code = code_el.get_attribute("textContent")
                order = document_code_list.index(document_code.strip())
                if order != -1:
                    # result_id ='_easyui_textbox_input10'
                    # 审查结论
                    document_dic = drcs.document_list[order]
                    el_get_result_el_xpath = f'//tr[@id="{tr_id}"]//td[@field="fileresult"]'
                    driver.find_element(By.XPATH, el_get_result_el_xpath).click()
                    file_result_el_xpath = f'//tr[@id="{tr_id}"]//td[@field="fileresult"]//*[contains(@id,"easyui_textbox_input")]'
                    file_result_el = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
                        lambda x: x.find_element(By.XPATH, file_result_el_xpath))
                    file_result_el.send_keys(document_dic["result"])
                    # 审查意见
                    file_opinion_el_xpath = f'//tr[@id="{tr_id}"]/td[@field="fileopinion"]//*[contains(@id,"_easyui_textbox_input")]'
                    el_get_opinion_el_xpath = f'//tr[@id="{tr_id}"]//td[@field="fileopinion"]'
                    driver.find_element(By.XPATH, el_get_opinion_el_xpath).click()
                    file_opinion_el = WebDriverWait(driver, timeout=element.TimeOut, poll_frequency=0.2).until(
                        lambda x: x.find_element(By.XPATH, file_opinion_el_xpath))
                    file_opinion_el.send_keys(document_dic["common"])
    save_drcs(driver)
    print("审查意见添加成功！")
    # endregion
    # region 添加附件页面
    Btn_athach_tab_list_xpath = '//div[@class="tabs-wrap"]/ul/li[3]/a'  # 附件信息标签
    if (drcs.attachment_list):
        driver.find_element(By.XPATH, Btn_athach_tab_list_xpath).click()
        btn_sys_load_id = 'sysFileUpLoad'
        btn_sys_load_el = element.wait_el_load(driver, btn_sys_load_id)
        for file_dic in drcs.attachment_list:
            if (file_dic["is_sys_file"]):
                btn_sys_load_el.click()
                # 系统文件
                element.wait_el_load(driver, "addFileXTWJ", By.ID)
                driver.switch_to.frame('frmDialogaddFileXTWJ')
                sendfile_el = driver.find_element(By.ID, 'txtHJCode')
                sendfile_el.send_keys(file_dic['code'])
                driver.find_element(By.ID, 'btnSearch').click()

                check_el = element.wait_el_load(driver, '//tr[@id="datagrid-row-r1-2-0"]//input[@name="ck"]', By.XPATH)
                check_el.click()
                driver.switch_to.default_content()
                # 保存文件
                Btn_save_file_xpath = '//div[@id="addFileXTWJ"]/following-sibling::*//a[1]'
                element.confirm_click_close(driver, Btn_save_file_xpath, By.XPATH)
            else:
                btn_loadflle_xpath = '//td[@id="tdBYKJUpload"]/a'
                driver.find_element(By.XPATH, btn_loadflle_xpath).click()

                # 上传附件
                frame_dialog_div_id = 'StandbyUpload'
                element.wait_el_load(driver, frame_dialog_div_id)
                driver.switch_to.frame('frmDialogStandbyUpload')
                sendfile_el = driver.find_element(value='ctl00_ContentPlaceHolder1_uploadFile')
                sendfile_el.send_keys(file_dic['filepath'])
                driver.switch_to.default_content()
                # 保存文件
                Btn_save_file_xpath = '//div[@id="StandbyUpload"]/following-sibling::*//a[1]'
                element.confirm_click_close(driver, Btn_save_file_xpath, By.XPATH)

    save_drcs(driver)

    # endregion

    print("保存选择成功")
    return True
