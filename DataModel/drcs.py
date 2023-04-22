from dataclasses import dataclass
from os import path
from Util.excel import Excel
from ProjVar.workvar import WorkData_path


@dataclass
class DRCS:
    # 需要用户输入的信息
    project_code: str = 'LF'
    receiver_code: str = "NHEY"
    sender_name: str = '中广核工程有限公司 设备采购与成套中心 核岛主设备分部 张斌 副经理'
    channel_code: str = None  # 文件通道号输入
    attachment_code: str = None  # 所附文号
    drcs_page_number: int = 2  # 正文页数
    attachment_page_number: int = None  # 附件页数

    drafter_code: str = None  # 联合编制人
    checker_code: str = None  # 审核人选择
    approver_code: str = None  # 批准人选择
    drcs_title: str = None  # 标题输入
    document_list: list = None
    attachment_list: list = None

    # 默认
    review_model: str = "并行（预设）"  # 审批模式选择
    secret: str = "内部公开"
    company: str = '工程公司'
    language: str = "中文"
    file_template: str = '设计审查单'
    outer_cc: str = None  # 外部抄送
    sender_code: str = "GENS"
    main_common: str = "详见详细意见"
    receiver_name: str = None  # 审核意见
    inner_cc: str = '#PPL;#PQC;#PQA'  # 内部抄送
    received_date: str = None  # 收文日期
    replied_file_code: str = channel_code  # 被答复函通道号
    request_reply_date: str = None


def getDrcsListByExcel(rows):
    pass


'''
    1、查找表格中的文件表格
    2、读取每一行的数据
    3、判断是否需要审查或已经审查
    4、生成DRCS列表,通道号作为key，value
'''


def getUnreviewDocument(excel_fileName):
    """
        获取Excel表格中包含'doc'的表格（文件表格）
        在各个表格中查找未审查且需要审查的文件列表
    """
    excel_path = path.join(WorkData_path, excel_fileName)
    excel = Excel(excel_path)
    sheet_names = excel.get_all_sheet_names()
    # 获取文件列表
    for name in sheet_names:
        if 'doc' in name:
            excel.sheet = excel.wb[name]
            excel.column_num = excel.sheet.max_column
            excel.row_num = excel.sheet.max_row
            attr_list = excel.get_excel_value_list(0)
            return attr_list


if __name__ == "__main__":
    excel_path = "projectdocument_EHEC.xlsx"
