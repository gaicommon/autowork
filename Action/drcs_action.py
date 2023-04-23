from DataModel import drcs

from WorkScript.drcs_script import create_drcs


def review_doc(excel_path_name):
    doc_info_dic = drcs.getUnreviewDocument(excel_path_name)
    excel = doc_info_dic['excel']
    doc_dic= doc_info_dic['doc_list']
    if doc_dic:
        for name in doc_dic.keys():
            channel_dic = doc_dic[name]
            for channel_code in channel_dic.keys():
                create_drcs(channel_dic[channel_code])


