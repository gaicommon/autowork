import re


def get_info_from_channel_code(channel_code: str):
    data_list = channel_code.split('-')
    if len(data_list) == 7:
        return {"project_code": data_list[1], 'sender_code': data_list[3], 'receiver_code': data_list[5]}
    if len(data_list) == 4:
        return {"project_code": data_list[0], 'sender_code': data_list[1], 'receiver_code': data_list[3]}


def str_to_list(list_str: str):
    return re.split(',|，|;|；|\n')


if __name__ == "__main__":
    channel_code = "SK-ANE-700000-HEY"
    print(get_info_from_channel_code(channel_code))
