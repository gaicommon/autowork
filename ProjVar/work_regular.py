def get_info_from_channel_code(channel_code: str):
    data_list = channel_code.split('-')
    if len(data_list) == 7:
        return {"project_code": data_list[1], 'sender_code': data_list[3], 'receiver_code': data_list[5]}


if __name__ == "__main__":
    channel_code = "007-LF-L-NHEY-F-GENS-700131"
    print(get_info_from_channel_code(channel_code))
