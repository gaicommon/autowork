from configparser import ConfigParser

from ProjVar.workvar import Config_path


def getConfigValue(section: str, valueKey: str):
    conf = ConfigParser()
    conf.read(filenames=Config_path, encoding='utf-8')

    return conf.get(section, valueKey)


# def writeConfigValue(section: str, valueKey: str):
#     conf = ConfigParser()
#     conf.write(filenames=Config_path, encoding='utf-8')
#
#     # return conf.get(section, valueKey, value)

if __name__ == "__main__":
    print(getConfigValue('sender_info', 'sender_name'))
