from os import path


from ProjVar.var import ProjDirPath
from Util.config import getConfigValue
Config_path = path.join(ProjDirPath, 'Conf', 'config.ini')
Sender_name = getConfigValue('sender_info', 'sender_name')
WorkData_path = path.join(ProjDirPath, 'WorkData')

