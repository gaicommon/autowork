from os import path


from ProjVar.var import ProjDirPath
from Conf.config import getConfigValue
Sender_name = getConfigValue('sender_info', 'sender_name')
WorkData_path = path.join(ProjDirPath, 'WorkData')

