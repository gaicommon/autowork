# encoding=utf-8
import logging.config
import logging
from ProjVar.var import ProjDirPath

print(ProjDirPath + "\\Conf\\" + "Logger.conf")

# 读取日志配置文件
logging.config.fileConfig(ProjDirPath + "\\Conf\\" + "Logger.conf")

# 选择一个日志格式
logger = logging.getLogger("example01")


def debug(message):
    print("debug")
    logger.debug(message)


def info(message):
    logger.info(message)


def error(message):
    print("error")
    logger.error(message)


def warning(message):
    logger.warning(message)


if __name__ == "__main__":
    debug("hi")
    info("gloryroad")
    warning("hello")
    error("something error!")
