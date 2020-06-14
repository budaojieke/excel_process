import logging

# logging.debug('debug级别，一般用来打印一些调试信息，级别最低')
# logging.info('info级别，一般用来打印一些正常的操作信息')
# logging.warning('waring级别，一般用来打印警告信息')
# logging.error('error级别，一般用来打印一些错误信息')
# logging.critical('critical级别，一般用来打印一些致命的错误信息，等级最高')
def set_log_level(log_level):
    logging.basicConfig(format='%(asctime)s - %(filename)s[%(lineno)d] - %(levelname)s: %(message)s',
                    level=log_level)

def logd(msg, *args):
    logging.debug(msg, *args)

def logi(msg, *args):
    logging.info(msg, *args)

def logw(msg, *args):
    logging.warning(msg, *args)

def loge(msg, *args):
    logging.error(msg, *args)