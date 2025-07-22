import logging
from datetime import datetime

# Лог-файл по имени задачи, с датой
def setup_logger(log_name):
    logger = logging.getLogger(log_name)
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(f'{log_name}_{datetime.now():%Y%m%d}.log', encoding='utf-8')
    fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))
    if not logger.handlers:
        logger.addHandler(fh)
    return logger
