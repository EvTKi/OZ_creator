# logging_config.py
import logging
import os
from datetime import datetime


def setup_logging(xlsx_filename: str):
    """
    Создаёт папку для логов и лог-файл с именем:
      [наименование файла]_[дд.мм.гггг].log
    """
    # Имя без расширения
    base = os.path.splitext(os.path.basename(xlsx_filename))[0]
    date = datetime.now().strftime("%d.%m.%Y")
    log_dir = "./log"
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"{base}_{date}.log")

    log_format = "%(asctime)s [%(levelname)s] %(message)s"
    date_format = "%d.%m.%Y"
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        datefmt=date_format,
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return log_file
