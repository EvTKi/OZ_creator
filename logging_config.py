# logging_config.py
"""
Модуль для настройки логирования приложения.
"""

import logging
import os
import sys
from datetime import datetime
from config import LOG_DIR


class LogManager:
    """
    Класс для управления логированием приложения.
    Настраивает логирование как в файл, так и в консоль.
    """

    def __init__(self, log_level: int = logging.INFO):  # Изменено: убираем filename из __init__
        """
        Инициализирует менеджер логирования.

        Args:
            log_level (int, optional): Уровень логирования. По умолчанию logging.INFO.
        """
        self.log_level = log_level
        self.logger_configured = False
        self._setup_base_logging()

    def _setup_base_logging(self):
        """Настраивает базовое логирование (консоль). Файловый хендлер будет добавлен позже."""
        if not self.logger_configured:
            log_format = "%(asctime)s [%(levelname)s] %(message)s"
            date_format = "%d.%m.%Y %H:%M:%S"

            console_handler = logging.StreamHandler()
            console_handler.setFormatter(
                logging.Formatter(log_format, datefmt=date_format))

            root_logger = logging.getLogger()
            root_logger.setLevel(self.log_level)
            root_logger.addHandler(console_handler)

            self.logger_configured = True

    def setup_file_handler(self, base_filename: str):
        """
        Добавляет файловый хендлер к логгеру с именем, основанным на base_filename.

        Args:
            base_filename (str): Имя выходного файла (без пути и расширения),
                                 например, 'output_data' для файла 'output_data.xml'.
        """
        # Учитываем PyInstaller при создании пути к директории логов
        if getattr(sys, 'frozen', False):
            log_dir_full = os.path.join(
                os.path.dirname(sys.executable), LOG_DIR)
        else:
            log_dir_full = LOG_DIR

        os.makedirs(log_dir_full, exist_ok=True)

        date = datetime.now().strftime("%d.%m.%Y")
        log_file = os.path.join(log_dir_full, f"{base_filename}_{date}.log")

        # Проверяем, не добавлен ли уже файловый хендлер
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            if isinstance(handler, logging.FileHandler):
                root_logger.removeHandler(handler)
                handler.close()

        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        log_format = "%(asctime)s [%(levelname)s] %(message)s"
        date_format = "%d.%m.%Y %H:%M:%S"
        file_handler.setFormatter(logging.Formatter(
            log_format, datefmt=date_format))

        root_logger.addHandler(file_handler)
        logging.info(
            f"Файловый логгер настроен. Логи будут записываться в: {log_file}")

    def get_logger(self, name: str = __name__) -> logging.Logger:
        """
        Получает экземпляр логгера.

        Args:
            name (str, optional): Имя логгера. По умолчанию __name__.

        Returns:
            logging.Logger: Экземпляр логгера.
        """
        return logging.getLogger(name)
