# logging_config.py
"""
Модуль для настройки логирования приложения.
"""

import logging
import os
import sys  # Добавлено для PyInstaller
from datetime import datetime
# Импортируем LOG_DIR из нового config.py
from config import LOG_DIR


class LogManager:
    """
    Класс для управления логированием приложения.
    Настраивает логирование как в файл, так и в консоль.
    """

    def __init__(self, xlsx_filename: str, log_level: int = logging.INFO):
        """
        Инициализирует менеджер логирования.

        Args:
            xlsx_filename (str): Путь к обрабатываемому Excel-файлу.
            log_level (int, optional): Уровень логирования. По умолчанию logging.INFO.
        """
        self.xlsx_filename = xlsx_filename
        self.log_level = log_level
        self.log_file = self._setup_logging()

    def _setup_logging(self) -> str:
        """
        Настраивает систему логирования.

        Создает директорию для логов (если не существует) и файл лога с именем,
        основанным на имени Excel-файла и текущей дате.

        Returns:
            str: Полный путь к созданному лог-файлу.
        """
        # Используем os.path.basename для надежности
        base = os.path.splitext(os.path.basename(self.xlsx_filename))[0]
        date = datetime.now().strftime("%d.%m.%Y")

        # Учитываем PyInstaller при создании пути к директории логов
        if getattr(sys, 'frozen', False):
            # Если .exe, создаем log в той же папке, что и .exe
            log_dir_full = os.path.join(
                os.path.dirname(sys.executable), LOG_DIR)
        else:
            log_dir_full = LOG_DIR

        os.makedirs(log_dir_full, exist_ok=True)
        log_file = os.path.join(log_dir_full, f"{base}_{date}.log")

        log_format = "%(asctime)s [%(levelname)s] %(message)s"
        date_format = "%d.%m.%Y"

        # Проверка, не настроено ли уже логирование
        if not logging.getLogger().handlers:
            logging.basicConfig(
                level=self.log_level,
                format=log_format,
                datefmt=date_format,
                handlers=[
                    logging.FileHandler(log_file, encoding="utf-8"),
                    logging.StreamHandler()
                ]
            )
        return log_file

    def get_logger(self, name: str = __name__) -> logging.Logger:
        """
        Получает экземпляр логгера.

        Args:
            name (str, optional): Имя логгера. По умолчанию __name__.

        Returns:
            logging.Logger: Экземпляр логгера.
        """
        return logging.getLogger(name)

# --- Пример использования ---
# logger_manager = LogManager("example.xlsx")
# logger = logger_manager.get_logger(__name__)
# logger.info("Это сообщение лога")
