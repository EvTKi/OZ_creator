# tests/test_logging_config.py
"""
Тесты для модуля logging_config.py.
"""
import pytest
import logging
import os
import sys
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch, MagicMock

# Импорты модулей проекта
from logging_config import LogManager
from config import LOG_DIR  # Импортируем путь к папке логов из config


class TestLogManager:
    """Тесты для класса LogManager."""

    # Используем временную директорию для каждого теста, где будут создаваться логи
    @pytest.fixture(autouse=True)
    def setup_and_teardown(self):
        """Создает временную директорию перед тестом и удаляет её после."""
        # Создаем временную директорию
        self.test_dir = tempfile.mkdtemp()

        # --- Исправлено: Патчим LOG_DIR на АБСОЛЮТНЫЙ путь к временной директории ---
        # Вместо относительного пути "test_logs_for_test" патчим на саму временную папку
        with patch('logging_config.LOG_DIR', self.test_dir):
            yield  # Тест выполняется здесь
        # После теста удаляем временную директорию
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_init_sets_up_console_logging(self):
        """Тест: Инициализация LogManager настраивает консольный логгер."""
        # 1. Подготовка: Очистим обработчики корневого логгера перед тестом для изоляции
        root_logger = logging.getLogger()
        original_handlers = root_logger.handlers[:]
        for handler in original_handlers:
            root_logger.removeHandler(handler)

        try:
            # 2. Действие
            # Используем DEBUG для проверки
            lm = LogManager(log_level=logging.DEBUG)
            # logger = lm.get_logger("test_init_console") # Не проверяем уровень этого логгера

            # 3. Утверждение
            # Проверяем, что уровень установлен для КОРНЕВОГО логгера
            assert root_logger.level == logging.DEBUG
            # Проверяем, что в корневом логгере есть хотя бы один обработчик (консольный)
            # Должен быть как минимум консольный
            assert len(root_logger.handlers) >= 1
            # Проверим, что среди обработчиков есть StreamHandler (консоль)
            console_handlers = [
                h for h in root_logger.handlers if isinstance(h, logging.StreamHandler)]
            assert len(console_handlers) >= 1
            # Проверим форматтер (простая проверка наличия)
            assert console_handlers[0].formatter is not None
            # Проверим, что файловый обработчик НЕ добавлен сразу
            file_handlers_before = [
                h for h in root_logger.handlers if isinstance(h, logging.FileHandler)]
            assert len(file_handlers_before) == 0

            # Уровень конкретного логгера, полученного через get_logger, обычно NOTSET (0)
            # если явно не установлен. Это нормально.
            # assert logger.level == logging.DEBUG # Это неверная проверка

        finally:
            # Восстановим оригинальные обработчики после теста
            for handler in root_logger.handlers[:]:
                root_logger.removeHandler(handler)
            for handler in original_handlers:
                root_logger.addHandler(handler)

    def test_setup_file_handler_creates_log_file(self):
        """Тест: setup_file_handler создает файл лога в правильной директории."""
        lm = LogManager()
        base_filename = "test_app_run"

        # 1. Действие
        lm.setup_file_handler(base_filename)

        # 2. Утверждение
        # Проверяем, что файловый обработчик добавлен
        root_logger = logging.getLogger()
        file_handlers = [h for h in root_logger.handlers if isinstance(
            h, logging.FileHandler)]
        assert len(file_handlers) == 1

        log_file_path_str = file_handlers[0].baseFilename
        log_path = Path(log_file_path_str)

        # Проверяем, что файл находится в папке, указанной в config (которая патчена во временную)
        # Теперь это должно работать, так как мы патчим logging_config.LOG_DIR на self.test_dir
        # --- Исправлено: assert теперь проверяет, что родительская папка лога - это self.test_dir ---
        assert log_path.parent == Path(self.test_dir).resolve()
        # Проверяем, что имя файла содержит base_filename и дату (простая проверка)
        assert base_filename in log_path.name
        assert log_path.suffix == '.log'
        # Проверяем, что файл был создан
        assert log_path.exists()

    def test_setup_file_handler_removes_old_file_handler(self):
        """Тест: setup_file_handler удаляет предыдущий файловый обработчик."""
        lm = LogManager()
        base_filename_1 = "first_run"
        base_filename_2 = "second_run"

        # 1. Первый вызов
        lm.setup_file_handler(base_filename_1)
        root_logger = logging.getLogger()
        file_handlers_1 = [
            h for h in root_logger.handlers if isinstance(h, logging.FileHandler)]
        assert len(file_handlers_1) == 1
        first_log_path = Path(file_handlers_1[0].baseFilename)
        assert first_log_path.exists()

        # Запишем что-нибудь в первый лог
        logging.info("Message to first log")

        # 2. Второй вызов
        lm.setup_file_handler(base_filename_2)
        file_handlers_2 = [
            h for h in root_logger.handlers if isinstance(h, logging.FileHandler)]
        # Проверка: должен остаться только один файловый обработчик
        assert len(file_handlers_2) == 1
        second_log_path = Path(file_handlers_2[0].baseFilename)
        assert second_log_path.exists()
        assert second_log_path != first_log_path  # Пути должны быть разными

        # Проверим, что первый файл больше не используется (простая проверка: он не удален, но новый создан)
        # Более точная проверка потребовала бы мокирования или сложной проверки ссылок на файл,
        # что выходит за рамки простого теста. Но мы проверили, что новый создан и пути разные.

    def test_get_logger_returns_logger(self):
        """Тест: get_logger возвращает экземпляр logging.Logger."""
        lm = LogManager()
        logger_name = "MyTestLogger"

        # 1. Действие
        logger = lm.get_logger(logger_name)

        # 2. Утверждение
        assert isinstance(logger, logging.Logger)
        assert logger.name == logger_name

    def test_logging_works_after_setup(self, caplog):
        """Тест: Логирование работает корректно после настройки файлового логгера.
        Использует встроенную фикстуру pytest 'caplog' для захвата сообщений консоли."""
        lm = LogManager()
        base_filename = "test_logging"
        test_message = "This is a test INFO message"
        test_error_message = "This is a test ERROR message"

        # 1. Действие
        lm.setup_file_handler(base_filename)
        logger = lm.get_logger("test_logger")

        # 2. Утверждение
        # Проверка через caplog (захват сообщений, которые идут в консоль/другие обработчики)
        # caplog нуждается в установке правильного уровня
        with caplog.at_level(logging.INFO):
            logger.info(test_message)
            logger.error(test_error_message)

        assert test_message in caplog.text
        assert test_error_message in caplog.text
        # Проверка уровня
        assert "INFO" in caplog.text
        assert "ERROR" in caplog.text
        # Проверка, что сообщения попали в файл
        root_logger = logging.getLogger()
        file_handlers = [h for h in root_logger.handlers if isinstance(
            h, logging.FileHandler)]
        assert len(file_handlers) == 1
        log_file_path = Path(file_handlers[0].baseFilename)

        # Прочитаем файл и проверим наличие сообщений
        # (ждем немного, если запись асинхронная, хотя обычно она синхронная)
        # Добавим небольшую задержку или попробуем открыть файл несколько раз, если он занят
        import time
        time.sleep(0.1)  # Небольшая задержка

        with open(log_file_path, 'r', encoding='utf-8') as f:
            log_content = f.read()

        assert test_message in log_content
        assert test_error_message in log_content
        # Проверим формат (примерно)
        # Формат из кода: "%(asctime)s [%(levelname)s] %(message)s"
        # Проверим, что есть запись с правильным форматом для info
        # Это сложная проверка из-за даты, делаем приблизительно
        assert "[INFO]" in log_content
        assert "[ERROR]" in log_content
        assert test_message in log_content
        assert test_error_message in log_content

    def test_pyinstaller_path_handling_frozen(self):
        """Тест: Правильная обработка пути к папке логов при запуске через PyInstaller (имитация)."""
        # Этот тест имитирует запуск из .exe, созданным PyInstaller
        # 1. Настройка
        lm = LogManager()
        base_filename = "pyi_test_frozen"

        # Создаем временную "папку исполняемого файла" PyInstaller
        # Это будет директория, где находится .exe
        pyi_exe_dir = tempfile.mkdtemp(dir=self.test_dir)

        # --- Исправлено: Используем относительный путь специально для этого теста ---
        # Переопределяем патч LOG_DIR на относительный путь внутри этого теста
        relative_log_dir_for_pyi_test = "logs_for_pyi_test"
        with patch('logging_config.LOG_DIR', relative_log_dir_for_pyi_test):
            # 2. Патчим sys.frozen и sys.executable
            with patch.object(sys, 'frozen', True, create=True):
                with patch.object(sys, 'executable', os.path.join(pyi_exe_dir, "app.exe")):
                    # 3. Действие
                    lm.setup_file_handler(base_filename)

        # 4. Утверждение
        root_logger = logging.getLogger()
        file_handlers = [h for h in root_logger.handlers if isinstance(
            h, logging.FileHandler)]
        assert len(
            file_handlers) == 1, "Должен быть добавлен один файловый обработчик."

        log_file_path_str = file_handlers[0].baseFilename
        log_path = Path(log_file_path_str)

        # --- Исправлено: Расчет ожидаемого пути теперь корректный ---
        # При sys.frozen == True, log_dir_full = os.path.join(os.path.dirname(sys.executable), LOG_DIR)
        # os.path.dirname(sys.executable) = pyi_exe_dir
        # LOG_DIR = relative_log_dir_for_pyi_test (относительный путь)
        # os.path.join(pyi_exe_dir, relative_log_dir_for_pyi_test) = pyi_exe_dir / relative_log_dir_for_pyi_test
        expected_log_dir_path = Path(
            pyi_exe_dir) / relative_log_dir_for_pyi_test

        # Проверяем, что файл лога создан в папке рядом с "исполняемым файлом"
        # согласно логике из logging_config.py при sys.frozen == True
        assert log_path.parent == expected_log_dir_path.resolve(), (
            f"Ожидаемый путь к папке логов: {expected_log_dir_path.resolve()}\n"
            f"Фактический путь к папке логов: {log_path.parent}\n"
            f"pyi_exe_dir: {pyi_exe_dir}\n"
            # Исправлено имя переменной в сообщении
            f"relative_log_dir_for_pyi_test: {relative_log_dir_for_pyi_test}"
        )
        assert log_path.exists(
        ), f"Файл лога должен существовать по пути: {log_path}"

        # Очищаем
        shutil.rmtree(pyi_exe_dir, ignore_errors=True)

    def test_file_handler_formatter(self):
        """Тест: Файловый обработчик использует правильный форматтер."""
        lm = LogManager()
        base_filename = "formatter_test"
        # 1. Действие
        lm.setup_file_handler(base_filename)
        # 2. Утверждение
        root_logger = logging.getLogger()  # Используем глобальный импорт logging
        file_handlers = [h for h in root_logger.handlers if isinstance(
            h, logging.FileHandler)]
        assert len(file_handlers) == 1
        file_handler = file_handlers[0]
        formatter = file_handler.formatter
        assert formatter is not None

        # Проверим формат строки (примерно)
        # Ожидаемый формат из кода: "%(asctime)s [%(levelname)s] %(message)s"
        # Убираем локальный import logging
        record = logging.LogRecord(  # Используем глобальный импорт logging
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="Test message", args=(), exc_info=None
        )
        formatted_message = formatter.format(record)
        # Проверим, что в отформатированном сообщении есть части формата
        # Учитываем, что формат даты может немного отличаться
        assert "[INFO]" in formatted_message
        assert "Test message" in formatted_message
        # Проверим, что есть дата/время (наличие двоеточий и точек может указывать на это)
        assert ":" in formatted_message  # Очень приблизительно
