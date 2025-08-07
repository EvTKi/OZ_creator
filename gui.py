# gui.py
"""
GUI-приложение для генератора CIM/XML из Excel.
Использует PyQt5 для создания интерфейса.
"""
import sys
import os
import logging
from datetime import datetime
from pathlib import Path
import pandas as pd  # Для получения списка листов

# Импорты PyQt5
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLineEdit, QTextEdit, QFileDialog, QLabel, QMessageBox, QProgressBar,
    QComboBox  # <-- Добавлено
)
from PyQt5.QtCore import QThread, pyqtSignal, QObject, Qt

# Импорты из вашего проекта
# Предполагается, что config.py, logging_config.py, xlsx_parser.py, cim_xml_creator.py находятся рядом
# <-- Добавлен импорт констант
from config import DEBUG_FILE_NAME, DEBUG_PARENT_UID, LOG_DIR, SHEET_CATEGORIES, SHEET_TEMPLATES
from logging_config import LogManager
from xlsx_parser import ExcelParser
from cim_xml_creator import CIMXMLGenerator

# --- 1. Кастомный лог-хендлер для перенаправления логов в GUI ---


class QtHandler(QObject, logging.Handler):
    """
    Пользовательский обработчик логов, который отправляет сообщения в Qt сигнал.
    Это позволяет обновлять GUI из потока логгера.
    """
    new_record = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        # Устанавливаем уровень, который будет обрабатываться
        self.setLevel(logging.DEBUG)

    def emit(self, record):
        """Отправляет запись лога через сигнал."""
        msg = self.format(record)
        self.new_record.emit(msg)

# --- 2. Рабочий поток для выполнения основной логики ---


class WorkerThread(QThread):
    """
    Поток для выполнения длительных операций (парсинг, генерация XML),
    чтобы не блокировать GUI.
    """
    # Сигналы для общения с основным потоком GUI
    finished = pyqtSignal(str)  # Успешное завершение, передаем путь к XML
    error = pyqtSignal(str)     # Ошибка, передаем сообщение об ошибке
    progress = pyqtSignal(str)  # Прогресс/информация, передаем сообщение

    # --- Изменения в __init__ ---
    # Добавлены параметры sheet_categories_name и sheet_templates_name
    def __init__(self, excel_file_path, parent_uid, sheet_categories_name, sheet_templates_name):
        super().__init__()
        self.excel_file_path = excel_file_path
        self.parent_uid = parent_uid
        # --- Новые атрибуты ---
        self.sheet_categories_name = sheet_categories_name
        self.sheet_templates_name = sheet_templates_name
        # -----------------------

    def run(self):
        """
        Основной метод потока. Здесь выполняется вся логика приложения.
        """
        try:
            # 1. Настройка логирования для этого запуска
            # Используем LogManager из вашего проекта
            logger_manager = LogManager(self.excel_file_path)
            logger = logger_manager.get_logger("WorkerThread")
            logger.info("Начало обработки...")

            # --- Изменения в передаче параметров ---
            # Передаем имена листов в ExcelParser
            self.progress.emit("Начало парсинга Excel-файла...")
            # parser = ExcelParser(self.excel_file_path, logger_manager) # Было
            parser = ExcelParser(
                self.excel_file_path, logger_manager,
                override_sheet_categories=self.sheet_categories_name,  # <-- Новое
                override_sheet_templates=self.sheet_templates_name    # <-- Новое
            )  # Стало
            structure = parser.build_structure()
            logger.info("Парсинг Excel-файла завершен.")
            # -------------------------------------

            # 3. Генерация XML
            self.progress.emit("Начало генерации CIM/XML...")
            generator = CIMXMLGenerator(
                structure, self.parent_uid, logger_manager)
            xml_content = generator.create_xml()
            logger.info("Генерация CIM/XML завершена.")

            # 4. Запись в файл
            base = os.path.splitext(os.path.basename(self.excel_file_path))[0]
            out_xml_path = f'{base}.xml'
            # Если .exe, сохраняем рядом с ним, иначе в текущей директории
            if getattr(sys, 'frozen', False):
                out_xml_path = os.path.join(
                    os.path.dirname(sys.executable), out_xml_path)
            with open(out_xml_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            logger.info(f"XML успешно сохранён: {out_xml_path}")
            # Отправляем сигнал об успешном завершении с путем к файлу
            self.finished.emit(out_xml_path)
        except Exception as e:
            # Отправляем сигнал об ошибке
            error_msg = f"Ошибка во время выполнения: {str(e)}"
            # exc_info=True для полного трейса в лог
            logger.error(error_msg, exc_info=True)
            self.error.emit(error_msg)

# --- 3. Главное окно приложения ---


class MainWindow(QMainWindow):
    """
    Главное окно GUI-приложения.
    """

    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.log_handler = None
        # --- Явная инициализация атрибутов ComboBox ---
        # Это устраняет предупреждения Pylance о доступе к атрибутам None
        self.categories_sheet_combo: Optional[QComboBox] = None
        self.templates_sheet_combo: Optional[QComboBox] = None
        # -----------------------------------------------
        self.init_ui()
        self.setup_logging()

    def init_ui(self):
        """Инициализация пользовательского интерфейса."""
        self.setWindowTitle("Генератор CIM/XML из Excel")
        # Увеличил высоту для новых элементов
        self.setGeometry(100, 100, 800, 650)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- Форма ввода ---
        form_layout = QFormLayout()

        # Поле для пути к Excel-файлу
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Путь к Excel-файлу...")
        self.file_path_edit.setText(str(Path(DEBUG_FILE_NAME).resolve()) if Path(
            DEBUG_FILE_NAME).exists() else DEBUG_FILE_NAME)
        self.browse_button = QPushButton("Обзор...")
        self.browse_button.clicked.connect(self.browse_file)
        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_button)
        form_layout.addRow(QLabel("Файл Excel:"), file_layout)

        # --- Инициализация и добавление ComboBox ---
        # Теперь мы уверены, что атрибуты существуют
        self.categories_sheet_combo = QComboBox()  # <-- Создание объекта
        self.categories_sheet_combo.setEditable(True)
        self.categories_sheet_combo.addItem(SHEET_CATEGORIES)
        form_layout.addRow(QLabel("Лист категорий:"),
                           self.categories_sheet_combo)

        self.templates_sheet_combo = QComboBox()  # <-- Создание объекта
        self.templates_sheet_combo.setEditable(True)
        self.templates_sheet_combo.addItem(SHEET_TEMPLATES)
        form_layout.addRow(QLabel("Лист шаблонов:"),
                           self.templates_sheet_combo)
        # ------------------------------------------

        # Поле для Parent UID
        self.uid_edit = QLineEdit()
        self.uid_edit.setPlaceholderText("Введите Parent UID...")
        self.uid_edit.setText(DEBUG_PARENT_UID)
        form_layout.addRow(QLabel("Parent UID:"), self.uid_edit)

        main_layout.addLayout(form_layout)

        # --- Кнопка запуска ---
        self.run_button = QPushButton("Запустить")
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)

        # --- Прогресс-бар ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # --- Область логов ---
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        font = self.log_text_edit.font()
        font.setFamily("Courier New")
        font.setPointSize(9)
        self.log_text_edit.setFont(font)
        main_layout.addWidget(QLabel("Лог выполнения:"))
        main_layout.addWidget(self.log_text_edit)

        # --- Кнопки действий после завершения ---
        self.actions_layout = QHBoxLayout()
        self.open_xml_button = QPushButton("Открыть XML")
        self.open_xml_button.clicked.connect(self.open_xml_file)
        self.open_xml_button.setEnabled(False)
        self.open_log_button = QPushButton("Открыть папку логов")
        self.open_log_button.clicked.connect(self.open_log_folder)
        self.actions_layout.addWidget(self.open_xml_button)
        self.actions_layout.addWidget(self.open_log_button)
        self.actions_layout.addStretch()
        main_layout.addLayout(self.actions_layout)

        self.output_xml_path = None

    def setup_logging(self):
        """Настройка системы логирования для перенаправления в GUI."""
        # Создаем кастомный хендлер
        self.log_handler = QtHandler()
        # Устанавливаем формат логов (без даты/времени, так как оно будет в текстовом поле)
        formatter = logging.Formatter('[%(levelname)s] %(message)s')
        self.log_handler.setFormatter(formatter)
        # Подключаем сигнал хендлера к слоту обновления текста
        self.log_handler.new_record.connect(self.append_log)
        # Добавляем хендлер к корневому логгеру
        root_logger = logging.getLogger()
        root_logger.addHandler(self.log_handler)
        root_logger.setLevel(logging.INFO)  # Устанавливаем уровень логирования

    def append_log(self, message):
        """
        Слот для добавления сообщения в текстовое поле логов.
        Вызывается при получении сигнала от QtHandler.
        """
        # Добавляем временную метку
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text_edit.append(formatted_message)
        # Прокручиваем вниз
        self.log_text_edit.moveCursor(self.log_text_edit.textCursor().End)

    def browse_file(self):
        """Открывает диалог выбора файла."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)

            # --- Новое: Загрузка списка листов при выборе файла ---
            self.load_sheet_names(file_path)
            # ----------------------------------------------------

    # --- Новый метод для загрузки имен листов ---
    def load_sheet_names(self, file_path):
        """
        Загружает список имен листов из Excel-файла и заполняет ComboBox.
        """
        try:
            # Проверка на None (для Pylance и безопасности)
            if self.categories_sheet_combo is None or self.templates_sheet_combo is None:
                self.append_log("Ошибка: ComboBox не инициализированы.")
                return

            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names

            if sheet_names:
                # Преобразуем список в список строк, чтобы избежать ошибок типов
                sheet_names_str = [str(name) for name in sheet_names]

                current_cat_text = self.categories_sheet_combo.currentText()
                current_tmpl_text = self.templates_sheet_combo.currentText()

                # Очищаем и заполняем ComboBox
                self.categories_sheet_combo.clear()
                self.templates_sheet_combo.clear()
                # Передаем список строк
                self.categories_sheet_combo.addItems(sheet_names_str)
                self.templates_sheet_combo.addItems(sheet_names_str)

                if current_cat_text in sheet_names_str:
                    self.categories_sheet_combo.setCurrentText(
                        current_cat_text)
                elif SHEET_CATEGORIES in sheet_names_str:
                    self.categories_sheet_combo.setCurrentText(
                        SHEET_CATEGORIES)
                else:
                    self.categories_sheet_combo.setCurrentIndex(0)

                if current_tmpl_text in sheet_names_str:
                    self.templates_sheet_combo.setCurrentText(
                        current_tmpl_text)
                elif SHEET_TEMPLATES in sheet_names_str:
                    self.templates_sheet_combo.setCurrentText(SHEET_TEMPLATES)
                else:
                    self.templates_sheet_combo.setCurrentIndex(0)

                self.append_log(f"Список листов загружен: {sheet_names}")
            else:
                self.append_log("Предупреждение: В файле не найдено листов.")
        except Exception as e:
            error_msg = f"Ошибка при загрузке списка листов: {e}"
            self.append_log(error_msg)

    def start_processing(self):
        """Запускает процесс обработки в отдельном потоке."""
        excel_file = self.file_path_edit.text().strip()
        parent_uid = self.uid_edit.text().strip()

        # Проверка на None (для Pylance и безопасности)
        if self.categories_sheet_combo is None or self.templates_sheet_combo is None:
            QMessageBox.critical(
                self, "Ошибка", "ComboBox не инициализированы.")
            return

        # --- Получение значений из ComboBox ---
        sheet_categories_name = self.categories_sheet_combo.currentText().strip()
        sheet_templates_name = self.templates_sheet_combo.currentText().strip()
        # -------------------------------------
        # Можно также показать QMessageBox.warning(self, "Ошибка", error_msg)

        # if not parent_uid: # UID может быть опциональным, если используется дефолтный
        #     QMessageBox.warning(self, "Ошибка ввода", "Пожалуйста, введите Parent UID.")
        #     return

        # Подготавливаем GUI
        self.run_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.log_text_edit.clear()
        self.open_xml_button.setEnabled(False)
        self.output_xml_path = None

        # --- Передача имен листов в WorkerThread ---
        self.worker_thread = WorkerThread(
            excel_file, parent_uid if parent_uid else None,
            sheet_categories_name, sheet_templates_name  # <-- Новые аргументы
        )
        # -------------------------------------------

        self.worker_thread.finished.connect(self.on_worker_finished)
        self.worker_thread.error.connect(self.on_worker_error)
        # Прямая передача сообщений прогресса в лог
        self.worker_thread.progress.connect(self.append_log)
        self.worker_thread.finished.connect(
            self.worker_thread.deleteLater)  # Авто-удаление потока
        self.worker_thread.error.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def on_worker_finished(self, xml_path):
        """Слот, вызываемый при успешном завершении рабочего потока."""
        self.progress_bar.setVisible(False)
        self.run_button.setEnabled(True)
        self.output_xml_path = xml_path
        self.open_xml_button.setEnabled(True)
        self.append_log("Обработка завершена успешно!")
        QMessageBox.information(
            self, "Готово", f"XML-файл успешно создан:\n{xml_path}")

    def on_worker_error(self, error_message):
        """Слот, вызываемый при ошибке в рабочем потоке."""
        self.progress_bar.setVisible(False)
        self.run_button.setEnabled(True)
        self.append_log(f"Ошибка: {error_message}")
        QMessageBox.critical(
            self, "Ошибка", f"Произошла ошибка:\n{error_message}")

    def open_xml_file(self):
        """Открывает созданный XML-файл."""
        if self.output_xml_path and os.path.exists(self.output_xml_path):
            try:
                # Открывает файл с помощью приложения по умолчанию
                if sys.platform == "win32":
                    os.startfile(self.output_xml_path)
                elif sys.platform == "darwin":  # macOS
                    os.system(f"open '{self.output_xml_path}'")
                else:  # linux
                    os.system(f"xdg-open '{self.output_xml_path}'")
            except Exception as e:
                QMessageBox.critical(
                    self, "Ошибка", f"Не удалось открыть файл:\n{e}")
        else:
            QMessageBox.warning(
                self, "Ошибка", "Файл XML не найден или еще не создан.")

    def open_log_folder(self):
        """Открывает папку с логами."""
        try:
            log_path = LOG_DIR
            # Если .exe, логи могут быть рядом с ним
            if getattr(sys, 'frozen', False):
                log_path = os.path.join(
                    os.path.dirname(sys.executable), LOG_DIR)
            if not os.path.exists(log_path):
                os.makedirs(log_path)  # Создаем, если не существует
            if sys.platform == "win32":
                os.startfile(log_path)
            elif sys.platform == "darwin":  # macOS
                os.system(f"open '{log_path}'")
            else:  # linux
                os.system(f"xdg-open '{log_path}'")
        except Exception as e:
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось открыть папку логов:\n{e}")

    def closeEvent(self, event):
        """
        Переопределяем событие закрытия окна для корректного завершения потока
        и безопасной очистки лог-хендлера до завершения Qt.
        """
        import logging  # Убедимся, что logging импортирован
        # --- 1. Завершение рабочего потока (если он есть и запущен) ---
        if self.worker_thread is not None:
            try:
                # Проверим, запущен ли поток. Если объект удален, это вызовет RuntimeError
                is_running = self.worker_thread.isRunning()
            except RuntimeError:
                # Объект уже удален, ничего делать не нужно
                is_running = False
                # Явно устанавливаем ссылку в None после ошибки
                self.worker_thread = None
            # --- ВАЖНО: Повторная проверка перед использованием ---
            # Pylance может ругаться, если не видит прямой связи между проверкой и использованием
            # Явная проверка устраняет это.
            if is_running and self.worker_thread is not None:  # <-- Добавлена повторная проверка
                self.append_log(
                    "Приложение закрывается, ожидание завершения потока...")
                # Еще одна проверка непосредственно перед вызовом
                worker = self.worker_thread  # Создаем временную ссылку
                if worker is not None:  # <-- Убеждаемся, что worker не None
                    worker.quit()  # Просим поток завершиться
                    # Ждем завершения, но не бесконечно (например, 2 секунды)
                    # Используем QThread.wait(), а не просто time.sleep
                    # <-- worker гарантированно не None здесь
                    if not worker.wait(2000):
                        self.append_log("Поток не завершился вовремя.")
                # Очищаем ссылку
                self.worker_thread = None
            elif self.worker_thread is not None:
                # Поток существует, но не запущен, просто очищаем ссылку
                self.worker_thread = None
        # --- 2. КРИТИЧЕСКИ ВАЖНО: Безопасное и раннее удаление лог-хендлера ---
        # Делаем это ДО завершения работы Qt, чтобы избежать ошибки в logging.shutdown()
        if self.log_handler:
            root_logger = logging.getLogger()
            # ВАЖНО: Удаляем хендлер из логгера ПЕРВЫМ ДЕЛОМ
            try:
                # Пытаемся удалить хендлер из корневого логгера
                # Это должно произойти до того, как Qt уничтожит объект C++ self.log_handler
                root_logger.removeHandler(self.log_handler)
                # Для отладки
                self.append_log(
                    "Лог-хендлер успешно отсоединен от корневого логгера.")
            except (ValueError, RuntimeError) as e:
                # ValueError: если хендлера уже нет в списке (редко, но возможно)
                # RuntimeError: если объект Python/Qt уже удален (наш случай)
                # В любом случае, мы больше не хотим, чтобы этот хендлер был в списке.
                pass  # Игнорируем ошибки, так как наша цель - убрать его из списка
            # --- 3. Очистка сигналов хендлера (опционально, но хорошо) ---
            try:
                # Отсоединяем все сигналы хендлера, чтобы избежать ошибок при его удалении
                # Это предотвращает попытки Qt отправить сигнал уже уничтоженному слоту
                if hasattr(self.log_handler, 'new_record'):
                    self.log_handler.new_record.disconnect()  # Отсоединяем все подключения сигнала
            except (RuntimeError, TypeError):
                # RuntimeError: если объект Qt уже удален
                # TypeError: если сигнал не был подключен или объект удален
                pass  # Игнорируем ошибки
            # --- 4. Явное обнуление ссылки на хендлер ---
            # Это помогает Python понять, что объект больше не нужен
            self.log_handler = None
        # Принимаем событие закрытия, позволяя Qt завершить работу
        event.accept()


# --- 4. Точка входа ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Устанавливаем Fusion стиль для лучшего внешнего вида (опционально)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
