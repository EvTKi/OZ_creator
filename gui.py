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
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QGridLayout,
    QPushButton, QLineEdit, QTextEdit, QFileDialog, QLabel, QMessageBox, QProgressBar,
    QComboBox, QFrame, QSplitter, QSizePolicy
)
from PyQt5.QtCore import QThread, pyqtSignal, QObject, Qt
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon

# Импорты из вашего проекта
from config import DEBUG_FILE_NAME, DEBUG_PARENT_UID, LOG_DIR, SHEET_CATEGORIES, SHEET_TEMPLATES
from logging_config import LogManager
from xlsx_parser import ExcelParser
from cim_xml_creator import CIMXMLGenerator

# --- 1. Кастомный лог-хендлер для перенаправления логов в GUI ---


class QtHandler(QObject, logging.Handler):
    new_record = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.setLevel(logging.DEBUG)

    def emit(self, record):
        msg = self.format(record)
        self.new_record.emit(msg)

# --- 2. Рабочий поток для выполнения основной логики ---


class WorkerThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, excel_file_path, parent_uid, sheet_categories_name, sheet_templates_name):
        super().__init__()
        self.excel_file_path = excel_file_path
        self.parent_uid = parent_uid
        self.sheet_categories_name = sheet_categories_name
        self.sheet_templates_name = sheet_templates_name

    def run(self):
        try:
            logger_manager = LogManager(self.excel_file_path)
            logger = logger_manager.get_logger("WorkerThread")
            logger.info("Начало обработки...")

            self.progress.emit("Начало парсинга Excel-файла...")
            parser = ExcelParser(
                self.excel_file_path, logger_manager,
                override_sheet_categories=self.sheet_categories_name,
                override_sheet_templates=self.sheet_templates_name
            )
            structure = parser.build_structure()
            logger.info("Парсинг Excel-файла завершен.")

            self.progress.emit("Начало генерации CIM/XML...")
            generator = CIMXMLGenerator(
                structure, self.parent_uid, logger_manager)
            xml_content = generator.create_xml()
            logger.info("Генерация CIM/XML завершена.")

            base = os.path.splitext(os.path.basename(self.excel_file_path))[0]
            out_xml_path = f'{base}.xml'
            if getattr(sys, 'frozen', False):
                out_xml_path = os.path.join(
                    os.path.dirname(sys.executable), out_xml_path)
            with open(out_xml_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            logger.info(f"XML успешно сохранён: {out_xml_path}")
            self.finished.emit(out_xml_path)
        except Exception as e:
            error_msg = f"Ошибка во время выполнения: {str(e)}"
            logger.error(error_msg, exc_info=True)
            self.error.emit(error_msg)

# --- 3. Главное окно приложения ---


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.log_handler = None
        self.categories_sheet_combo: Optional[QComboBox] = None
        self.templates_sheet_combo: Optional[QComboBox] = None
        self.init_ui()
        self.setup_logging()
        self.apply_styles()

    def init_ui(self):
        self.setWindowTitle("Генератор CIM/XML из Excel")
        self.setGeometry(100, 100, 1000, 700)
        self.setMinimumSize(800, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Заголовок
        header_label = QLabel("Генератор CIM/XML")
        header_font = QFont()
        header_font.setPointSize(18)
        header_font.setBold(True)
        header_label.setFont(header_font)
        header_label.setAlignment(Qt.AlignCenter)
        header_label.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        main_layout.addWidget(header_label)

        # Основной контейнер
        content_widget = QFrame()
        content_widget.setFrameStyle(QFrame.StyledPanel)
        content_widget.setStyleSheet(
            "QFrame { background-color: #ffffff; border-radius: 8px; }")
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.addWidget(content_widget)

        # --- Форма ввода ---
        input_group = QFrame()
        input_group.setStyleSheet(
            "QFrame { background-color: #f8f9fa; border-radius: 6px; }")
        input_layout = QVBoxLayout(input_group)
        input_layout.setContentsMargins(15, 15, 15, 15)

        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignRight)
        form_layout.setHorizontalSpacing(15)
        form_layout.setVerticalSpacing(12)

        # Поле для пути к Excel-файлу
        file_widget = QWidget()
        file_layout = QHBoxLayout(file_widget)
        file_layout.setContentsMargins(0, 0, 0, 0)
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Выберите Excel-файл...")
        self.file_path_edit.setText(str(Path(DEBUG_FILE_NAME).resolve()) if Path(
            DEBUG_FILE_NAME).exists() else DEBUG_FILE_NAME)
        self.file_path_edit.setStyleSheet(
            "padding: 8px; border: 1px solid #ced4da; border-radius: 4px;")
        self.browse_button = QPushButton("Обзор")
        self.browse_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
        """)
        self.browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.browse_button)
        form_layout.addRow(QLabel("Файл Excel:"), file_widget)

        # Выбор листов
        sheets_layout = QHBoxLayout()
        sheets_layout.setSpacing(10)

        self.categories_sheet_combo = QComboBox()
        self.categories_sheet_combo.setEditable(True)
        self.categories_sheet_combo.addItem(SHEET_CATEGORIES)
        self.categories_sheet_combo.setStyleSheet(
            "padding: 6px; border: 1px solid #ced4da; border-radius: 4px;")

        self.templates_sheet_combo = QComboBox()
        self.templates_sheet_combo.setEditable(True)
        self.templates_sheet_combo.addItem(SHEET_TEMPLATES)
        self.templates_sheet_combo.setStyleSheet(
            "padding: 6px; border: 1px solid #ced4da; border-radius: 4px;")

        sheets_layout.addWidget(QLabel("Категории:"))
        sheets_layout.addWidget(self.categories_sheet_combo)
        sheets_layout.addWidget(QLabel("Шаблоны:"))
        sheets_layout.addWidget(self.templates_sheet_combo)
        form_layout.addRow(QLabel("Листы Excel:"), sheets_layout)

        # Поле для Parent UID
        self.uid_edit = QLineEdit()
        self.uid_edit.setPlaceholderText("Введите Parent UID (опционально)")
        self.uid_edit.setText(DEBUG_PARENT_UID)
        self.uid_edit.setStyleSheet(
            "padding: 8px; border: 1px solid #ced4da; border-radius: 4px;")
        form_layout.addRow(QLabel("Parent UID:"), self.uid_edit)

        input_layout.addLayout(form_layout)
        content_layout.addWidget(input_group)

        # --- Кнопка запуска ---
        self.run_button = QPushButton("Запустить генерацию")
        self.run_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                padding: 12px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.run_button.clicked.connect(self.start_processing)
        content_layout.addWidget(self.run_button)

        # --- Прогресс-бар и логи ---
        bottom_widget = QFrame()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ced4da;
                border-radius: 4px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                width: 10px;
            }
        """)
        bottom_layout.addWidget(self.progress_bar)

        # Область логов
        log_group = QFrame()
        log_group.setStyleSheet(
            "QFrame { background-color: #f8f9fa; border-radius: 6px; }")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(10, 10, 10, 10)

        log_header = QLabel("Лог выполнения")
        log_header.setStyleSheet(
            "font-weight: bold; color: #2c3e50; margin-bottom: 5px;")
        log_layout.addWidget(log_header)

        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.log_text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-family: 'Courier New', monospace;
                font-size: 9pt;
            }
        """)
        log_layout.addWidget(self.log_text_edit)
        bottom_layout.addWidget(log_group)

        content_layout.addWidget(bottom_widget)

        # --- Кнопки действий ---
        actions_layout = QHBoxLayout()
        actions_layout.addStretch()

        self.open_xml_button = QPushButton("Открыть XML")
        self.open_xml_button.setEnabled(False)
        self.open_xml_button.clicked.connect(self.open_xml_file)
        self.open_xml_button.setStyleSheet("""
            QPushButton {
                background-color: #34495e;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2c3e50;
            }
            QPushButton:pressed {
                background-color: #212f3c;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)

        self.open_log_button = QPushButton("Папка логов")
        self.open_log_button.clicked.connect(self.open_log_folder)
        self.open_log_button.setStyleSheet("""
            QPushButton {
                background-color: #34495e;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2c3e50;
            }
            QPushButton:pressed {
                background-color: #212f3c;
            }
        """)

        actions_layout.addWidget(self.open_xml_button)
        actions_layout.addWidget(self.open_log_button)
        content_layout.addLayout(actions_layout)

        self.output_xml_path = None

    def apply_styles(self):
        # Применение глобальных стилей
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ecf0f1;
            }
            QLabel {
                color: #2c3e50;
            }
            QFormLayout QLabel {
                font-weight: bold;
            }
        """)

    def setup_logging(self):
        self.log_handler = QtHandler()
        formatter = logging.Formatter('[%(levelname)s] %(message)s')
        self.log_handler.setFormatter(formatter)
        self.log_handler.new_record.connect(self.append_log)
        root_logger = logging.getLogger()
        root_logger.addHandler(self.log_handler)
        root_logger.setLevel(logging.INFO)

    def append_log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text_edit.append(formatted_message)
        self.log_text_edit.moveCursor(self.log_text_edit.textCursor().End)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel-файл", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)
            self.load_sheet_names(file_path)

    def load_sheet_names(self, file_path):
        try:
            if self.categories_sheet_combo is None or self.templates_sheet_combo is None:
                self.append_log("Ошибка: ComboBox не инициализированы.")
                return

            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names

            if sheet_names:
                sheet_names_str = [str(name) for name in sheet_names]

                current_cat_text = self.categories_sheet_combo.currentText()
                current_tmpl_text = self.templates_sheet_combo.currentText()

                self.categories_sheet_combo.clear()
                self.templates_sheet_combo.clear()
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
        excel_file = self.file_path_edit.text().strip()
        parent_uid = self.uid_edit.text().strip()

        if self.categories_sheet_combo is None or self.templates_sheet_combo is None:
            QMessageBox.critical(
                self, "Ошибка", "ComboBox не инициализированы.")
            return

        sheet_categories_name = self.categories_sheet_combo.currentText().strip()
        sheet_templates_name = self.templates_sheet_combo.currentText().strip()

        if not excel_file:
            QMessageBox.warning(self, "Ошибка ввода",
                                "Пожалуйста, выберите Excel-файл.")
            return
        if not os.path.exists(excel_file):
            QMessageBox.critical(self, "Ошибка файла",
                                 f"Файл не найден: {excel_file}")
            return
        if not sheet_categories_name:
            QMessageBox.warning(self, "Ошибка ввода",
                                "Пожалуйста, введите или выберите имя листа категорий.")
            return
        if not sheet_templates_name:
            QMessageBox.warning(self, "Ошибка ввода",
                                "Пожалуйста, введите или выберите имя листа шаблонов.")
            return

        self.run_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.log_text_edit.clear()
        self.open_xml_button.setEnabled(False)
        self.output_xml_path = None

        self.worker_thread = WorkerThread(
            excel_file, parent_uid if parent_uid else None,
            sheet_categories_name, sheet_templates_name
        )

        self.worker_thread.finished.connect(self.on_worker_finished)
        self.worker_thread.error.connect(self.on_worker_error)
        self.worker_thread.progress.connect(self.append_log)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.error.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def on_worker_finished(self, xml_path):
        self.progress_bar.setVisible(False)
        self.run_button.setEnabled(True)
        self.output_xml_path = xml_path
        self.open_xml_button.setEnabled(True)
        self.append_log("Обработка завершена успешно!")
        QMessageBox.information(
            self, "Готово", f"XML-файл успешно создан:\n{xml_path}")

    def on_worker_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.run_button.setEnabled(True)
        self.append_log(f"Ошибка: {error_message}")
        QMessageBox.critical(
            self, "Ошибка", f"Произошла ошибка:\n{error_message}")

    def open_xml_file(self):
        if self.output_xml_path and os.path.exists(self.output_xml_path):
            try:
                if sys.platform == "win32":
                    os.startfile(self.output_xml_path)
                elif sys.platform == "darwin":
                    os.system(f"open '{self.output_xml_path}'")
                else:
                    os.system(f"xdg-open '{self.output_xml_path}'")
            except Exception as e:
                QMessageBox.critical(
                    self, "Ошибка", f"Не удалось открыть файл:\n{e}")
        else:
            QMessageBox.warning(
                self, "Ошибка", "Файл XML не найден или еще не создан.")

    def open_log_folder(self):
        try:
            log_path = LOG_DIR
            if getattr(sys, 'frozen', False):
                log_path = os.path.join(
                    os.path.dirname(sys.executable), LOG_DIR)
            if not os.path.exists(log_path):
                os.makedirs(log_path)
            if sys.platform == "win32":
                os.startfile(log_path)
            elif sys.platform == "darwin":
                os.system(f"open '{log_path}'")
            else:
                os.system(f"xdg-open '{log_path}'")
        except Exception as e:
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось открыть папку логов:\n{e}")

    def closeEvent(self, event):
        import logging
        if self.worker_thread is not None:
            try:
                is_running = self.worker_thread.isRunning()
            except RuntimeError:
                is_running = False
                self.worker_thread = None
            if is_running and self.worker_thread is not None:
                self.append_log(
                    "Приложение закрывается, ожидание завершения потока...")
                worker = self.worker_thread
                if worker is not None:
                    worker.quit()
                    if not worker.wait(2000):
                        self.append_log("Поток не завершился вовремя.")
                self.worker_thread = None
            elif self.worker_thread is not None:
                self.worker_thread = None
        if self.log_handler:
            root_logger = logging.getLogger()
            try:
                root_logger.removeHandler(self.log_handler)
                self.append_log(
                    "Лог-хендлер успешно отсоединен от корневого логгера.")
            except (ValueError, RuntimeError):
                pass
            try:
                if hasattr(self.log_handler, 'new_record'):
                    self.log_handler.new_record.disconnect()
            except (RuntimeError, TypeError):
                pass
            self.log_handler = None
        event.accept()


# --- 4. Точка входа ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
