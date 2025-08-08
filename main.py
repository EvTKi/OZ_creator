# main.py
"""
Точка входа в приложение.
"""

# Импортируем константы из нового config.py
from modules.config import DEBUG_FILE_NAME, DEBUG_PARENT_UID
from modules.logging_config import LogManager
from modules.xlsx_parser import ExcelParser
from modules.cim_xml_creator import CIMXMLGenerator
import os


def main():
    """Главная функция приложения."""
    # Инициализация логирования
    logger_manager = LogManager(DEBUG_FILE_NAME)
    logger = logger_manager.get_logger(__name__)

    try:
        # --- 1. Парсинг Excel ---
        logger.info("Начало парсинга Excel-файла...")
        parser = ExcelParser(DEBUG_FILE_NAME, logger_manager)
        structure = parser.build_structure()
        logger.info("Парсинг Excel-файла завершен.")

        # --- 2. Генерация XML ---
        logger.info("Начало генерации CIM/XML...")
        generator = CIMXMLGenerator(
            structure, DEBUG_PARENT_UID, logger_manager)
        xml_content = generator.create_xml()
        logger.info("Генерация CIM/XML завершена.")

        # --- 3. Запись в файл ---
        base = os.path.splitext(os.path.basename(DEBUG_FILE_NAME))[0]
        out_xml_path = f'{base}.xml'
        with open(out_xml_path, 'w', encoding='utf-8') as f:
            f.write(xml_content)
        logger.info(f"XML успешно сохранён: {out_xml_path}")
        print(f"✅ XML сформирован и сохранён в {out_xml_path}")

    except Exception as e:
        logger.error(f"Критическая ошибка в main: {e}")
        print(f"[ERROR] {e}")
        # Можно добавить sys.exit(1) для выхода с кодом ошибки


if __name__ == "__main__":
    main()
