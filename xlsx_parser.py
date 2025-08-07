# xlsx_parser.py
"""
Модуль для парсинга структуры категорий из Excel-файла.
"""

import pandas as pd
from collections import defaultdict
import logging
from typing import List, Set, Dict, Any, Optional
import openpyxl
from openpyxl.styles.colors import Color
# Импортируем константы из нового config.py
from config import (
    RED_LIKE_COLORS, RED_INDEXED_COLORS, RED_THEME_INDICES,
    RED_THEME_TINT_THRESHOLD, DEFAULT_STORAGE_DEPTH,
    SHEET_CATEGORIES, SHEET_TEMPLATES,
    COL_CATEGORY_TYPE, COL_CATEGORY,
    COL_TEMPLATE_CATEGORY, COL_TEMPLATE_EXPRESSION
)
from logging_config import LogManager


class ExcelParser:
    """
    Класс для парсинга структуры категорий из Excel-файла.
    """

    def __init__(self, filename: str, logger_manager: LogManager):
        """
        Инициализирует парсер Excel.

        Args:
            filename (str): Путь к Excel-файлу.
            logger_manager (LogManager): Менеджер логирования.
        """
        self.filename = filename
        self.logger = logger_manager.get_logger(self.__class__.__name__)
        self._logger_manager = logger_manager

    def _is_red_color(self, color: Optional[Color]) -> bool:
        """
        Проверяет, является ли цвет (Font.color или Fill.fgColor) красным.

        Args:
            color (Optional[Color]): Цвет для проверки.

        Returns:
            bool: True, если цвет считается красным, иначе False.
        """
        if color is None:
            return False
        # 1. Проверка по RGB
        try:
            if hasattr(color, 'rgb') and color.rgb is not None:
                rgb_val = color.rgb
                rgb_str = None
                if isinstance(rgb_val, str):
                    rgb_str = rgb_val
                elif hasattr(rgb_val, 'rgb') and isinstance(rgb_val.rgb, str):
                    rgb_str = rgb_val.rgb
                elif str(rgb_val).startswith('RGB:'):
                    try:
                        rgb_hex = str(rgb_val).split(':', 1)[1].strip()
                        if len(rgb_hex) in (6, 8) and all(c in '0123456789ABCDEFabcdef' for c in rgb_hex):
                            rgb_str = rgb_hex
                    except Exception as e:
                        self.logger.debug(
                            f"Не удалось распарсить RGB из строки: {rgb_val}, ошибка: {e}")
                if rgb_str:
                    rgb_str = rgb_str.upper()
                    if len(rgb_str) == 8 and rgb_str.startswith('FF'):
                        full_rgb = rgb_str
                    elif len(rgb_str) == 8:
                        full_rgb = 'FF' + rgb_str[2:]
                    elif len(rgb_str) == 6:
                        full_rgb = 'FF' + rgb_str
                    else:
                        full_rgb = None
                    if full_rgb and full_rgb in RED_LIKE_COLORS:
                        return True
        except Exception as e:
            self.logger.debug(f"Ошибка при обработке RGB цвета {color}: {e}")

        # 2. Theme color
        try:
            if hasattr(color, 'theme') and color.theme is not None:
                if color.theme in RED_THEME_INDICES:
                    tint = getattr(color, 'tint', 0.0)
                    if isinstance(tint, (int, float)) and tint > RED_THEME_TINT_THRESHOLD:
                        return True
        except Exception as e:
            self.logger.debug(f"Ошибка при обработке theme цвета {color}: {e}")

        # 3. Indexed color
        try:
            if hasattr(color, 'indexed'):
                if color.indexed in RED_INDEXED_COLORS:
                    return True
        except Exception as e:
            self.logger.debug(
                f"Ошибка при обработке indexed цвета {color}: {e}")

        return False

    def _get_non_red_rows(self, sheet_name: str, header_row_idx: int, nrows: int) -> List[int]:
        """
        Возвращает индексы строк (относительно DataFrame), в которых отсутствует красный цвет 
        в заливке ячеек или цвете шрифта.

        Args:
            sheet_name (str): Имя листа.
            header_row_idx (int): Номер строки заголовка в Excel (индексация с 1).
            nrows (int): Количество строк для анализа.

        Returns:
            List[int]: Список индексов строк (в индексации pandas), без красного цвета.
        """
        self.logger.debug(f"Начало фильтрации строк по цвету для листа '{sheet_name}', "
                          f"строки {header_row_idx + 1}-{header_row_idx + nrows}")

        try:
            # data_only=True для получения значений, а не формул
            wb = openpyxl.load_workbook(self.filename, data_only=True)
            self.logger.debug(
                f"Рабочая книга '{self.filename}' загружена для анализа цвета.")
        except FileNotFoundError:
            self.logger.error(
                f"Файл '{self.filename}' не найден для анализа цвета.")
            raise
        except Exception as e:
            self.logger.error(
                f"Ошибка при загрузке файла '{self.filename}' для анализа цвета: {e}")
            raise

        try:
            ws = wb[sheet_name]
            self.logger.debug(f"Лист '{sheet_name}' открыт для анализа цвета.")
        except KeyError:
            self.logger.error(
                f"Лист '{sheet_name}' не найден в файле '{self.filename}'.")
            raise
        except Exception as e:
            self.logger.error(
                f"Ошибка при открытии листа '{sheet_name}' для анализа цвета: {e}")
            raise

        result = []
        red_count = 0

        for ex_idx in range(header_row_idx + 1, header_row_idx + 1 + nrows):
            is_red = False
            try:
                for cell in ws[ex_idx]:
                    # Проверка заливки
                    fill = cell.fill
                    fg_color = getattr(fill, 'fgColor', None)
                    if fg_color is not None and self._is_red_color(fg_color):
                        is_red = True
                        self.logger.debug(
                            f"Строка Excel {ex_idx} помечена как 'красная' из-за заливки ячейки {cell.coordinate}")
                        break
                    # Проверка шрифта
                    font_color = getattr(cell.font, 'color', None)
                    if font_color is not None and self._is_red_color(font_color):
                        is_red = True
                        self.logger.debug(
                            f"Строка Excel {ex_idx} помечена как 'красная' из-за цвета шрифта ячейки {cell.coordinate}")
                        break
            except Exception as e:
                self.logger.warning(
                    f"Ошибка при анализе строки {ex_idx} на листе '{sheet_name}': {e}")
                # Не прерываем весь процесс из-за ошибки в одной строке, продолжаем анализ

            if not is_red:
                df_index = ex_idx - (header_row_idx + 1)
                result.append(df_index)
            else:
                red_count += 1

        self.logger.info(f"Фильтрация по цвету завершена для листа '{sheet_name}'. "
                         f"Проанализировано: {nrows}, Пропущено (красных): {red_count}, Оставлено: {len(result)}")
        return result

    def _read_table_by_header_df(self, sheet_name: str, required_headers: Set[str]) -> pd.DataFrame:
        """
        Читает таблицу из Excel-листа, начиная со строки заголовка.

        Args:
            sheet_name (str): Имя листа Excel.
            required_headers (Set[str]): Множество обязательных заголовков столбцов.

        Returns:
            pd.DataFrame: DataFrame с данными таблицы, отфильтрованный по цвету.

        Raises:
            Exception: Если строка заголовка не найдена.
        """
        self.logger.info(f"Начало чтения таблицы с листа '{sheet_name}'")

        try:
            # 1. чтение листа
            df_whole = pd.read_excel(
                self.filename, sheet_name=sheet_name, header=None)
            self.logger.debug(
                f"Лист '{sheet_name}' прочитан в DataFrame. Размер: {df_whole.shape}")
        except Exception as e:
            self.logger.error(
                f"Ошибка при чтении листа '{sheet_name}' из файла '{self.filename}': {e}")
            raise

        # 2. поиск строки шапки
        self.logger.debug(
            f"Поиск строки заголовка с колонками {required_headers} на листе '{sheet_name}'")
        found = False
        header_row = None
        rows_searched = 0
        for idx, row in df_whole.iterrows():
            rows_searched += 1
            headers = [str(x).strip() for x in row if pd.notna(x)]
            if required_headers.issubset(set(headers)):
                found = True
                header_row = idx
                self.logger.info(
                    f"Строка заголовка найдена на позиции {idx + 1} (индекс pandas: {idx})")
                break

        if not found:
            error_msg = (f"Не найден заголовок с колонками {required_headers} "
                         f"на листе '{sheet_name}'. Проверено {rows_searched} строк.")
            self.logger.error(error_msg)
            raise Exception(error_msg)

        # Убедимся, что header_row не None (это должно быть верно из-за проверки выше)
        assert header_row is not None

        # 3. обработка DataFrame
        self.logger.debug(
            f"Обработка данных после строки заголовка {header_row + 1}")
        df_table = df_whole.iloc[header_row:]
        df_table.columns = df_table.iloc[0]
        df_table = df_table[1:]
        df_table = df_table.reset_index(drop=True)
        cols = [col for col in required_headers if col in df_table.columns]
        df_out = df_table[cols]
        self.logger.debug(
            f"DataFrame после обработки заголовка. Размер: {df_out.shape}")

        # 4. фильтрация по красному цвету
        self.logger.info(
            f"Начало фильтрации строк по цвету для таблицы на листе '{sheet_name}'")
        clean_rows = self._get_non_red_rows(
            sheet_name, header_row + 1, len(df_out))
        df_out = df_out.iloc[clean_rows].reset_index(drop=True)
        self.logger.info(
            f"Таблица с листа '{sheet_name}' успешно прочитана и отфильтрована. Итоговый размер: {df_out.shape}")

        return df_out

    def build_structure(self) -> List[Dict[str, Any]]:
        """
        Строит структуру категорий из Excel-файла.

        Returns:
            List[Dict[str, Any]]: Список словарей, представляющих структуру категорий.
        """
        self.logger.info(
            f"Начало построения структуры из файла Excel: {self.filename}")

        try:
            self.logger.info(f"Открытие файла Excel: {self.filename}")
            with pd.ExcelFile(self.filename) as xls:
                self.logger.info(f"Файл {self.filename} успешно открыт.")
                self.logger.debug(
                    f"Доступные листы в файле: {xls.sheet_names}")

                self.logger.info(
                    f"Чтение таблицы категорий с листа '{SHEET_CATEGORIES}'...")
                cats_df = self._read_table_by_header_df(SHEET_CATEGORIES, {
                                                        COL_CATEGORY_TYPE, COL_CATEGORY})
                self.logger.info(
                    f"Таблица категорий прочитана. Размер: {cats_df.shape}")

                self.logger.info(
                    f"Чтение таблицы шаблонов с листа '{SHEET_TEMPLATES}'...")
                expr_df = self._read_table_by_header_df(SHEET_TEMPLATES, {
                                                        COL_TEMPLATE_CATEGORY, COL_TEMPLATE_EXPRESSION})
                self.logger.info(
                    f"Таблица шаблонов прочитана. Размер: {expr_df.shape}")

            self.logger.info(f"Файл {self.filename} закрыт.")
        except Exception as e:
            self.logger.error(
                f"Ошибка при открытии/чтении файла {self.filename}: {e}")
            raise

        initial_cat_count = len(cats_df)
        self.logger.info(
            f"Начало обработки {initial_cat_count} строк категорий")
        cats_df[COL_CATEGORY_TYPE] = cats_df[COL_CATEGORY_TYPE].ffill()
        self.logger.debug(
            "Заполнение пропущенных значений 'Тип категории' завершено")

        cat_types = []
        categories_by_type = defaultdict(list)
        skipped_cat_rows = 0

        for index, row in cats_df.iterrows():
            ctype = str(row[COL_CATEGORY_TYPE]).strip()
            cname = str(row[COL_CATEGORY]).strip()
            if not ctype or ctype.lower() == 'nan':
                self.logger.warning(
                    f"Пропущена строка {index + 1} с пустой '{COL_CATEGORY_TYPE}'")
                skipped_cat_rows += 1
                continue
            if not cname or cname.lower() == 'nan':
                self.logger.warning(
                    f"Пропущена строка {index + 1} с пустой '{COL_CATEGORY}'")
                skipped_cat_rows += 1
                continue
            if ctype not in cat_types:
                cat_types.append(ctype)
                self.logger.debug(f"Найден новый тип категории: '{ctype}'")
            categories_by_type[ctype].append({'name': cname, 'templates': []})
            self.logger.debug(
                f"Добавлена категория '{cname}' к типу '{ctype}'")

        self.logger.info(f"Обработка категорий завершена. "
                         f"Всего строк: {initial_cat_count}, Пропущено: {skipped_cat_rows}, "
                         f"Уникальных типов: {len(cat_types)}, Всего категорий: {sum(len(cats) for cats in categories_by_type.values())}")

        initial_expr_count = len(expr_df)
        self.logger.info(
            f"Начало обработки {initial_expr_count} строк шаблонов")
        templates_by_cat = defaultdict(list)
        skipped_expr_rows = 0

        for index, row in expr_df.iterrows():
            catname = str(row[COL_TEMPLATE_CATEGORY]).strip()
            expr = str(row[COL_TEMPLATE_EXPRESSION]).strip()
            if expr and expr.lower() != 'nan':
                templates_by_cat[catname].append(expr)
                self.logger.debug(
                    f"Добавлен шаблон к категории '{catname}': {expr[:50]}{'...' if len(expr) > 50 else ''}")
            else:
                self.logger.debug(
                    f"Пропущена строка шаблона {index + 1} из-за пустого выражения")
                skipped_expr_rows += 1

        self.logger.info(f"Обработка шаблонов завершена. "
                         f"Всего строк: {initial_expr_count}, Пропущено: {skipped_expr_rows}, "
                         f"Категорий с шаблонами: {len(templates_by_cat)}")

        self.logger.info("Начало формирования финальной структуры")
        structure = []
        total_cats_in_structure = 0
        total_templates_in_structure = 0

        for ctype in cat_types:
            cats_out = []
            for cat in categories_by_type[ctype]:
                cname = cat['name']
                templates = templates_by_cat.get(cname, [])
                cat_dict = {
                    'name': cname,
                    'storageDepth': DEFAULT_STORAGE_DEPTH,
                    'templates': templates
                }
                cats_out.append(cat_dict)
                total_cats_in_structure += 1
                total_templates_in_structure += len(templates)
                if templates:
                    self.logger.debug(
                        f"Категория '{cname}' связана с {len(templates)} шаблонами")

            type_dict = {'name': ctype, 'categories': cats_out}
            structure.append(type_dict)

        self.logger.info(f"Формирование структуры завершено.")
        self.logger.info(f"Итоговая структура: {len(structure)} типов категорий, "
                         f"{total_cats_in_structure} категорий, "
                         f"{total_templates_in_structure} шаблонов")

        return structure


# --- Пример использования ---
if __name__ == "__main__":
    from config import DEBUG_FILE_NAME
    logger_manager = LogManager(DEBUG_FILE_NAME)
    parser = ExcelParser(DEBUG_FILE_NAME, logger_manager)
    structure = parser.build_structure()
    print(structure)
