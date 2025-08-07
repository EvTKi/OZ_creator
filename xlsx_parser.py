# xlsx_parser.py
import pandas as pd
from collections import defaultdict
import pprint
import logging
from logging_config import setup_logging
from config import *
import openpyxl
from config import *


def get_non_red_rows(excel_filename, sheet_name, header_row_idx, nrows):
    """
    Возвращает индексы строк (относительно DataFrame), в которых отсутствует красный цвет 
    в заливке ячеек или цвете шрифта.

    Функция анализирует указанный диапазон строк в Excel-файле с использованием `openpyxl`, 
    чтобы проверить, содержит ли хотя бы одна ячейка в строке цвет, распознаваемый как "красный".
    Строки, содержащие красный цвет, считаются "помеченными на удаление" или "неактивными" 
    и исключаются из результата.

    Поддерживаемые типы красного цвета:
    - RGB-цвета в формате ARGB или RGB (например, 'FFFF0000', 'FFCC0000').
    - Индексированные цвета Excel (например, индекс 10 — стандартный красный).
    - Тематические цвета (theme colors), часто используемые в стилях Excel, 
      особенно при применении встроенных тем оформления.
    - Учитывается параметр `tint` (затемнение/осветление), чтобы избежать ложных срабатываний 
      на сильно осветлённых оттенках.

    Алгоритм:
    1. Загружает рабочую книгу с помощью `openpyxl`.
    2. Для каждой строки в указанном диапазоне проверяет все ячейки.
    3. Проверяет цвет заливки (`fill.fgColor`) и цвет шрифта (`font.color`).
    4. Цвет считается красным, если:
        - Его RGB-значение входит в предопределённый набор "красных" цветов, ИЛИ
        - Это тематический цвет с номером, ассоциированным с красным (5, 6, 7), 
          и не сильно осветлён (`tint > -0.5`), ИЛИ
        - Это индексированный цвет, соответствующий красному (например, 3, 10, 46).
    5. Если хотя бы одна ячейка в строке содержит красный цвет — строка отбрасывается.

    Примечания:
    - Используется `data_only=True` — читаются значения формул, а не формулы.
    - Функция устойчива к отсутствующим атрибутам цвета (обрабатывает `None`, `auto` и т.п.).
    - Индексация в Excel начинается с 1; индексация в pandas — с 0.
      Результат возвращается в индексации pandas (относительно начала таблицы после заголовка).

    Args:
        excel_filename (str): Путь к Excel-файлу (.xlsx).
        sheet_name (str): Имя листа, на котором выполняется проверка.
        header_row_idx (int): Номер строки заголовка в Excel (индексация с 1).
                              Следующие строки (header_row_idx + 1 и далее) анализируются.
        nrows (int): Количество строк для анализа (начиная сразу после заголовка).

    Returns:
        list[int]: Список индексов строк (в индексации pandas/DataFrame), 
                   в которых НЕТ красного цвета ни в одной ячейке.
                   Например, если первая строка данных (после заголовка) не красная, 
                   в список попадёт 0.

    Raises:
        FileNotFoundError: Если файл не найден.
        KeyError: Если лист с указанным именем отсутствует.
        Exception: При ошибках чтения файла (недоступность, повреждение и т.п.).

    Example:
        >>> get_non_red_rows("events.xlsx", "Категории", 3, 10)
        [0, 1, 2, 4, 5, 7, 8]  # 3 строки с красным цветом были отфильтрованы
    """

    def is_red_color(color):
        """
        Проверяет, является ли цвет (Font.color или Fill.fgColor) красным.
        Поддерживает строки, объекты RGB, theme, indexed.
        Обернута в try-except для устойчивости.
        """
        if color is None:
            return False

        # 1. Проверка по RGB (разные форматы)
        try:
            if hasattr(color, 'rgb') and color.rgb is not None:
                rgb_val = color.rgb
                rgb_str = None

                if isinstance(rgb_val, str):
                    rgb_str = rgb_val
                elif hasattr(rgb_val, 'rgb') and isinstance(rgb_val.rgb, str):
                    rgb_str = rgb_val.rgb  # случай: RGB(fFFFF0000)
                elif str(rgb_val).startswith('RGB:'):
                    try:
                        rgb_hex = str(rgb_val).split(':', 1)[1].strip()
                        if len(rgb_hex) in (6, 8) and all(c in '0123456789ABCDEFabcdef' for c in rgb_hex):
                            rgb_str = rgb_hex
                    except Exception as e:
                        logging.debug(
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
            logging.debug(f"Ошибка при обработке RGB цвета {color}: {e}")

        # 2. Theme color
        try:
            if hasattr(color, 'theme') and color.theme is not None:
                if color.theme in (5, 6, 7):
                    tint = getattr(color, 'tint', 0.0)
                    if isinstance(tint, (int, float)) and tint > -0.5:
                        return True
        except Exception as e:
            logging.debug(f"Ошибка при обработке theme цвета {color}: {e}")

        # 3. Indexed color
        try:
            if hasattr(color, 'indexed'):
                if color.indexed in (3, 10, 46):
                    return True
        except Exception as e:
            logging.debug(f"Ошибка при обработке indexed цвета {color}: {e}")

        return False

    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    ws = wb[sheet_name]
    result = []

    for ex_idx in range(header_row_idx + 1, header_row_idx + 1 + nrows):
        is_red = False
        for cell in ws[ex_idx]:
            # Проверка заливки
            fill = cell.fill
            fg_color = getattr(fill, 'fgColor', None)
            if fg_color is not None and is_red_color(fg_color):
                is_red = True
                break

            # Проверка шрифта
            font_color = getattr(cell.font, 'color', None)
            if font_color is not None and is_red_color(font_color):
                is_red = True
                break

        if not is_red:
            df_index = ex_idx - (header_row_idx + 1)
            result.append(df_index)

    return result


def read_table_by_header_df(excel_filename, sheet_name, required_headers):
    # 1. чтение листа
    df_whole = pd.read_excel(
        excel_filename, sheet_name=sheet_name, header=None)
    # 2. поиск строки шапки
    found = False
    header_row = None
    for idx, row in df_whole.iterrows():
        headers = [str(x).strip() for x in row if pd.notna(x)]
        if set(required_headers).issubset(set(headers)):
            found = True
            header_row = idx
            break
    if not found:
        raise Exception(
            f"Не найден заголовок с колонками {required_headers} на листе {sheet_name}")
    # 3. обработка DataFrame
    df_table = df_whole.iloc[header_row:]
    df_table.columns = df_table.iloc[0]
    df_table = df_table[1:]
    df_table = df_table.reset_index(drop=True)
    cols = [col for col in required_headers if col in df_table.columns]
    df_out = df_table[cols]

    # 4. фильтрация по красному цвету
    clean_rows = get_non_red_rows(
        excel_filename, sheet_name, header_row+1, len(df_out))
    df_out = df_out.iloc[clean_rows].reset_index(drop=True)
    return df_out


def build_structure_from_excel(filename):
    setup_logging(filename)
    try:
        logging.info(f"Открытие файла Excel: {filename}")
        with pd.ExcelFile(filename) as xls:
            logging.info(f"Файл {filename} успешно открыт.")
            cats_df = read_table_by_header_df(filename, 'Категории событий', {
                                              'Тип категории', 'Категория'})
            expr_df = read_table_by_header_df(filename, 'Ключевые выражения категорий', {
                                              'Категория события', 'Ключевое выражение'})
        logging.info(f"Файл {filename} закрыт.")
    except Exception as e:
        logging.error(f"Ошибка при открытии файла {filename}: {e}")
        raise

    cats_df['Тип категории'] = cats_df['Тип категории'].ffill()

    cat_types = []
    categories_by_type = defaultdict(list)
    for _, row in cats_df.iterrows():
        ctype = str(row['Тип категории']).strip()
        cname = str(row['Категория']).strip()
        if not ctype or ctype.lower() == 'nan':
            logging.warning(f"Пропущена строка с пустой 'Тип категории'")
            continue
        if not cname or cname.lower() == 'nan':
            logging.warning(f"Пропущена строка с пустой 'Категория'")
            continue
        if ctype not in cat_types:
            cat_types.append(ctype)
        categories_by_type[ctype].append({'name': cname, 'templates': []})

    templates_by_cat = defaultdict(list)
    for _, row in expr_df.iterrows():
        catname = str(row['Категория события']).strip()
        expr = str(row['Ключевое выражение']).strip()
        if expr and expr.lower() != 'nan':
            templates_by_cat[catname].append(expr)

    structure = []
    for ctype in cat_types:
        cats_out = []
        for cat in categories_by_type[ctype]:
            cname = cat['name']
            templates = templates_by_cat.get(cname, [])
            cats_out.append(
                {'name': cname, 'storageDepth': DEFAULT_STORAGE_DEPTH, 'templates': templates})
        structure.append({'name': ctype, 'categories': cats_out})

    logging.info(f"Построена структура: {len(structure)} типов категорий")
    return structure


if __name__ == "__main__":
    filename = DEBUG_FILE_NAME
    setup_logging(filename)
    try:
        logging.info(f"Открытие файла Excel: {filename}")
        with pd.ExcelFile(filename) as xls:
            logging.info(f"Файл {filename} успешно открыт.")
            cats_df = read_table_by_header_df(filename, 'Категории событий', {
                                              'Тип категории', 'Категория'})
            cats_df['Тип категории'] = cats_df['Тип категории'].ffill()
            expr_df = read_table_by_header_df(filename, 'Ключевые выражения категорий', {
                                              'Категория события', 'Ключевое выражение'})
        logging.info(f"Файл {filename} закрыт.")
        print("\n--- DataFrame: Категории событий ---")
        print(cats_df)
        print("\n--- DataFrame: Ключевые выражения категорий ---")
        print(expr_df)
        structure = build_structure_from_excel(filename)
        print("\n--- Структура для XML (фрагмент) ---")
        pprint.pprint(structure[:3])
        print(f"\nВсего типов категорий: {len(structure)}")
    except Exception as e:
        logging.error(f"Ошибка при работе с файлом {filename}: {e}")
        print(f"[ERROR] {e}")
