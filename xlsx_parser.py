import pandas as pd
from collections import defaultdict
import pprint
import logging
from logging_config import setup_logging
from debug_config import *
import openpyxl


def get_non_red_rows(excel_filename, sheet_name, header_row_idx, nrows):
    """
    Возвращает список индексов (относительно DataFrame!) строк, в которых НЕТ красного цвета ни в шрифте, ни в заливке.
    header_row_idx — строка с заголовками (индекс в openpyxl, начинается с 1).
    nrows — сколько строк читать (длина DataFrame после header).
    """
    wb = openpyxl.load_workbook(excel_filename, data_only=True)
    ws = wb[sheet_name]
    result = []
    for ex_idx in range(header_row_idx+1, header_row_idx+1+nrows):
        is_red = False
        for cell in ws[ex_idx]:
            # Проверяем заливку
            fill = cell.fill
            fgColor = getattr(fill, 'fgColor', None)
            color_str = ""
            try:
                if fgColor is not None:
                    if hasattr(fgColor, 'rgb') and isinstance(fgColor.rgb, str):
                        color_str = fgColor.rgb
                    elif hasattr(fgColor, 'rgb') and fgColor.rgb is not None:
                        color_str = str(fgColor.rgb)
                if color_str and color_str.upper() == 'FFFF0000':
                    is_red = True
            except Exception:
                pass
            # Проверяем шрифт
            font_color = getattr(cell.font, 'color', None)
            font_color_str = ""
            try:
                if font_color is not None:
                    if hasattr(font_color, 'rgb') and isinstance(font_color.rgb, str):
                        font_color_str = font_color.rgb
                    elif hasattr(font_color, 'rgb') and font_color.rgb is not None:
                        font_color_str = str(font_color.rgb)
                if font_color_str and font_color_str.upper() == 'FFFF0000':
                    is_red = True
            except Exception:
                pass
            if is_red:
                break
        if not is_red:
            # индекс относительно DataFrame
            result.append(ex_idx - (header_row_idx+1))
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
                {'name': cname, 'storageDepth': 1095, 'templates': templates})
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
