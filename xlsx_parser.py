import pandas as pd
from collections import defaultdict
import pprint
import logging
from logging_config import setup_logging
from debug_config import *


def read_table_by_header_df(df_whole, required_headers):
    found = False
    header_row = None
    for idx, row in df_whole.iterrows():
        headers = [str(x).strip() for x in row if pd.notna(x)]
        if set(required_headers).issubset(set(headers)):
            found = True
            header_row = idx
            logging.info(
                f"Заголовок найден в строке {header_row+1}: {headers}")
            break
    if not found:
        logging.error(f"Не найден заголовок с колонками {required_headers}")
        raise Exception(f"Не найден заголовок с колонками {required_headers}")
    df_table = df_whole.iloc[header_row:]
    df_table.columns = df_table.iloc[0]
    df_table = df_table[1:]
    df_table = df_table.reset_index(drop=True)
    cols = [col for col in required_headers if col in df_table.columns]
    df_out = df_table[cols]
    logging.info(f"Обработано {len(df_out)} строк с колонками {cols}")
    return df_out


def build_structure_from_excel(filename):
    setup_logging(filename)
    try:
        logging.info(f"Открытие файла Excel: {filename}")
        with pd.ExcelFile(filename) as xls:
            cats_whole = pd.read_excel(
                xls, sheet_name='Категории событий', header=None)
            expr_whole = pd.read_excel(
                xls, sheet_name='Ключевые выражения категорий', header=None)
            logging.info(f"Файл {filename} успешно открыт.")
            cats_df = read_table_by_header_df(
                cats_whole, {'Тип категории', 'Категория'})
            expr_df = read_table_by_header_df(
                expr_whole, {'Категория события', 'Ключевое выражение'})
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
            cats_whole = pd.read_excel(
                xls, sheet_name='Категории событий', header=None)
            expr_whole = pd.read_excel(
                xls, sheet_name='Ключевые выражения категорий', header=None)
            logging.info(f"Файл {filename} успешно открыт.")
            cats_df = read_table_by_header_df(
                cats_whole, {'Тип категории', 'Категория'})
            cats_df['Тип категории'] = cats_df['Тип категории'].ffill()
            expr_df = read_table_by_header_df(
                expr_whole, {'Категория события', 'Ключевое выражение'})
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
