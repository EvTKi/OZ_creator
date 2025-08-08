# config.py
"""
Модуль конфигурации проекта.
Загружает настройки из config.json.
"""

import json
import os
import uuid
from typing import Dict, Any
import sys

# --- Путь к конфигурационному файлу ---
# Определяем путь к config.json относительно этого модуля
# Это важно для PyInstaller, который может запускать из временной папки


def get_config_path() -> str:
    """Получает путь к config.json, учитывая PyInstaller."""
    if getattr(sys, 'frozen', False):
        # Если запущено как .exe, ищем рядом с исполняемым файлом
        application_path = os.path.dirname(sys.executable)
    else:
        # Если запущено как скрипт, ищем в текущей директории
        application_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(application_path, 'config.json')

# --- Загрузка конфигурации ---


def load_config() -> Dict[str, Any]:
    """Загружает конфигурацию из config.json."""
    config_path = get_config_path()
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(
            f"Конфигурационный файл не найден: {config_path}")
    except json.JSONDecodeError as e:
        raise ValueError(f"Ошибка парсинга config.json: {e}")


# --- Глобальная константа с конфигурацией ---
CONFIG = load_config()

# --- Экспорт значений для удобства ---
# Debug
DEBUG_PARENT_UID: str = CONFIG['debug']['parent_uid']
DEBUG_FILE_NAME: str = CONFIG['debug']['file_name']

# Colors
RED_LIKE_COLORS: set = set(CONFIG['colors']['red_like_colors'])
RED_INDEXED_COLORS: set = set(CONFIG['colors']['red_indexed_colors'])
RED_THEME_INDICES: set = set(CONFIG['colors']['red_theme_indices'])
RED_THEME_TINT_THRESHOLD: float = CONFIG['colors']['red_theme_tint_threshold']

# Defaults
DEFAULT_STORAGE_DEPTH: int = CONFIG['defaults']['storage_depth']
DEFAULT_ORDER: int = CONFIG['defaults']['order']
DEFAULT_MESSAGE_REQUIRED: bool = CONFIG['defaults']['message_required']
DEFAULT_REPORT_HIGHER: bool = CONFIG['defaults']['report_higher']
DEFAULT_SHIFT_RESTRICTED: bool = CONFIG['defaults']['shift_restricted']
DEFAULT_ACTION_TIME_SHIFT: bool = CONFIG['defaults']['action_time_shift']

# Formatting
MODEL_CREATED_FORMAT: str = CONFIG['formatting']['model_created_format']

# Namespaces
NSMAP: Dict[str, str] = CONFIG['namespaces']

# Sheet & Column Names
SHEET_CATEGORIES: str = CONFIG['sheet_names']['categories']
SHEET_TEMPLATES: str = CONFIG['sheet_names']['templates']
COL_CATEGORY_TYPE: str = CONFIG['column_names']['category_type']
COL_CATEGORY: str = CONFIG['column_names']['category']
COL_TEMPLATE_CATEGORY: str = CONFIG['column_names']['template_category']
COL_TEMPLATE_EXPRESSION: str = CONFIG['column_names']['template_expression']

# Paths
LOG_DIR: str = CONFIG['paths']['log_dir']

# --- Функции ---


def generate_uid() -> str:
    """Генерирует универсальный уникальный идентификатор (UUID4).

    Returns:
        str: Строка UUID4 в формате 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'.
    """
    return str(uuid.uuid4())

# --- Для PyInstaller ---
# Убедимся, что config.json будет включен в сборку
# Это нужно указать в .spec файле PyInstaller:
# datas=[('config.json', '.')],


# --- Пример использования ---
# print(f"Debug file: {DEBUG_FILE_NAME}")
# print(f"Red colors: {RED_LIKE_COLORS}")
# print(f"Namespaces: {NSMAP}")
