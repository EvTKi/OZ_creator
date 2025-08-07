# config.py
from dataclasses import dataclass
from typing import List, Set, Optional
import uuid


# ———————————————————————
# 🎨 Цвета, ассоциированные с "красным" (в формате FFRRGGBB)
# ———————————————————————
RED_LIKE_COLORS: Set[str] = {
    'FFFF0000',  # Чистый красный
    'FFCC0000',  # Тёмно-красный
    'FF990000',
    'FF660000',
    'FF330000',
    'FF8B0000',  # Deep Red
    'FFFF3333',  # Светло-красный
    'FFFF6666',
    'FFFF9999',
    'FFFFCCCC',
    'FFD32121',  # Red (Office)
    'FFB80C0C',  # Dark Red
    'FFA52222',
    'FFE66161',  # Light Red
    'FFE74C3C',  # Flat UI Red
    'FFC0392B',  # Alizarin
    'FFD91E18',  # Material Red
}

# ———————————————————————
# 📊 Индексированные цвета Excel, ассоциированные с красным
# https://openpyxl.readthedocs.io/en/stable/styles.html#colours
# ———————————————————————
RED_INDEXED_COLORS: Set[int] = {3,   # Red (в стандартной палитре)
                                10,  # Bright Red
                                46}  # Accent2 Red (в некоторых темах)

# ———————————————————————
# 🎭 Темы Excel, которые могут быть красными (theme + tint)
# ———————————————————————
RED_THEME_INDICES: Set[int] = {5, 6, 7}  # Условно: красные темы
# tint > -0.5 считается "достаточно красным"
RED_THEME_TINT_THRESHOLD: float = -0.5

# ———————————————————————
# ⏳ Хранение: глубина по умолчанию (в днях)
# ———————————————————————
DEFAULT_STORAGE_DEPTH: int = 1095  # 3 года

# ———————————————————————
# 📅 Формат даты для Model.created
# ———————————————————————
MODEL_CREATED_FORMAT: str = "%Y-%m-%dT%H:%M:%S.000Z"

# ———————————————————————
# 🧩 RDF/MD/CIM Namespaces
# ———————————————————————
NSMAP = {
    'rdf': 'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
    'md': 'http://iec.ch/TC57/61970-552/ModelDescription/1#',
    'cim': 'http://iec.ch/TC57/2014/CIM-schema-cim16#',
    'me': 'http://monitel.com/2014/schema-cim16#',
}

# ———————————————————————
# 🧱 Значения по умолчанию для полей
# ———————————————————————
DEFAULT_ORDER: int = 0
DEFAULT_MESSAGE_REQUIRED: bool = False
DEFAULT_REPORT_HIGHER: bool = False
DEFAULT_SHIFT_RESTRICTED: bool = False
DEFAULT_ACTION_TIME_SHIFT: bool = False

# ———————————————————————
# 🆔 Глобальный UID (только для отладки!)
# В реальной системе должен передаваться!
# ———————————————————————


def generate_uid() -> str:
    return str(uuid.uuid4())


# Можно заменить на конкретный, если нужно
DEBUG_PARENT_UID: str = '0377FACB-0EA4-4990-A4DD-DC9DE6BFB5B4'
DEBUG_FILE_NAME = 'Опросный лист ОЖ. Станции.xlsx'
