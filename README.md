# Генератор CIM/XML из Excel (Категории событий)

Этот проект преобразует структуру категорий событий, описанную в специально форматированном Excel-файле, в XML-документ, соответствующий стандарту CIM (Common Information Model) и кастомной схеме Monitel. Особенностью является возможность "скрытия" строк данных в Excel путем их заливки или окрашивания шрифта в красный цвет.

## 📁 Структура проекта
project_root/
├── main.py # Точка входа в приложение
├── config.json # Файл конфигурации
├── xlsx_parser.py # Модуль для парсинга Excel-файла
├── cim_xml_creator.py # Модуль для генерации CIM/XML
├── logging_config.py # Модуль для настройки логирования
├── requirements.txt # Зависимости проекта
└── README.md # Этот файл

## 🧱 Основные классы

### `ExcelParser` (`xlsx_parser.py`)

**Назначение:** Отвечает за чтение и анализ Excel-файла, извлечение структуры категорий с учетом фильтрации по цвету.

**Методы:**

*   `__init__(self, filename: str, logger_manager: LogManager)`: Инициализирует парсер.
    *   `filename`: Путь к Excel-файлу.
    *   `logger_manager`: Экземпляр `LogManager` для логирования.
*   `_is_red_color(self, color: Optional[openpyxl.styles.colors.Color]) -> bool`: (Приватный) Проверяет, является ли цвет (заливки или шрифта ячейки) "красным" согласно настройкам в `config.json`.
*   `_get_non_red_rows(self, sheet_name: str, header_row_idx: int, nrows: int) -> List[int]`: (Приватный) Возвращает список индексов строк (в терминах pandas DataFrame), которые **не** содержат красной заливки или красного шрифта. Использует `openpyxl` для анализа стилей.
    *   `sheet_name`: Имя листа Excel.
    *   `header_row_idx`: Номер строки заголовка в Excel (начиная с 1).
    *   `nrows`: Количество строк данных для анализа.
*   `_read_table_by_header_df(self, sheet_name: str, required_headers: Set[str]) -> pd.DataFrame`: (Приватный) Читает таблицу с указанного листа, начиная со строки, содержащей все `required_headers`. Фильтрует строки с помощью `_get_non_red_rows`.
    *   `sheet_name`: Имя листа Excel.
    *   `required_headers`: Множество обязательных заголовков столбцов.
*   `build_structure(self) -> List[Dict[str, Any]]`: Основной метод. Читает оба листа Excel ('Категории событий', 'Ключевые выражения категорий'), обрабатывает данные и строит внутреннюю структуру категорий.

### `CIMXMLGenerator` (`cim_xml_creator.py`)

**Назначение:** Преобразует внутреннюю структуру категорий, созданную `ExcelParser`, в XML-документ формата RDF/CIM.

**Методы:**

*   `__init__(self, category_types: List[Dict[str, Any]], parent_obj_uid: Optional[str], logger_manager: LogManager)`: Инициализирует генератор.
    *   `category_types`: Структура категорий, полученная от `ExcelParser`.
    *   `parent_obj_uid`: UID родительского объекта CIM в XML. Если `None`, используется значение из `config.json`.
    *   `logger_manager`: Экземпляр `LogManager` для логирования.
*   `_create_full_model(self, root_element: lxml.etree._Element) -> str`: (Приватный) Создает и добавляет элемент `<md:FullModel>` в корневой элемент XML. Возвращает UID созданного элемента.
*   `_create_category_type(self, root_element: lxml.etree._Element, ct_data: Dict[str, Any], ct_index: int) -> str`: (Приватный) Создает элемент `<me:DjCategoryType>` и все связанные с ним `<me:DjCategory>`. Возвращает UID созданного типа категории.
*   `_create_category(self, root_element: lxml.etree._Element, cat_data: Dict[str, Any], cat_uid: str, ct_uid: str, cat_index: int)`: (Приватный) Создает элемент `<me:DjCategory>` и связывает его с типом категории. Также создает связанные `<me:DjRecordTemplate>`.
*   `_create_template(self, root_element: lxml.etree._Element, tmpl_text: str, cat_uid: str, cat_elem: lxml.etree._Element, tmpl_index: int)`: (Приватный) Создает элемент `<me:DjRecordTemplate>` и связывает его с категорией.
*   `create_xml(self) -> str`: Основной метод. Координирует создание всех элементов XML и возвращает итоговую XML-строку с декларацией и отступами.

### `LogManager` (`logging_config.py`)

**Назначение:** Централизованное управление логированием приложения. Настраивает логирование как в файл (в папку `./log`), так и в консоль.

**Методы:**

*   `__init__(self, xlsx_filename: str, log_level: int = logging.INFO)`: Инициализирует менеджер логирования и настраивает обработчики.
    *   `xlsx_filename`: Имя Excel-файла, используется для формирования имени лог-файла.
    *   `log_level`: Уровень логирования (по умолчанию `INFO`).
*   `_setup_logging(self) -> str`: (Приватный) Выполняет фактическую настройку `logging`. Создает папку `./log` и файл лога. Возвращает путь к лог-файлу.
*   `get_logger(self, name: str = __name__) -> logging.Logger`: Возвращает настроенный экземпляр логгера для использования в других модулях/классах.

## ⚙️ Конфигурация (`config.json`)

Файл `config.json` содержит все настройки проекта.

*   **`debug`**: Настройки для отладочного запуска.
    *   `parent_uid`: UID родительского объекта по умолчанию.
    *   `file_name`: Имя Excel-файла по умолчанию.
*   **`colors`**: Параметры для определения "красного" цвета в Excel.
    *   `red_like_colors`: Список ARGB-значений, считающихся красными.
    *   `red_indexed_colors`: Список индексов стандартных цветов Excel, считающихся красными.
    *   `red_theme_indices`: Список индексов тематических цветов, которые могут быть красными.
    *   `red_theme_tint_threshold`: Пороговое значение `tint` для тематических цветов.
*   **`defaults`**: Значения по умолчанию для полей в XML.
    *   `storage_depth`: Глубина хранения данных (в днях).
    *   `order`: Порядок по умолчанию.
    *   `message_required` и др.: Булевы флаги по умолчанию.
*   **`formatting`**: Форматы дат и других данных.
    *   `model_created_format`: Формат даты для поля `Model.created`.
*   **`namespaces`**: Пространства имен XML (RDF, MD, CIM, ME).
*   **`sheet_names`**: Имена листов в Excel-файле.
*   **`column_names`**: Названия обязательных колонок в Excel.
*   **`paths`**: Пути к директориям.
    *   `log_dir`: Директория для лог-файлов.

## 🧠 Логика работы

1.  **Инициализация (`main.py`)**:
    *   Создается экземпляр `LogManager`, который настраивает логирование в файл и консоль.
    *   Создаются экземпляры `ExcelParser` и `CIMXMLGenerator`.

2.  **Парсинг Excel (`ExcelParser`)**:
    *   `build_structure()` вызывает `_read_table_by_header_df()` для каждого из двух листов Excel.
    *   `_read_table_by_header_df()`:
        *   Читает весь лист в `pandas.DataFrame`.
        *   Ищет строку, содержащую все обязательные заголовки.
        *   Создает новый DataFrame, начиная с этой строки.
        *   Вызывает `_get_non_red_rows()` для определения индексов строк без красного цвета.
        *   Фильтрует DataFrame, оставляя только "некрасные" строки.
    *   `_get_non_red_rows()`:
        *   Использует `openpyxl` для загрузки файла (`data_only=True`).
        *   Итерируется по указанному диапазону строк.
        *   Для каждой строки проверяет все ячейки на наличие красной заливки (`fill.fgColor`) или красного шрифта (`font.color`).
        *   Использует `_is_red_color()` для проверки каждого цвета.
        *   Возвращает список индексов строк, не содержащих красного.
    *   `build_structure()` объединяет данные из двух таблиц, связывая категории с их шаблонами, и формирует внутреннюю структуру данных (список словарей).

3.  **Генерация XML (`CIMXMLGenerator`)**:
    *   `create_xml()` создает корневой элемент `<rdf:RDF>` с необходимыми namespace'ами.
    *   Создает элемент `<md:FullModel>` с метаинформацией.
    *   Итерируется по структуре категорий, полученной от `ExcelParser`.
    *   Для каждого типа категории вызывает `_create_category_type()`.
    *   `_create_category_type()` создает `<me:DjCategoryType>` и затем для каждой категории вызывает `_create_category()`.
    *   `_create_category()` создает `<me:DjCategory>`, заполняет его атрибуты и вызывает `_create_template()` для каждого шаблона.
    *   `_create_template()` создает `<me:DjRecordTemplate>`.
    *   Все связи между элементами (например, `ParentObject`, `Categories`, `CategoryType`, `Templates`) устанавливаются с использованием атрибутов `rdf:resource`.
    *   `create_xml()` сериализует построенное XML-дерево в строку с отступами и XML-декларацией.

4.  **Завершение (`main.py`)**:
    *   Полученная XML-строка записывается в файл с именем, соответствующим имени исходного Excel-файла (например, `data.xml`).
    *   В консоль и лог-файл выводится сообщение об успешном завершении.

## 📦 Зависимости

*   `pandas`: Для удобной работы с табличными данными из Excel.
*   `openpyxl`: Для чтения Excel-файлов и анализа стилей ячеек.
*   `lxml`: Для эффективного построения и сериализации XML-документов.

## 🛠 Установка и запуск

1.  Установите Python 3.x.
2.  Установите зависимости: `pip install -r requirements.txt`.
3.  Настройте `config.json` (при необходимости).
4.  Запустите приложение: `python main.py`.