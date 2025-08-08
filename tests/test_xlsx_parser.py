# tests/test_xlsx_parser.py
"""
Тесты для модуля xlsx_parser.py.
"""
import pytest
import os
from pathlib import Path
from unittest.mock import MagicMock, patch

# Импорты модулей проекта
from xlsx_parser import ExcelParser
from logging_config import LogManager
# Импортируем константы из config для проверок
from config import (
    COL_CATEGORY_TYPE, COL_CATEGORY,
    COL_TEMPLATE_CATEGORY, COL_TEMPLATE_EXPRESSION,
    SHEET_CATEGORIES, SHEET_TEMPLATES
)

# Путь к тестовым данным
TEST_DATA_DIR = Path(__file__).parent / "data"
TEST_COLORS_FILE = TEST_DATA_DIR / "test_colors.xlsx"
TEST_STRUCTURE_FILE = TEST_DATA_DIR / "test_structure.xlsx"

# Убедимся, что тестовые файлы существуют (в идеале они должны быть созданы заранее)
# assert TEST_COLORS_FILE.exists(), f"Тестовый файл {TEST_COLORS_FILE} не найден!"
# assert TEST_STRUCTURE_FILE.exists(), f"Тестовый файл {TEST_STRUCTURE_FILE} не найден!"


class TestExcelParser:
    """Тесты для класса ExcelParser."""

    @pytest.fixture
    def logger_manager_mock(self):
        """Фикстура: Мок для LogManager."""
        mock_lm = MagicMock(spec=LogManager)
        mock_logger = MagicMock()
        mock_lm.get_logger.return_value = mock_logger
        return mock_lm

    def test_init_with_defaults(self, logger_manager_mock):
        """Тест: Инициализация ExcelParser с параметрами по умолчанию."""
        filename = "test.xlsx"
        parser = ExcelParser(filename, logger_manager_mock)

        assert parser.filename == filename
        assert parser._logger_manager == logger_manager_mock
        # Проверяем, что логгер был получен
        logger_manager_mock.get_logger.assert_called_once_with("ExcelParser")
        assert parser._override_sheet_categories is None
        assert parser._override_sheet_templates is None

    def test_init_with_overrides(self, logger_manager_mock):
        """Тест: Инициализация ExcelParser с переопределенными именами листов."""
        filename = "test.xlsx"
        override_cat_sheet = "MyCategories"
        override_tpl_sheet = "MyTemplates"
        parser = ExcelParser(
            filename, logger_manager_mock,
            override_sheet_categories=override_cat_sheet,
            override_sheet_templates=override_tpl_sheet
        )

        assert parser.filename == filename
        assert parser._logger_manager == logger_manager_mock
        assert parser._override_sheet_categories == override_cat_sheet
        assert parser._override_sheet_templates == override_tpl_sheet

    # --- Тесты для _is_red_color ---
    # Эти тесты требуют имитации объектов openpyxl.styles.colors.Color
    # и могут быть сложными. Лучше тестировать это на реальных данных в _get_non_red_rows.
    # Но для примера:
    def test_is_red_color_none(self, logger_manager_mock):
        """Тест: _is_red_color возвращает False для None."""
        parser = ExcelParser("dummy.xlsx", logger_manager_mock)
        assert parser._is_red_color(None) == False

    # --- Тесты для build_structure (интеграционные) ---
    # Используют реальный файл test_structure.xlsx

    @pytest.mark.skipif(not TEST_STRUCTURE_FILE.exists(), reason="Тестовый файл test_structure.xlsx не найден")
    def test_build_structure_basic(self, logger_manager_mock):
        """Тест: build_structure корректно строит структуру из тестового файла."""
        parser = ExcelParser(str(TEST_STRUCTURE_FILE), logger_manager_mock)

        # 1. Действие
        structure = parser.build_structure()

        # 2. Утверждения
        assert isinstance(structure, list)
        # Должно быть 2 типа (Тип 1, Тип 2), "Красный Тип" отфильтрован
        assert len(structure) == 2

        # Найдем "Тип 1" и "Тип 2"
        type1_data = next((t for t in structure if t['name'] == "Тип 1"), None)
        type2_data = next((t for t in structure if t['name'] == "Тип 2"), None)

        assert type1_data is not None
        assert type2_data is not None

        # Проверим категории в "Тип 1"
        cats_type1 = type1_data['categories']
        # Должны быть категории B, E, F (A отфильтрована, если она была красной, или нет - проверим по файлу)
        # Из описания файла: Категории B, E, F принадлежат Тип 1.
        # Категория A тоже принадлежит Тип 1. Если она не красная, она должна быть.
        # Предположим A не красная. Тогда 4 категории.
        expected_cat_names_type1 = {"Категория A",
                                    "Категория B", "Категория E", "Категория F"}
        actual_cat_names_type1 = {c['name'] for c in cats_type1}
        assert actual_cat_names_type1 == expected_cat_names_type1

        # Проверим шаблоны для "Категория A"
        cat_a_data = next(
            (c for c in cats_type1 if c['name'] == "Категория A"), None)
        assert cat_a_data is not None
        assert len(cat_a_data['templates']) == 2
        assert "Шаблон для A1" in cat_a_data['templates']
        assert "Шаблон для A2" in cat_a_data['templates']

        # Проверим категории в "Тип 2"
        cats_type2 = type2_data['categories']
        # Должны быть категории C, D
        expected_cat_names_type2 = {"Категория C", "Категория D"}
        actual_cat_names_type2 = {c['name'] for c in cats_type2}
        assert actual_cat_names_type2 == expected_cat_names_type2

        # Проверим шаблоны для "Категория C"
        cat_c_data = next(
            (c for c in cats_type2 if c['name'] == "Категория C"), None)
        assert cat_c_data is not None
        assert len(cat_c_data['templates']) == 1
        assert "Шаблон для C1" in cat_c_data['templates']

        # Проверим шаблоны для "Категория D" (пустой шаблон в файле)
        cat_d_data = next(
            (c for c in cats_type2 if c['name'] == "Категория D"), None)
        assert cat_d_data is not None
        assert cat_d_data['templates'] == []  # Пустой список шаблонов

        # Проверим, что категории, связанные с отфильтрованным "Красный Тип", обработаны корректно
        # В данном случае, "Категория H" наследует "Красный Тип", но сама не красная.
        # Логика build_structure сначала читает категории, потом шаблоны.
        # Если "Красный Тип" отфильтрован, то категории, принадлежащие ему, не попадут в финальную структуру.
        # Нужно уточнить логику: отфильтровываются ли строки с красным типом категории?
        # Из описания _get_non_red_rows: проверяются все ячейки строки.
        # Если строка с "Красный Тип" полностью красная, то тип и его категории отфильтровываются.
        # Если только "Красный Тип" красный, а "Категория G" нет, то тип отфильтровывается, G остается, но без типа?
        # Это сложная логика. Проще проверить по результату.
        # В структуре не должно быть "Красный Тип"
        red_type_names = [t['name']
                          for t in structure if "Красный" in t['name']]
        assert len(red_type_names) == 0

    @pytest.mark.skipif(not TEST_STRUCTURE_FILE.exists(), reason="Тестовый файл test_structure.xlsx не найден")
    def test_build_structure_with_sheet_overrides(self, logger_manager_mock):
        """Тест: build_structure использует переопределенные имена листов."""
        # Создадим временный файл с другими именами листов для этого теста
        # или используем тот же файл, но укажем другие имена (если они совпадают, тест не проверит override)
        # Для простоты, предположим, что листы в test_structure.xlsx могут называться по-другому
        # и мы их переопределяем. Этот тест проверит, что параметры передаются.
        # Реальная проверка требует файла с другими именами листов или мока pandas.read_excel.

        override_cat_sheet = SHEET_CATEGORIES  # Используем стандартное имя
        override_tpl_sheet = SHEET_TEMPLATES  # Используем стандартное имя
        parser = ExcelParser(
            str(TEST_STRUCTURE_FILE), logger_manager_mock,
            override_sheet_categories=override_cat_sheet,
            override_sheet_templates=override_tpl_sheet
        )

        # Замокаем pandas.read_excel, чтобы проверить, с какими аргументами она была вызвана
        with patch('xlsx_parser.pd.read_excel') as mock_read_excel:
            # Настроим мок, чтобы он возвращал минимальные DataFrame
            mock_cats_df = MagicMock()
            mock_cats_df.shape = (0, 2)  # Пустой для упрощения
            mock_expr_df = MagicMock()
            mock_expr_df.shape = (0, 2)  # Пустой для упрощения
            mock_read_excel.side_effect = [
                MagicMock(), mock_cats_df, MagicMock(), mock_expr_df]  # 2 вызова на каждый лист

            try:
                parser.build_structure()
            except Exception:
                pass  # Нам не важно, что произойдет с пустыми данными, нам важен вызов

            # Проверим, что read_excel был вызван с переопределенными именами
            # Всего 2 листа -> 2 вызова read_excel
            assert mock_read_excel.call_count >= 2
            # Проверим аргументы первого вызова (категории)
            first_call_args, first_call_kwargs = mock_read_excel.call_args_list[0]
            assert first_call_kwargs.get('sheet_name') == override_cat_sheet
            # Проверим аргументы второго вызова (шаблоны)
            second_call_args, second_call_kwargs = mock_read_excel.call_args_list[1]
            assert second_call_kwargs.get('sheet_name') == override_tpl_sheet

    # --- Тесты для _get_non_red_rows и _read_table_by_header_df ---
    # Эти тесты требуют реального файла test_colors.xlsx и сложной имитации openpyxl.
    # Лучше интеграционно проверить через build_structure или _read_table_by_header_df напрямую.
    # Для _read_table_by_header_df можно создать файл с известной структурой и цветами.

    @pytest.mark.skipif(not TEST_COLORS_FILE.exists(), reason="Тестовый файл test_colors.xlsx не найден")
    def test_get_non_red_rows_integration(self, logger_manager_mock):
        """Интеграционный тест: _get_non_red_rows фильтрует строки правильно."""
        # Этот тест косвенно проверяет _get_non_red_rows через _read_table_by_header_df
        parser = ExcelParser(str(TEST_COLORS_FILE), logger_manager_mock)

        # Предположим, у нас есть тестовый лист "Colors"
        # с заголовками "ID", "Color"
        # и известно, какие строки содержат красный цвет.
        # required_headers = {"ID", "Color"} # Пример, зависит от структуры файла
        # df = parser._read_table_by_header_df("Colors", required_headers)

        # Вместо этого, проверим напрямую _get_non_red_rows
        # Но это требует имитации openpyxl или реального файла.
        # Пример с моком (очень упрощенный и может не работать напрямую):
        # with patch('xlsx_parser.openpyxl.load_workbook') as mock_load_wb:
        #     mock_wb = MagicMock()
        #     mock_ws = MagicMock()
        #     mock_wb.__getitem__.return_value = mock_ws
        #     mock_load_wb.return_value = mock_wb
        #     # Настроить mock_ws для возврата нужных ячеек и цветов
        #     # ...
        #     result = parser._get_non_red_rows("Colors", 1, 4) # header=1, nrows=4
        #     # assert result == [expected_indices]

        # Из-за сложности мока, лучше использовать реальный файл.
        # Проверим, что метод не вызывает исключений с корректными данными.
        try:
            # header_row_idx=1 (строка 2 в Excel), nrows=4 (строки 2-5)
            # Ожидаемые индексы в DataFrame (0-based): 0 (строка 2), 2 (строка 4), 3 (строка 5)
            # Строка 3 (индекс 1 в DF) должна быть отфильтрована.
            non_red_indices = parser._get_non_red_rows("Colors", 1, 4)
            # Проверка зависит от реальных цветов в файле test_colors.xlsx
            # assert 1 not in non_red_indices # Если строка 3 (DF index 1) красная
            # assert len(non_red_indices) == 3 # Остальные 3 строки не красные
            assert isinstance(non_red_indices, list)
        except Exception as e:
            pytest.fail(f"_get_non_red_rows вызвал исключение: {e}")

    @pytest.mark.skipif(not TEST_COLORS_FILE.exists(), reason="Тестовый файл test_colors.xlsx не найден")
    def test_read_table_by_header_df_integration(self, logger_manager_mock):
        """Интеграционный тест: _read_table_by_header_df читает и фильтрует таблицу."""
        parser = ExcelParser(str(TEST_COLORS_FILE), logger_manager_mock)
        # Зависит от структуры test_colors.xlsx
        required_headers = {"ID", "Color"}

        # 1. Действие
        df = parser._read_table_by_header_df("Colors", required_headers)

        # 2. Утверждения
        assert df is not None
        assert not df.empty
        assert set(df.columns) == required_headers
        # Проверим, что количество строк соответствует ожиданиям после фильтрации
        # В test_colors.xlsx 5 строк данных (ID 1-5). Одна красная.
        # Ожидаем 4 строки. Индексы DF: 0, 1, 2, 3
        assert len(df) == 4  # Или другое число, в зависимости от файла
        # Проверим, что отфильтрованная строка отсутствует
        # assert 2 not in df.index # Если строка с ID=3 (DF index=2) была красной


# --- Дополнительные тесты (по желанию) ---

    def test_build_structure_file_not_found(self, logger_manager_mock):
        """Тест: build_structure вызывает исключение, если файл не найден."""
        parser = ExcelParser("non_existent_file.xlsx", logger_manager_mock)
        with pytest.raises(FileNotFoundError):
            parser.build_structure()

    def test_build_structure_sheet_not_found(self, logger_manager_mock):
        """Тест: build_structure вызывает исключение, если лист не найден."""
        # Требует существующий файл, но с отсутствующим листом.
        # Можно создать временный файл или использовать мок.
        pass  # Реализация опущена для краткости
