# cim_xml_creator.py
"""
Модуль для генерации CIM/XML-файла из структуры категорий.
"""

import logging
from datetime import datetime
from typing import List, Dict, Any, Optional
from lxml import etree
# Импортируем константы из нового config.py
from config import (
    NSMAP, MODEL_CREATED_FORMAT, DEFAULT_STORAGE_DEPTH,
    DEFAULT_ORDER, DEBUG_PARENT_UID, generate_uid
)
from logging_config import LogManager


class CIMXMLGenerator:
    """
    Класс для генерации CIM/XML-документа из структуры категорий.
    """

    def __init__(self, category_types: List[Dict[str, Any]], parent_obj_uid: Optional[str], logger_manager: LogManager):
        """
        Инициализирует генератор XML.

        Args:
            category_types (List[Dict[str, Any]]): Структура категорий.
            parent_obj_uid (Optional[str]): UID родительского объекта.
            logger_manager (LogManager): Менеджер логирования.
        """
        self.category_types = category_types
        self.parent_obj_uid = parent_obj_uid or DEBUG_PARENT_UID
        self.logger = logger_manager.get_logger(self.__class__.__name__)

    def _create_full_model(self, root_element: etree._Element) -> str:
        """
        Создает и добавляет элемент FullModel в корневой элемент XML.

        Args:
            root_element (etree._Element): Корневой элемент RDF XML-документа.

        Returns:
            str: Сгенерированный UID для элемента FullModel.
        """
        fm_uid = generate_uid()
        fm = etree.SubElement(
            root_element,
            '{%s}FullModel' % NSMAP['md'],
            attrib={"{%s}about" % NSMAP['rdf']: f"#_{fm_uid}"}
        )
        created = datetime.utcnow().strftime(MODEL_CREATED_FORMAT)
        etree.SubElement(fm, '{%s}Model.created' % NSMAP['md']).text = created
        etree.SubElement(fm, '{%s}Model.version' %
                         NSMAP['md']).text = "ver:sample;"
        etree.SubElement(fm, '{%s}Model.name' % NSMAP['me']).text = "CIM16"
        return fm_uid

    def _create_category_type(self, root_element: etree._Element, ct_data: Dict[str, Any], ct_index: int) -> str:
        """
        Создает элемент DjCategoryType и связанные с ним DjCategory.

        Args:
            root_element (etree._Element): Корневой элемент XML.
            ct_data (Dict[str, Any]): Данные типа категории.
            ct_index (int): Индекс типа категории.

        Returns:
            str: UID созданного элемента DjCategoryType.
        """
        ct_name = ct_data.get('name', 'Unnamed Type').strip()
        if not ct_name or ct_name.lower() == 'nan':
            self.logger.warning("Пропущен тип категории с пустым именем")
            return ""  # или raise exception

        ct_uid = generate_uid()
        ct_elem = etree.SubElement(
            root_element,
            '{%s}DjCategoryType' % NSMAP['me'],
            attrib={"{%s}about" % NSMAP['rdf']: f"#_{ct_uid}"}
        )
        etree.SubElement(
            ct_elem,
            '{%s}IdentifiedObject.name' % NSMAP['cim']
        ).text = ct_name
        etree.SubElement(
            ct_elem,
            '{%s}IdentifiedObject.ParentObject' % NSMAP['me'],
            attrib={'{%s}resource' % NSMAP['rdf']: f"#_{self.parent_obj_uid}"}
        )
        etree.SubElement(
            ct_elem,
            '{%s}DjCategoryType.order' % NSMAP['me']
        ).text = str(ct_index)

        # Генерируем UID для категорий один раз
        categories = ct_data.get('categories', [])
        cat_uids = [generate_uid() for _ in categories]

        # Связь: DjCategoryType → DjCategory (через Categories)
        for uid in cat_uids:
            etree.SubElement(
                ct_elem,
                '{%s}DjCategoryType.Categories' % NSMAP['me'],
                attrib={'{%s}resource' % NSMAP['rdf']: f"#_{uid}"}
            )

        # Создаём объекты DjCategory
        for idx, cat_data in enumerate(categories):
            self._create_category(root_element, cat_data,
                                  cat_uids[idx], ct_uid, idx)

        return ct_uid

    def _create_category(self, root_element: etree._Element, cat_data: Dict[str, Any], cat_uid: str, ct_uid: str, cat_index: int):
        """
        Создает элемент DjCategory.

        Args:
            root_element (etree._Element): Корневой элемент XML.
            cat_data (Dict[str, Any]): Данные категории.
            cat_uid (str): UID категории.
            ct_uid (str): UID родительского типа категории.
            cat_index (int): Индекс категории внутри типа.
        """
        cat_name = cat_data.get('name', 'Unnamed Category').strip()
        if not cat_name or cat_name.lower() == 'nan':
            self.logger.warning(
                f"Пропущена категория без имени в типе с UID '{ct_uid}'")
            return

        cat_elem = etree.SubElement(
            root_element,
            '{%s}DjCategory' % NSMAP['me'],
            attrib={"{%s}about" % NSMAP['rdf']: f"#_{cat_uid}"}
        )
        etree.SubElement(
            cat_elem,
            '{%s}IdentifiedObject.name' % NSMAP['cim']
        ).text = cat_name
        etree.SubElement(
            cat_elem,
            '{%s}IdentifiedObject.ParentObject' % NSMAP['me'],
            attrib={'{%s}resource' % NSMAP['rdf']: f"#_{ct_uid}"}
        )
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.order' % NSMAP['me']
        ).text = str(cat_index)
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.actionTimeWithinShift' % NSMAP['me']
        ).text = "false"
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.messageRequired' % NSMAP['me']
        ).text = "false"
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.reportForHigherOperationalStaff' % NSMAP['me']
        ).text = "false"
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.shiftPersonRestricted' % NSMAP['me']
        ).text = "false"
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.storageDepth' % NSMAP['me']
        ).text = str(cat_data.get('storageDepth', DEFAULT_STORAGE_DEPTH))

        # Связь: DjCategory → DjCategoryType
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.CategoryType' % NSMAP['me'],
            attrib={'{%s}resource' % NSMAP['rdf']: f"#_{ct_uid}"}
        )

        # Шаблоны
        templates = cat_data.get('templates', [])
        for idx, tmpl_text in enumerate(templates):
            if not tmpl_text or str(tmpl_text).lower() in ('nan', 'none', ''):
                continue
            self._create_template(root_element, tmpl_text,
                                  cat_uid, cat_elem, idx)

    def _create_template(self, root_element: etree._Element, tmpl_text: str, cat_uid: str, cat_elem: etree._Element, tmpl_index: int):
        """
        Создает элемент DjRecordTemplate и связывает его с категорией.

        Args:
            root_element (etree._Element): Корневой элемент XML.
            tmpl_text (str): Текст шаблона.
            cat_uid (str): UID категории, к которой принадлежит шаблон.
            cat_elem (etree._Element): Элемент категории, к которому добавляется связь.
            tmpl_index (int): Индекс шаблона внутри категории.
        """
        tmpl_uid = generate_uid()
        # Связь: DjCategory → DjRecordTemplate
        etree.SubElement(
            cat_elem,
            '{%s}DjCategory.Templates' % NSMAP['me'],
            attrib={'{%s}resource' % NSMAP['rdf']: f"#_{tmpl_uid}"}
        )
        tmpl_elem = etree.SubElement(
            root_element,
            '{%s}DjRecordTemplate' % NSMAP['me'],
            attrib={"{%s}about" % NSMAP['rdf']: f"#_{tmpl_uid}"}
        )
        etree.SubElement(
            tmpl_elem,
            '{%s}DjRecordTemplate.order' % NSMAP['me']
        ).text = str(tmpl_index)
        etree.SubElement(
            tmpl_elem,
            '{%s}DjRecordTemplate.text' % NSMAP['me']
        ).text = str(tmpl_text)
        etree.SubElement(
            tmpl_elem,
            '{%s}DjRecordTemplate.Category' % NSMAP['me'],
            attrib={'{%s}resource' % NSMAP['rdf']: f"#_{cat_uid}"}
        )

    def create_xml(self) -> str:
        """
        Преобразует структуру категорий в CIM/XML в формате RDF.

        Returns:
            str: XML-строка с декларацией и отступами.

        Raises:
            Exception: При ошибках построения XML.
        """
        try:
            self.logger.info("Запуск формирования CIM-XML…")

            # Корень RDF
            root = etree.Element('{%s}RDF' % NSMAP['rdf'], nsmap=NSMAP)

            # FullModel
            self._create_full_model(root)

            count_types = 0
            count_cats = 0
            count_templates = 0

            for idx, ct in enumerate(self.category_types):
                ct_uid = self._create_category_type(root, ct, idx)
                if ct_uid:  # Если тип был создан успешно
                    count_types += 1
                    # Подсчет категорий и шаблонов можно улучшить, передавая счетчики в методы
                    # или возвращая их из методов. Пока оставим как есть для простоты.
                    categories = ct.get('categories', [])
                    count_cats += len([c for c in categories if c.get('name',
                                      '').strip().lower() != 'nan'])
                    for cat in categories:
                        count_templates += len([t for t in cat.get('templates', [])
                                               if t and str(t).lower() not in ('nan', 'none', '')])

            self.logger.info(
                f"Сформировано: {count_types} типов, {count_cats} категорий, {count_templates} шаблонов"
            )
            self.logger.info("CIM-XML построен успешно.")

            return etree.tostring(
                root,
                encoding='utf-8',
                pretty_print=True,
                xml_declaration=True
            ).decode('utf-8')

        except Exception as e:
            self.logger.error(f"Ошибка при формировании CIM-XML: {e}")
            raise


# --- Пример использования ---
if __name__ == "__main__":
    # Предполагается, что structure уже получена
    from config import DEBUG_FILE_NAME, DEBUG_PARENT_UID
    logger_manager = LogManager(DEBUG_FILE_NAME)
    # structure = ... # откуда-то получаем структуру
    # generator = CIMXMLGenerator(structure, DEBUG_PARENT_UID, logger_manager)
    # xml_content = generator.create_xml()
    # print(xml_content)
