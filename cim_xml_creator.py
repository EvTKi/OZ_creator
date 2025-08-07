# cim_xml_creator
"""
Модуль для генерации CIM/XML-файла из структуры категорий.
Использует lxml для построения XML в соответствии с CIM-стандартом и кастомной схемой Monitel.
"""

import logging
from datetime import datetime
from lxml import etree
from config import (
    NSMAP,
    MODEL_CREATED_FORMAT,
    DEFAULT_STORAGE_DEPTH,
    DEFAULT_ORDER,
    DEBUG_PARENT_UID,
    generate_uid,
)


def create_cim_xml(category_types, parent_obj_uid=None):
    """
    Преобразует структуру категорий в CIM/XML в формате RDF.

    Args:
        category_types (list[dict]): Список типов категорий в формате:
            [
              {
                'name': 'Тип 1',
                'categories': [
                  {
                    'name': 'Категория A',
                    'storageDepth': 1095,
                    'templates': ['шаблон1', 'шаблон2']
                  }
                ]
              }
            ]
        parent_obj_uid (str, optional): UID родительского объекта. Если None — берётся из config.

    Returns:
        str: XML-строка с декларацией и отступами.

    Raises:
        Exception: При ошибках построения XML.
    """
    try:
        logging.info("Запуск формирования CIM-XML…")
        parent_obj_uid = parent_obj_uid or DEBUG_PARENT_UID

        # Корень RDF
        root = etree.Element('{%s}RDF' % NSMAP['rdf'], nsmap=NSMAP)

        # FullModel
        fm_uid = generate_uid()
        fm = etree.SubElement(
            root,
            '{%s}FullModel' % NSMAP['md'],
            attrib={"{%s}about" % NSMAP['rdf']: f"#_{fm_uid}"}
        )
        created = datetime.utcnow().strftime(MODEL_CREATED_FORMAT)
        etree.SubElement(fm, '{%s}Model.created' % NSMAP['md']).text = created
        etree.SubElement(fm, '{%s}Model.version' %
                         NSMAP['md']).text = "ver:sample;"
        etree.SubElement(fm, '{%s}Model.name' % NSMAP['me']).text = "CIM16"

        count_types = 0
        count_cats = 0
        count_templates = 0

        for idx, ct in enumerate(category_types):
            ct_name = ct.get('name', 'Unnamed Type').strip()
            if not ct_name or ct_name.lower() == 'nan':
                logging.warning("Пропущен тип категории с пустым именем")
                continue

            ct_uid = generate_uid()
            ct_elem = etree.SubElement(
                root,
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
                attrib={'{%s}resource' % NSMAP['rdf']: f"#_{parent_obj_uid}"}
            )
            etree.SubElement(
                ct_elem,
                '{%s}DjCategoryType.order' % NSMAP['me']
            ).text = str(idx)  # или DEFAULT_ORDER

            # Генерируем UID для категорий один раз
            categories = ct.get('categories', [])
            cat_uids = [generate_uid() for _ in categories]

            # Связь: DjCategoryType → DjCategory (через Categories)
            for uid in cat_uids:
                etree.SubElement(
                    ct_elem,
                    '{%s}DjCategoryType.Categories' % NSMAP['me'],
                    attrib={'{%s}resource' % NSMAP['rdf']: f"#_{uid}"}
                )

            # Создаём объекты DjCategory
            for idx, cat in enumerate(categories):
                cat_name = cat.get('name', 'Unnamed Category').strip()
                if not cat_name or cat_name.lower() == 'nan':
                    logging.warning(
                        f"Пропущена категория без имени в типе '{ct_name}'")
                    continue

                cat_uid = cat_uids[idx]
                cat_elem = etree.SubElement(
                    root,
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
                ).text = str(idx)  # или DEFAULT_ORDER
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
                ).text = str(cat.get('storageDepth', DEFAULT_STORAGE_DEPTH))

                # Связь: DjCategory → DjCategoryType
                etree.SubElement(
                    cat_elem,
                    '{%s}DjCategory.CategoryType' % NSMAP['me'],
                    attrib={'{%s}resource' % NSMAP['rdf']: f"#_{ct_uid}"}
                )

                # Шаблоны
                templates = cat.get('templates', [])
                for idx, tmpl_text in enumerate(templates):
                    if not tmpl_text or str(tmpl_text).lower() in ('nan', 'none', ''):
                        continue

                    tmpl_uid = generate_uid()
                    # Связь: DjCategory → DjRecordTemplate
                    etree.SubElement(
                        cat_elem,
                        '{%s}DjCategory.Templates' % NSMAP['me'],
                        attrib={'{%s}resource' % NSMAP['rdf']: f"#_{tmpl_uid}"}
                    )
                    tmpl_elem = etree.SubElement(
                        root,
                        '{%s}DjRecordTemplate' % NSMAP['me'],
                        attrib={"{%s}about" % NSMAP['rdf']: f"#_{tmpl_uid}"}
                    )
                    etree.SubElement(
                        tmpl_elem,
                        '{%s}DjRecordTemplate.order' % NSMAP['me']
                    ).text = str(idx)  # или DEFAULT_ORDER
                    etree.SubElement(
                        tmpl_elem,
                        '{%s}DjRecordTemplate.text' % NSMAP['me']
                    ).text = str(tmpl_text)
                    etree.SubElement(
                        tmpl_elem,
                        '{%s}DjRecordTemplate.Category' % NSMAP['me'],
                        attrib={'{%s}resource' % NSMAP['rdf']: f"#_{cat_uid}"}
                    )
                    count_templates += 1

                count_cats += 1
            count_types += 1

        logging.info(
            f"Сформировано: {count_types} типов, {count_cats} категорий, {count_templates} шаблонов"
        )
        logging.info("CIM-XML построен успешно.")

        return etree.tostring(
            root,
            encoding='utf-8',
            pretty_print=True,
            xml_declaration=True
        ).decode('utf-8')

    except Exception as e:
        logging.error(f"Ошибка при формировании CIM-XML: {e}")
        raise


# ———————————————————————————————————
# 🧪 Точка входа для отладки
# ———————————————————————————————————
if __name__ == "__main__":
    import sys
    from xlsx_parser import build_structure_from_excel
    from config import *
    import os

    # Настройка логирования
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler("cim_xml_debug.log", encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    try:
        from config import DEBUG_PARENT_UID
        filename = DEBUG_FILE_NAME
        base = os.path.splitext(os.path.basename(filename))[0]

        structure = build_structure_from_excel(filename)
        xml_str = create_cim_xml(structure, parent_obj_uid=DEBUG_PARENT_UID)
        with open(f"{base}.xml", "w", encoding="utf-8") as f:
            f.write(xml_str)
        logging.info("XML успешно сохранён в output.xml")
        print("✅ XML сформирован и сохранён в output.xml")
    except Exception as err:
        logging.error(f"Ошибка при выполнении __main__: {err}")
        print(f"[ERROR] {err}")
        sys.exit(1)
