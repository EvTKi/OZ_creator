import uuid
from lxml import etree
import logging
import os
from logging_config import setup_logging
from debug_config import *
NSMAP = {
    'rdf': 'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
    'md': 'http://iec.ch/TC57/61970-552/ModelDescription/1#',
    'cim': 'http://iec.ch/TC57/2014/CIM-schema-cim16#',
    'me': 'http://monitel.com/2014/schema-cim16#',
}


def rdf_resource(uid):
    return ("{%s}resource" % NSMAP['rdf'], f"#_{uid}")


def gen_uid():
    return str(uuid.uuid4())


def create_cim_xml(category_types, parent_obj_uid=None):
    try:
        logging.info("Запуск формирования CIM-XML…")
        parent_obj_uid = parent_obj_uid if parent_obj_uid is not None else DEBUG_PARENT_UID

        root = etree.Element('{%s}RDF' % NSMAP['rdf'], nsmap=NSMAP)
        fm_uid = gen_uid()
        fm = etree.SubElement(root, '{%s}FullModel' % NSMAP['md'],
                              attrib={"{%s}about" % NSMAP['rdf']: f"#_{fm_uid}"})
        etree.SubElement(fm, '{%s}Model.created' %
                         NSMAP['md']).text = "2025-08-06T06:45:42.000Z"
        etree.SubElement(fm, '{%s}Model.version' %
                         NSMAP['md']).text = "ver:sample;"
        etree.SubElement(fm, '{%s}Model.name' % NSMAP['me']).text = "CIM16"

        count_types = 0
        count_cats = 0
        count_templates = 0

        for ct in category_types:
            ct_uid = gen_uid()
            ct_elem = etree.SubElement(root, '{%s}DjCategoryType' % NSMAP['me'],
                                       attrib={"{%s}about" % NSMAP['rdf']: f"#_{ct_uid}"})
            etree.SubElement(
                ct_elem, '{%s}IdentifiedObject.name' % NSMAP['cim']).text = ct['name']
            etree.SubElement(ct_elem, '{%s}IdentifiedObject.ParentObject' % NSMAP['me'],
                             attrib=dict([rdf_resource(parent_obj_uid)]))
            etree.SubElement(
                ct_elem, '{%s}DjCategoryType.order' % NSMAP['me']).text = "0"
            child_cat_uids = []
            for cat in ct.get('categories', []):
                child_cat_uid = gen_uid()
                child_cat_uids.append(child_cat_uid)
                etree.SubElement(ct_elem, '{%s}IdentifiedObject.ChildObjects' % NSMAP['me'],
                                 attrib=dict([rdf_resource(child_cat_uid)]))
            for cat in ct.get('categories', []):
                child_cat_uid = gen_uid()
                child_cat_uids.append(child_cat_uid)
                etree.SubElement(ct_elem, '{%s}DjCategoryType.Categories' % NSMAP['me'],
                                 attrib=dict([rdf_resource(child_cat_uid)]))
            count_types += 1
            for idx, cat in enumerate(ct.get('categories', [])):
                cat_uid = child_cat_uids[idx]
                cat_elem = etree.SubElement(root, '{%s}DjCategory' % NSMAP['me'],
                                            attrib={"{%s}about" % NSMAP['rdf']: f"#_{cat_uid}"})
                etree.SubElement(
                    cat_elem, '{%s}IdentifiedObject.name' % NSMAP['cim']).text = cat['name']
                etree.SubElement(cat_elem, '{%s}IdentifiedObject.ParentObject' % NSMAP['me'],
                                 attrib=dict([rdf_resource(ct_uid)]))
                etree.SubElement(
                    cat_elem, '{%s}DjCategory.actionTimeWithinShift' % NSMAP['me']).text = "false"
                etree.SubElement(
                    cat_elem, '{%s}DjCategory.messageRequired' % NSMAP['me']).text = "false"
                etree.SubElement(
                    cat_elem, '{%s}DjCategory.order' % NSMAP['me']).text = "0"
                etree.SubElement(
                    cat_elem, '{%s}DjCategory.reportForHigherOperationalStaff' % NSMAP['me']).text = "false"
                etree.SubElement(
                    cat_elem, '{%s}DjCategory.shiftPersonRestricted' % NSMAP['me']).text = "false"
                etree.SubElement(
                    cat_elem, '{%s}DjCategory.storageDepth' % NSMAP['me']).text = "1095"
                etree.SubElement(cat_elem, '{%s}DjCategory.CategoryType' % NSMAP['me'],
                                 attrib=dict([rdf_resource(ct_uid)]))
                count_cats += 1
                for tmpl_text in cat.get('templates', []):
                    tmpl_uid = gen_uid()
                    etree.SubElement(cat_elem, '{%s}DjCategory.Templates' % NSMAP['me'],
                                     attrib=dict([rdf_resource(tmpl_uid)]))
                    tmpl_elem = etree.SubElement(root, '{%s}DjRecordTemplate' % NSMAP['me'],
                                                 attrib={"{%s}about" % NSMAP['rdf']: f"#_{tmpl_uid}"})
                    etree.SubElement(
                        tmpl_elem, '{%s}DjRecordTemplate.order' % NSMAP['me']).text = "0"
                    etree.SubElement(
                        tmpl_elem, '{%s}DjRecordTemplate.text' % NSMAP['me']).text = tmpl_text
                    etree.SubElement(tmpl_elem, '{%s}DjRecordTemplate.Category' % NSMAP['me'],
                                     attrib=dict([rdf_resource(cat_uid)]))
                    count_templates += 1

        logging.info(
            f"Сформировано: {count_types} типов, {count_cats} категорий, {count_templates} шаблонов")
        logging.info("CIM-XML построен успешно.")
        return etree.tostring(root, encoding='utf-8', pretty_print=True, xml_declaration=True).decode()
    except Exception as e:
        logging.error(f"Ошибка при формировании CIM-XML: {e}")
        raise


if __name__ == "__main__":
    from xlsx_parser import build_structure_from_excel
    import sys

    filename = DEBUG_FILE_NAME
    base = os.path.splitext(os.path.basename(filename))[0]
    setup_logging(filename)
    try:
        structure = build_structure_from_excel(filename)
        xml_str = create_cim_xml(structure)
        with open(f'{base}.xml', "w", encoding="utf-8") as f:
            f.write(xml_str)
            logging.info("XML успешно сохранён в categories.xml")
        print("XML сформирован и сохранён.")
    except Exception as err:
        logging.error(f"Ошибка в основном запуске: {err}")
        print(f"[ERROR] {err}")
        sys.exit(1)
