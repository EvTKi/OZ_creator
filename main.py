# main.py
from config import *
from xlsx_parser import build_structure_from_excel
from cim_xml_creator import create_cim_xml
import os
if __name__ == "__main__":
    xlsx_path = DEBUG_FILE_NAME
    base = os.path.splitext(os.path.basename(xlsx_path))[0]
    out_xml_path = f'{base}.xml'
    structure = build_structure_from_excel(xlsx_path)
    your_parent_uid = DEBUG_PARENT_UID
    xml = create_cim_xml(structure, parent_obj_uid=your_parent_uid)
    with open(out_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml)
    print("XML успешно сформирован:", out_xml_path)
