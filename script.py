import zipfile
import os
import shutil
import tempfile
from lxml import etree as ET

def modify_cells(xlsm_file, sheet_xml, updates):
    try:
        # Crear un directorio temporal
        temp_dir = tempfile.mkdtemp()
        compression_info = {}

        # Extraer el archivo .xlsm preservando la compresión
        with zipfile.ZipFile(xlsm_file, 'r') as zf:
            for info in zf.infolist():
                compression_info[info.filename] = info.compress_type
            zf.extractall(temp_dir)

        # Modificar sharedStrings.xml si es necesario
        shared_strings_path = os.path.join(temp_dir, "xl", "sharedStrings.xml")
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

        # Cargar sharedStrings.xml si existe
        if os.path.exists(shared_strings_path):
            parser = ET.XMLParser(remove_blank_text=True)
            ss_tree = ET.parse(shared_strings_path, parser)
            ss_root = ss_tree.getroot()
            si_elements = ss_root.findall(f"{namespace}si")
        else:
            ss_root = ET.Element(f"{namespace}sst", count="0", uniqueCount="0")
            ss_tree = ET.ElementTree(ss_root)
            si_elements = []

        # Modificar la hoja de cálculo
        sheet_path = os.path.join(temp_dir, "xl", "worksheets", f"{sheet_xml}.xml")
        parser = ET.XMLParser(remove_blank_text=True)
        tree = ET.parse(sheet_path, parser)
        root = tree.getroot()
        sheet_data = root.find(f"{namespace}sheetData")

        for row, (start_col, values) in updates.items():
            start_col_letter = start_col.upper()

            for idx, value in enumerate(values):
                current_col = chr(ord(start_col_letter) + idx)
                cell_ref = f"{current_col}{row}"

                # Buscar la fila existente o crearla
                row_elem = sheet_data.find(f'.//{namespace}row[@r="{row}"]')
                if row_elem is None:
                    row_elem = ET.SubElement(sheet_data, f"{namespace}row", r=str(row))

                # Buscar la celda existente o crear una nueva
                cell_element = row_elem.find(f'.//{namespace}c[@r="{cell_ref}"]')
                if cell_element is None:
                    cell_element = ET.SubElement(row_elem, f'{namespace}c', r=cell_ref)

                # Determinar tipo de dato
                try:
                    float(value.replace(',', '.'))
                    cell_element.set("t", "n")  # Número
                    cell_value = ET.SubElement(cell_element, f"{namespace}v")
                    cell_value.text = str(value)
                except ValueError:
                    # Si es texto, usar sharedStrings.xml
                    cell_element.set("t", "s")

                    # Verificar si el valor ya existe en sharedStrings.xml
                    existing_index = next((i for i, si in enumerate(si_elements) if si.find(f"{namespace}t") is not None and si.find(f"{namespace}t").text == value), None)

                    if existing_index is None:
                        si = ET.SubElement(ss_root, f"{namespace}si")
                        ET.SubElement(si, f"{namespace}t").text = str(value)
                        si_elements.append(si)
                        ss_index = len(si_elements) - 1
                    else:
                        ss_index = existing_index

                    cell_value = ET.SubElement(cell_element, f"{namespace}v")
                    cell_value.text = str(ss_index)

        # Guardar cambios en sharedStrings.xml
        ss_root.set("count", str(len(si_elements)))
        ss_root.set("uniqueCount", str(len(set(si.find(f"{namespace}t").text for si in si_elements if si.find(f"{namespace}t") is not None))))
        with open(shared_strings_path, "wb") as f:
            ss_tree.write(f, encoding="UTF-8", xml_declaration=True, pretty_print=True)

        # Guardar la hoja modificada
        tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)

        # Reconstruir el archivo .xlsm sin modificar certificados de macros
        new_file = os.path.join(".", "Final_" + os.path.basename(xlsm_file))
        with zipfile.ZipFile(new_file, 'w') as new_zip:
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    abs_path = os.path.join(root_dir, file)
                    rel_path = os.path.relpath(abs_path, temp_dir).replace(os.sep, '/')
                    compress_type = compression_info.get(rel_path, zipfile.ZIP_DEFLATED)

                    # Evitar compresión en archivos binarios (especialmente macros)
                    if "vbaProject.bin" in rel_path or rel_path.endswith(".bin"):
                        compress_type = zipfile.ZIP_STORED
                    
                    new_zip.write(abs_path, rel_path, compress_type=compress_type)

        print(f"✅ Celdas actualizadas con éxito!\nArchivo guardado: {new_file}")

    except Exception as e:
        print(f"❌ Error: {str(e)}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    # Datos a insertar
    data_to_insert = {
        5: ["C", [
            "20", "HYPOTHEQUE CONVENTION", "29/12/2011", "05/01/2028",
            "CREDIT DU NORD (456 504 851)", "SCI DES 9 FILS (569 208 825)", "AQ14",
            "*", "1, 2, 38, 107", "1.000.000,00", "200.000,00", "EUR",
            "3,21000%", "NOT CLERC / NEUILLY SUR SEINE", "27/01/2012", "2012V",
            "183", 'Bordereau rectificatif', "NON", "*", "*", "*", "*", "*", "*", "*"
        ]],
        7: ["C", [
            "5", "PRIVILEGE DE PRETEUR DE DENIERS", "26/04/2002", "05/05/2029",
            "ABBEY NATIONAL FRANCE", "ROS (01/09/1968)", "AQ14", "*",
            "9, 252", "447.056,74", "89.411,34", "EUR", "4,96700%",
            "NOT GUERIN-BERTRAND-GREMONT / PARIS", "25/06/2002", "2002V",
            "1139", "Regularisation suite", "TOTALE", "25/09/2007",
            "01/10/2007", "B214P01 2007D8342", "AQ13", "*", "9, 252", "*"
        ]]
    }

    original_file = "template_with_signed_macro.xlsm"
    sheet_to_modify = "sheet1"

    modify_cells(original_file, sheet_to_modify, data_to_insert)

