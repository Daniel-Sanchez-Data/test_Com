import zipfile
import os
import shutil
import tempfile
from lxml import etree as ET

def modify_cells(xlsm_file, sheet_xml, updates):
    try:
        # Create a temporary directory
        temp_dir = tempfile.mkdtemp()
        compression_info = {}

        # 1. Extract the .xlsm file while preserving compression
        with zipfile.ZipFile(xlsm_file, 'r') as zf:
            for info in zf.infolist():
                compression_info[info.filename] = info.compress_type
            zf.extractall(temp_dir)

        # 2. Modify sharedStrings.xml if necessary
        shared_strings_path = os.path.join(temp_dir, "xl", "sharedStrings.xml")
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

        # Load sharedStrings.xml
        if os.path.exists(shared_strings_path):
            parser = ET.XMLParser(remove_blank_text=True)
            ss_tree = ET.parse(shared_strings_path, parser)
            ss_root = ss_tree.getroot()
            si_elements = ss_root.findall(f"{namespace}si")
        else:
            ss_root = ET.Element(f"{namespace}sst", count="0", uniqueCount="0")
            ss_tree = ET.ElementTree(ss_root)
            si_elements = []

        # 3. Modify the target sheet
        sheet_path = os.path.join(temp_dir, "xl", "worksheets", f"{sheet_xml}.xml")
        parser = ET.XMLParser(remove_blank_text=True)
        tree = ET.parse(sheet_path, parser)
        root = tree.getroot()
        sheet_data = root.find(f"{namespace}sheetData")

        for row, (start_col, values) in updates.items():
            # Calculate the starting column letter
            start_col_letter = start_col.upper()

            for idx, value in enumerate(values):
                # Calculate the current column letter
                current_col = chr(ord(start_col_letter) + idx)
                cell_ref = f"{current_col}{row}"

                # Find existing cell
                cell_element = sheet_data.find(f'.//{namespace}c[@r="{cell_ref}"]')
                if cell_element is not None:
                    cell_type = cell_element.get("t", "n")
                    if cell_type == "s":
                        # Get shared string index
                        ss_index = int(cell_element.find(f"{namespace}v").text)
                        # Modify shared string
                        if ss_index < len(si_elements):
                            si_elements[ss_index].find(f"{namespace}t").text = str(value)
                        else:
                            # Add new shared string
                            si = ET.SubElement(ss_root, f"{namespace}si")
                            ET.SubElement(si, f"{namespace}t").text = str(value)
                            si_elements.append(si)
                            # Update cell's shared string index
                            cell_element.find(f"{namespace}v").text = str(len(si_elements) - 1)
                    else:
                        # Handle other types (inlineStr, number)
                        v_elem = cell_element.find(f"{namespace}v")
                        if v_elem is not None:
                            v_elem.text = str(value)
                        else:
                            ET.SubElement(cell_element, f"{namespace}v").text = str(value)
                else:
                    # Create a new cell
                    # Determine data type (assume text; can be improved by detecting types)
                    cell_type = "s"

                    # If the value is numeric, adjust the type
                    try:
                        float(value.replace(',', '.'))
                        cell_type = "n"
                    except ValueError:
                        pass  # Keep type "s"

                    if cell_type == "s":
                        # Handle shared strings
                        si = ET.SubElement(ss_root, f"{namespace}si")
                        ET.SubElement(si, f"{namespace}t").text = str(value)
                        ss_index = len(si_elements)
                        si_elements.append(si)

                        # Add cell to the sheet
                        row_num = str(row)
                        row_elem = sheet_data.find(f'{namespace}row[@r="{row_num}"]')
                        if row_elem is None:
                            row_elem = ET.SubElement(sheet_data, f'{namespace}row', r=row_num)
                        new_cell = ET.SubElement(row_elem, f'{namespace}c', r=cell_ref, t="s")
                        ET.SubElement(new_cell, f"{namespace}v").text = str(ss_index)
                    else:
                        # Add numeric cell
                        row_num = str(row)
                        row_elem = sheet_data.find(f'{namespace}row[@r="{row_num}"]')
                        if row_elem is None:
                            row_elem = ET.SubElement(sheet_data, f'{namespace}row', r=row_num)
                        new_cell = ET.SubElement(row_elem, f'{namespace}c', r=cell_ref, t="n")
                        ET.SubElement(new_cell, f"{namespace}v").text = str(value)

        # 4. Save changes to sharedStrings.xml
        ss_root.set("count", str(len(si_elements)))
        ss_root.set("uniqueCount", str(len(si_elements)))
        with open(shared_strings_path, "wb") as f:
            ss_tree.write(f, encoding="UTF-8", xml_declaration=True, pretty_print=True)

        # 5. Save the modified sheet
        tree.write(
            sheet_path,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
            pretty_print=False
        )

        # 6. Rebuild the .xlsm file (in the current directory)
        new_file = os.path.join(".", "FINAL_" + os.path.basename(xlsm_file))
        with zipfile.ZipFile(new_file, 'w') as new_zip:
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    abs_path = os.path.join(root_dir, file)
                    rel_path = os.path.relpath(abs_path, temp_dir).replace(os.sep, '/')
                    compress_type = compression_info.get(rel_path, zipfile.ZIP_DEFLATED)
                    
                    # Force no compression for vbaProject.bin
                    if "vbaProject.bin" in rel_path:
                        compress_type = zipfile.ZIP_STORED
                    
                    new_zip.write(abs_path, rel_path, compress_type=compress_type)

        print(f"✅ Cells updated successfully!\nFile: {new_file}")

    except Exception as e:
        print(f"❌ Error: {str(e)}")
    finally:
        # Clean up the temporary directory
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    # Define the data to insert
    data_to_insert = {
        5: [  # Row 5
            "C",  # Starting column
            [
                "20",
                "HYPOTHEQUE CONVENTION",
                "29/12/2011",
                "05/01/2028",
                "CREDIT DU NORD (456 504 851)",
                "SCI DES 9 FILS (569 208 825)",
                "AQ14",
                "*",
                "1, 2, 38, 107",
                "1.000.000,00",
                "200.000,00",
                "EUR",
                "3,21000%",
                "NOT CLERC / NEUILLY SUR SEINE",
                "27/01/2012",
                "2012V",
                "183",
                'Bordereau rectificatif en ce qui concerne el párrafo "en virtud de"',
                "NON",
                "*",
                "*",
                "*",
                "*",
                "*",
                "*",
                "*"
            ]
        ],
        7: [  # Row 7
            "C",  # Starting column
            [
                "5",
                "PRIVILEGE DE PRETEUR DE DENIERS",
                "26/04/2002",
                "05/05/2029",
                "ABBEY NATIONAL FRANCE",
                "ROS (01/09/1968)",
                "AQ14",
                "*",
                "9, 252",
                "447.056,74",
                "89.411,34",
                "EUR",
                "4,96700%",
                "NOT GUERIN-BERTRAND-GREMONT / PARIS",
                "25/06/2002",
                "2002V",
                "1139",
                "Regularisation suite a publication du titre.",
                "TOTALE",
                "25/09/2007",
                "01/10/2007",
                "B214P01 2007D8342",
                "AQ13",
                "*",
                "9, 252",
                "*"
            ]
        ]
    }

    # Name of the uploaded .xlsm file in your Codespace
    original_file = "template_with_signed_macro.xlsm"
    
    # Name of the sheet you want to modify (ensure it matches the sheet name in Excel)
    sheet_to_modify = "sheet1"

    modify_cells(original_file, sheet_to_modify, data_to_insert)
