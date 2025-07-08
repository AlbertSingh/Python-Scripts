import os
import xml.etree.ElementTree as ET
import pandas as pd


def extract_from_xml(xml_file, data_mapping):
    """
    Extract data from a single XML file using a specified XPath mapping.

    Parameters:
        xml_file (str): Path to the XML file.
        data_mapping (dict): Mapping of column names to XPath expressions.

    Returns:
        dict: Extracted data as a dictionary, or None on parse error.
    """
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        record = {}

        for column_name, xml_path in data_mapping.items():
            element = root.find(xml_path)
            record[column_name] = element.text if element is not None else None

        return record

    except ET.ParseError as e:
        print(f"[ERROR] Failed to parse XML file '{xml_file}': {e}")
        return None


def extract_from_met(met_file, data_mapping):
    """
    Extract data from a .met file using a key-value mapping.

    Parameters:
        met_file (str): Path to the .met file.
        data_mapping (dict): Mapping of column names to keys in the .met file.

    Returns:
        dict: Extracted data as a dictionary, or None on read error.
    """
    try:
        with open(met_file, 'r') as f:
            lines = f.readlines()

        file_data = {}
        for line in lines:
            if ':' in line:
                key, value = line.split(':', 1)
                file_data[key.strip()] = value.strip()

        record = {
            column_name: file_data.get(met_key)
            for column_name, met_key in data_mapping.items()
        }

        return record

    except IOError as e:
        print(f"[ERROR] Failed to read .met file '{met_file}': {e}")
        return None


def process_folder(folder_path, xml_mapping, met_mapping, output_excel_path):
    """
    Process all XML and .met files in a folder, extract data using mappings,
    and write results to an Excel file.

    Parameters:
        folder_path (str): Directory containing input files.
        xml_mapping (dict): Mapping for extracting data from XML files.
        met_mapping (dict): Mapping for extracting data from .met files.
        output_excel_path (str): Path to output Excel file.
    """
    all_records = []
    print(f"\n[INFO] Scanning folder: {folder_path}")

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        record = None

        if filename.lower().endswith(".xml"):
            print(f"  -> Extracting XML: {filename}")
            record = extract_from_xml(file_path, xml_mapping)

        # elif filename.lower().endswith(".met"):
        #     print(f"  -> Extracting MET: {filename}")
        #     record = extract_from_met(file_path, met_mapping)

        if record:
            all_records.append(record)

    if all_records:
        all_columns = list(dict.fromkeys(list(xml_mapping.keys()) + list(met_mapping.keys())))
        df = pd.DataFrame(all_records, columns=all_columns)
        df.to_excel(output_excel_path, index=False)
        print(f"\n[SUCCESS] Data exported to '{output_excel_path}'")
    else:
        print("\n[WARNING] No valid data extracted from the provided folder.")


if __name__ == "__main__":
    xml_mapping = {
        "Full Name": "./details/name",
        "Age": "./details/age",
        "Email Address": "./contact/email"
    }

    met_mapping = {
        "Full Name": "Full Name",
        "Age": "Age",
        "Email Address": "Email Address"
    }

    input_folder = r"path/to/your/files"  # <-- Update this path
    output_file = "output_data.xlsx"

    process_folder(input_folder, xml_mapping, met_mapping, output_file)
