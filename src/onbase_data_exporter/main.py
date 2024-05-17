"""
Written by Sophie Garrett
Program to translate OnBase data dump into usable form.
"""

import os.path
from config import create_config, read_config
from readwrite import import_data, export_csv, export_excel
from combine import combine_files


def main():
    print("Starting OnBase data conversion...")

    # Read configuration file
    create_config()
    config_data = read_config()

    # Determine file paths from config values
    import_filepath = os.path.join(config_data.get("data_directory"), config_data.get("import_filename"))
    csv_export_filepath = os.path.join(config_data.get("data_directory"), config_data.get("csv_export_filename"))
    excel_export_filepath = os.path.join(config_data.get("data_directory"), config_data.get("excel_export_filename"))

    document_attributes = ["DocTypeName", "DocDate", "Fiscal Year", "Provider Name", "Program Name", "Department",
                           "Description", "Document Section", "Doc Handle Link", "Document Handle"]

    file_attributes = ["DiskgroupNum", "VolumeNum", "FileSize", "NumOfPages", "DocRevNum", "Rendition", "PhysicalPageNum",
                       "ItemPageNum", "FileTypeNum", "ImageType", "Compress", "Xdpi", "Ydpi", "TextEncoding", "FileName"]

    # Import metadata
    document_metadata = import_data(import_filepath, document_attributes, file_attributes)

    # Combine files
    if config_data.get("combine_files"):
        combine_files(document_metadata, config_data.get("data_directory"), config_data.get("combined_file_directory"))

    # Write output to CSV file
    if config_data.get("export_csv"):
        export_csv(document_metadata, csv_export_filepath, document_attributes)

    # Write output to Excel file
    if config_data.get("export_excel"):
        export_excel(document_metadata, excel_export_filepath, document_attributes)

    print("OnBase data conversion complete.")


if __name__ == "__main__":
    main()
