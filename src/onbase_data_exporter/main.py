"""OnBase Data Exporter

A Python program that translates OnBase data dumps into a usable format.
"""

import os.path
from config import create_config, read_config
from readwrite import import_data, export_csv, export_excel
from combine import combine_files


def main():
    """Main function.

    Raises:
        NotADirectoryError: One of the required directories (data directory, export directory, or combined files
                            directory) does not exist.
    """

    print("Starting OnBase data export...")

    # Read configuration file
    create_config()
    config_data = read_config()

    # Check that required directories exist
    if not os.path.isdir(config_data.get("data_directory")):
        raise NotADirectoryError("Data directory does not exist.")
    if not os.path.isdir(config_data.get("export_directory")):
        raise NotADirectoryError("Export directory does not exist.")
    if config_data.get("combine_files") and not os.path.isdir(config_data.get("combined_file_directory")):
        raise NotADirectoryError("Combined file directory does not exist.")

    # Determine file paths from config values
    csv_export_filepath = os.path.join(config_data.get("export_directory"), config_data.get("csv_export_filename"))
    excel_export_filepath = os.path.join(config_data.get("export_directory"), config_data.get("excel_export_filename"))

    document_attributes = ["DocTypeName", "DocDate", "Fiscal Year", "Provider Name", "Program Name", "Department",
                           "Description", "Document Section", "Doc Handle Link", "Document Handle", "File Link"]

    file_attributes = ["DiskgroupNum", "VolumeNum", "FileSize", "NumOfPages", "DocRevNum", "Rendition",
                       "PhysicalPageNum", "ItemPageNum", "FileTypeNum", "ImageType", "Compress", "Xdpi", "Ydpi",
                       "TextEncoding", "FileName"]

    # Import metadata
    document_metadata = import_data(config_data.get("import_file_path"), document_attributes, file_attributes)

    # Combine files
    if config_data.get("combine_files"):
        combine_files(document_metadata, config_data.get("data_directory"),
                      config_data.get("combined_file_directory"), config_data.get("combine_skip_existing"))

    # Write output to CSV file
    if config_data.get("export_csv"):
        export_csv(document_metadata, csv_export_filepath, document_attributes)

    # Write output to Excel file
    if config_data.get("export_excel"):
        export_excel(document_metadata, excel_export_filepath, document_attributes)

    print("OnBase data export complete.")


if __name__ == "__main__":
    main()
