"""
Module for importing data from the disk dump and exporting to CSV and Excel files.

Functions:
    import_data(string, list of strings, list of strings) -> list of dictionaries
    export_csv(list of dictionaries, string, list of strings)
    export_excel(list of dictionaries, string, list of strings)
"""

import csv
import xlsxwriter


def import_data(import_filepath, document_attributes, file_attributes):
    """Import file metadata from OnBase data dump. Return a list of dictionaries containing document metadata.

    Parameters:
        import_filepath (string): Path to the OnBase data dump text file
        document_attributes (list): List of strings containing the names of document attributes
        file_attributes (list): List of strings containing the names of file attributes

    Returns:
        output_data (list): List of dictionaries containing document metadata
                            Each dictionary corresponds to one document.
                            Each dictionary contains the values for each document attribute
                                and a list of all files in the document.

    Raises:
        ValueError: Invalid attribute. This occurs if an attribute is found that is not in the document_attributes
                    or file_attributes lists.
    """

    print("Importing data from file...")
    # Create a list to hold document metadata
    output_data = []
    # Create a set to hold document handles (makes checking for duplicates faster)
    doc_handles = set()

    # Read data from file
    with open(import_filepath, mode='r') as input_file:
        input_data = input_file.read()

    # Remove first line and end section
    input_data = input_data.split("BEGIN:", 1)[1]
    input_data = input_data.split("END:", 1)[0]

    # Split the text by BEGIN:
    sections = input_data.split("BEGIN:")

    # Loop over the sections
    for section in sections:
        # Skip empty documents
        if section == "":
            continue

        # Split the section by newline
        lines = section.split("\n")

        # Create a dictionary to hold the document attributes and files
        doc = {"Files": []}

        # Loop over the lines
        for line in lines:
            # Skip empty lines
            if line == "":
                continue

            # Split the line by colon
            parts = line.split(":")

            # Get the attribute name and title
            attribute = parts[0]
            value = parts[1]

            # Trim the spaces
            attribute = attribute.strip('>')
            value = value.strip()

            # If the attribute is a document attribute, add to dictionary
            if attribute in document_attributes:
                doc[attribute] = value

            # If the attribute is a filename, add it to the Files list
            elif attribute == "FileName":
                if value not in doc.get("Files"):
                    doc.get("Files").append(value)

            elif attribute in file_attributes:
                pass

            else:
                raise ValueError(f"Invalid attribute: {attribute}")

        # Check if any document with the same handle exists. If so, append a number to it
        if doc.get("Document Handle") in doc_handles:
            counter = 2
            new_handle = doc.get("Document Handle") + "_" + str(counter)
            while new_handle in doc_handles:
                new_handle = doc.get("Document Handle") + "_" + str(counter)
            doc["Document Handle"] = new_handle

        # Append document to output data
        output_data.append(doc)
        doc_handles.add(doc.get("Document Handle"))

    print("Data imported.")
    return output_data


def export_csv(output_data, output_csv_filepath, document_attributes):
    """Export data to a CSV file.

    Parameters:
        output_data (list): List of dictionaries containing document metadata
                            Each dictionary corresponds to one document.
                            Each dictionary contains the values for each document attribute and a path to the file.
        output_csv_filepath (raw string): Path to the CSV file to be created
        document_attributes (list): List of strings containing the names of document attributes
    """

    print("Writing data to CSV file...")
    try:
        with open(output_csv_filepath, mode='w') as output_file:
            output_writer = csv.DictWriter(f=output_file, fieldnames=document_attributes, restval="",
                                           extrasaction="ignore", dialect="excel", lineterminator='\n')
            output_writer.writeheader()
            output_writer.writerows(output_data)
    except Exception as error:
        print(f"Error: Could not export to CSV. {error}")
    else:
        print("Data written to CSV file.")


def export_excel(output_data, output_excel_filepath, document_attributes):
    """Export data to an Excel file.

    Parameters:
        output_data (list): List of dictionaries containing document metadata
                            Each dictionary corresponds to one document.
                            Each dictionary contains the values for each document attribute and a path to the file.
        output_excel_filepath (raw string): Path to the Excel file to be created
        document_attributes (list): List of strings containing the names of document attributes
    """
    print("Writing data to Excel file...")
    try:
        workbook = xlsxwriter.Workbook(output_excel_filepath)
        worksheet = workbook.add_worksheet()

        # Write column headers
        for pos, attr in enumerate(document_attributes):
            worksheet.write(0, pos, attr)

        # Write document data
        for row, doc in enumerate(output_data, start=1):
            for col, attr in enumerate(document_attributes):
                if attr == "File Link" and doc.get(attr) is not None:
                    worksheet.write_url(row, col, doc.get(attr))
                else:
                    worksheet.write(row, col, doc.get(attr))

        workbook.close()
    except Exception as error:
        print(f"Error: Could not export to Excel. {error}")
    else:
        print("Data written to Excel file.")
