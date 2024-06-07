"""
Module for combining and renaming files.

Functions:
    combine_files(list of dictionaries, string, string, boolean)
"""

import os
import shutil
from PIL import Image


def combine_files(data, input_directory, output_directory, skip_existing):
    """Combine and rename document files.

    Documents containing only one file will be copied to the output directory and renamed to their document handle.
        The original file extension and most file metadata will be kept.
    Documents containing multiple files will be combined into single PDFs. These will be saved in the output directory.
        The file name will be {document handle}.pdf.

    Parameters:
        data (list): List of dictionaries containing document metadata
                     Each dictionary corresponds to one document.
                     Each dictionary contains the values for each document attribute
                     and a list of all files in the document.
        input_directory (string): Path to the directory containing input files
        output_directory (string): Path to the directory to save output (combined) files in
        skip_existing (bool): Setting for whether to skip existing files or not
                              If True, files that already exist in the output directory will be skipped.
                              If False, files that already exist in the output directory will be re-combined.
    """

    print("Combining files...")

    # Loop over each document
    for doc in data:
        try:
            if len(doc.get("Files")) == 0:
                raise ValueError("No files associated with document.")

            elif len(doc.get("Files")) == 1:
                old_filepath = os.path.join(input_directory, doc["Files"][0].split('\\', 1)[1])
                file_ext = doc["Files"][0].split('.')[1]
                new_filepath = os.path.join(output_directory, doc.get("Document Handle") + '.' + file_ext)

                # If skip_existing option is enabled, skip copying file if it already exists in the output directory
                if skip_existing and os.path.isfile(new_filepath):
                    # Save new filepath
                    doc["File Link"] = new_filepath
                else:
                    try:
                        shutil.copy2(str(old_filepath), str(new_filepath))
                    except FileNotFoundError as error:
                        print(f"Error: Could not copy document #{doc.get("Document Handle")}. {error}")
                    else:
                        # Save new filepath
                        doc["File Link"] = new_filepath

            else:
                new_filepath = os.path.join(output_directory, doc.get("Document Handle") + '.pdf')

                # If skip_existing option is enabled, skip copying file if it already exists in the output directory
                if skip_existing and os.path.isfile(new_filepath):
                    # Save new filepath
                    doc["File Link"] = new_filepath

                else:
                    images = []
                    for file in doc.get("Files"):
                        old_filepath = os.path.join(input_directory, file.split('\\', 1)[1])
                        try:
                            img = Image.open(old_filepath)
                        except FileNotFoundError as error:
                            print(f"Error: File not found for document #{doc.get("Document Handle")}. {error}")
                        else:
                            images.append(img)

                    try:
                        images[0].save(new_filepath, save_all="True", append_images=images[1:])
                    except IndexError as error:
                        print(f"Error: Index out of range for document #{doc.get("Document Handle")}. {error}")
                    except FileNotFoundError as error:
                        print(f"Error: Could not combine files for document #{doc.get("Document Handle")}. {error}")
                    else:
                        # Save new filepath
                        doc["File Link"] = new_filepath
        except Exception as error:
            print(f"An unexpected error occurred while processing document #{doc.get("Document Handle")}: {error}")

    print("Files combined successfully.")
