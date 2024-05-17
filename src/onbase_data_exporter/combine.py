"""
combine.py
Written by Sophie Garrett
Functions for combining and renaming files.
"""

import os
import shutil
from PIL import Image


def combine_files(data, input_directory, output_directory):
    print("Combining files...")

    # NOTE: file extensions present:
    # ['pdf', 'PDF', 'jpeg', 'tiff', 'png', 'tif', 'ctx', 'TIF', 'csv', 'pub', 'xls', 'jpg',
    # 'doc', 'ppt', 'xlsx', 'docx', 'gif', 'pptx', 'txt']

    # File extensions with multiple files per document:
    # ['jpeg', 'tiff', 'png', 'tif', 'TIF']

    # Loop over each document
    for doc in data:
        if len(doc.get("Files")) == 0:
            raise ValueError("No files associated with document.")

        elif len(doc.get("Files")) == 1:
            volume = doc["Files"][0].split('\\')[1]

            # FOR NOW, only working with V10-16 as test data
            # if volume in {"V10", "V11", "V12", "V13", "V14", "V15", "V16"}:
            if volume in {"V10"}:
                old_filepath = os.path.join(input_directory, doc["Files"][0].split('\\', 1)[1])
                file_ext = doc["Files"][0].split('.')[1]
                new_filepath = os.path.join(output_directory, doc.get("Document Handle") + '.' + file_ext)
                try:
                    shutil.copy2(str(old_filepath), str(new_filepath))
                except FileNotFoundError as error:
                    print(f"Error: {error}")
                else:
                    # Save new filepath
                    doc["File Link"] = new_filepath

        else:
            output_volume = doc["Files"][0].split('\\')[1]

            # FOR NOW, only working with V10-16 as test data
            # if output_volume in {"V10", "V11", "V12", "V13", "V14", "V15", "V16"}:
            if output_volume in {"V13"}:

                images = []
                for file in doc.get("Files"):
                    volume = file.split('\\')[1]
                    file_ext = file.split('.')[1]

                    # FOR NOW, only working with V10-16 as test data
                    # if volume in {"V10", "V11", "V12", "V13", "V14", "V15", "V16"}:
                    if output_volume in {"V13"}:
                        old_filepath = os.path.join(input_directory, file.split('\\', 1)[1])
                        try:
                            img = Image.open(old_filepath)
                        except FileNotFoundError as error:
                            print(f"Error: {error}")
                        else:
                            images.append(img)

                new_filepath = os.path.join(output_directory, doc.get("Document Handle") + '.pdf')
                try:
                    images[0].save(new_filepath, save_all="True", append_images=images[1:])
                except IndexError:
                    print(f"Error: List index out of range for {doc.get("Document Handle")}.")
                else:
                    # Save new filepath
                    doc["File Link"] = new_filepath

    print("Files combined successfully.")
