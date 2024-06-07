"""
Module for creating and reading configuration files.

Functions:
    create_config()
    read_config() -> dictionary

Creates and reads a file called config.ini in the same directory as the module.
"""

# configparser module is part of the Python standard library
import configparser


def create_config():
    """Create and save the default configuration file if it does not exist."""
    config = configparser.ConfigParser()

    # Add default config values
    config["Import"] = {"data_directory": "", "import_file_path": ""}
    config["Export"] = {"export_directory": "",
                        "export_csv": True,
                        "csv_export_filename": "onbase_data.csv",
                        "export_excel": True,
                        "excel_export_filename": "onbase_data.xlsx"}
    config["Combine"] = {"combine_files": True,
                         "combined_file_directory": "",
                         "combine_skip_existing": False}

    # Write the configuration to a file if it does not exist yet
    try:
        with open('config.ini', 'x') as configfile:
            config.write(configfile)
    except FileExistsError:
        pass


def read_config():
    """Read and parse the configuration file. Return a dictionary of configuration values."""
    config = configparser.ConfigParser()
    config.read("config.ini")

    # Read config values from file
    config_values = {
        "data_directory": config.get("Import", "data_directory", raw=True),
        "import_file_path": config.get("Import", "import_file_path", raw=True),
        "export_directory": config.get("Export", "export_directory", raw=True),
        "export_csv": config.getboolean("Export", "export_csv"),
        "csv_export_filename": config.get("Export", "csv_export_filename"),
        "export_excel": config.getboolean("Export", "export_excel"),
        "excel_export_filename": config.get("Export", "excel_export_filename"),
        "combine_files": config.getboolean("Combine", "combine_files"),
        "combined_file_directory": config.get("Combine", "combined_file_directory", raw=True),
        "combine_skip_existing": config.getboolean("Combine", "combine_skip_existing")
    }

    return config_values
