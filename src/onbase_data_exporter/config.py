# Configuration module

import configparser


def create_config():
    config = configparser.ConfigParser()

    # Add default config values
    config["General"] = {"data_directory": "",
                         "combine_files": True}
    config["Import"] = {"import_filename": ""}
    config["Export"] = {"export_csv": True,
                        "csv_export_filename": "output.csv",
                        "export_excel": True,
                        "excel_export_filename": "output.xlsx"}

    # Write the configuration to a file if it does not exist yet
    try:
        with open('config.ini', 'x') as configfile:
            config.write(configfile)
    except FileExistsError:
        pass


def read_config():
    config = configparser.ConfigParser()
    config.read("config.ini")

    # Read config values from file
    config_values = {
        "data_directory": config.get("General", "data_directory", raw=True),
        "combine_files": config.getboolean("General", "combine_files"),
        "import_filename": config.get("Import", "import_filename"),
        "export_csv": config.getboolean("Export", "export_csv"),
        "csv_export_filename": config.get("Export", "csv_export_filename"),
        "export_excel": config.getboolean("Export", "export_excel"),
        "excel_export_filename": config.get("Export", "excel_export_filename")
    }

    return config_values
