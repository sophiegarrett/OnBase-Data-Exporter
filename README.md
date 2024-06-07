# OnBase Data Exporter

[Docs](https://github.com/sophiegarrett/OnBase-Data-Exporter/wiki) | [Quick Start Guide](https://github.com/sophiegarrett/OnBase-Data-Exporter/wiki/Quick-Start-Guide) | [Configuration](https://github.com/sophiegarrett/OnBase-Data-Exporter/wiki/Configuration) | [Troubleshooting](https://github.com/sophiegarrett/OnBase-Data-Exporter/wiki/Troubleshooting)

### A program that translates OnBase data dumps into a usable format.

### Features:

- Translate OnBase data dump text files into readable, searchable CSV and Excel files
- Automatically combine documents containing multiple files into single PDF files

### Supported File Types:

- Data export will work for documents of any file type.
- Only jpeg, png, and tiff files can be combined. If requested, more file types could be supported in the future.

### Dependencies:

- [pillow](https://pypi.org/project/pillow/)
- [XlsxWriter](https://pypi.org/project/XlsxWriter/)
