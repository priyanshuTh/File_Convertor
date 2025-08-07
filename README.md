# FileConverter

A Python command-line utility to convert between various file formats including CSV, Excel, text, JSON, HTML, PDF, PPTX, JPG, PNG, and GIF.

## Features

- **CSV <> Excel**: Convert CSV files to Excel (`.xls`/`.xlsx`) and vice versa.
- **Text <> JSON**: Serialize plain text (line‑by‑line) to JSON and back.
- **Text <> CSV**: Store each line of a text file as a row in CSV.
- **HTML → PDF**: Render simple HTML documents as PDF.
- **PDF → HTML**: Extract text from PDF and wrap in HTML tags.
- **PDF <> PPTX**: Embed PDF pages as images in PowerPoint slides, and export PPTX decks to PDF via LibreOffice.
- **PDF → JPG/PNG/GIF**: Rasterize PDF pages to images.
- **JPG/PNG/GIF → PDF**: Convert image formats into PDF (single- or multi-page).
- **GIF → PDF** and **PDF → GIF**: Convert animated GIF frames to PDF and vice versa.

## Requirements

- Python 3.7+
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [fpdf](https://pyfpdf.github.io/)
- [pdf2image](https://github.com/Belval/pdf2image)
- [Pillow](https://python-pillow.org/)
- [python-pptx](https://python-pptx.readthedocs.io/)
- [PyPDF2](https://pypi.org/project/PyPDF2/)
- **LibreOffice** installed and in your PATH for PPTX-to-PDF conversion

Install dependencies via pip:

```bash
pip install pandas openpyxl fpdf2 pdf2image pillow python-pptx PyPDF2
```

On Debian/Ubuntu, install LibreOffice:

```bash
sudo apt-get update && sudo apt-get install libreoffice
```

## Usage

1. **Clone or copy** the `FileConverter` script:
   ```bash
   git clone https://github.com/priyanshuTh/File_Convertor.git
   cd File_Convertor
   ```
2. **Run** the script:
   ```bash
   python FileConverter.py
   ```
3. **Select** the desired conversion from the interactive menu by entering its number.
4. **Provide** source and destination paths when prompted.

### Examples

- Convert CSV to Excel:

  ```bash
  Choice: 1
  Source path:  C:\Users\Downloads\report.csv
  Destination path: C:\Users\Downloads\report.xlsx
  ```

- Convert PPTX to PDF:

  ```bash
  Choice: 9
  Source path: C:\Users\Downloads\presentation.pptx
  Destination path: C:\Users\Downloads\presentation.pdf
  ```

- Batch-export PDF pages to PNG images:

  ```bash
  Choice: 13
  Source path: C:\Users\Downloads\document.pdf
  Output folder: ./images
  ```

## Error Handling

- The utility validates file extensions before conversion and reports descriptive errors if mismatched.
- Ensure correct paths and install required dependencies.

## License

MIT License. See [LICENSE](LICENSE) for details.

## Contributing

Feel free to open issues or submit pull requests for additional formats or features.

