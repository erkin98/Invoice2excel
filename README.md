# Invoice2excel

This is a professional tool to convert PDF invoice files to Excel.

PDF must be in this [format](https://www.oaib.org.tr/files/downloads/Bilgi-merkezi/ihracat-merkezi/dis-ticarette-kullnlan-fatu/gum-gen-soz-1.jpg).

Currently, the script supports only searchable PDFs.

### Installation

1.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

### How to use

Run the application:

```bash
python pdf_last.py
```
or
```bash
python main.py
```

1.  From the opened dialog box, select the PDF folder.
2.  Select the PDF file from the list.
3.  Click "Submit".
4.  Wait for processing.
5.  A save dialog box will appear. Select the destination to save the Excel file.

### Structure

The project has been refactored into a modular structure:
- `src/`: Contains the source code.
    - `extractor.py`: PDF processing logic.
    - `ui.py`: User interface (PySimpleGUI).
    - `exporter.py`: Excel/CSV export logic.
- `main.py`: Main entry point.
- `pdf_last.py`: Backward compatibility entry point.

### Especially thanks for [pdfplumber team](https://github.com/jsvine/pdfplumber)
