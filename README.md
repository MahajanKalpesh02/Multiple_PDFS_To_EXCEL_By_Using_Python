# PDF_To_EXCEL_INVOICES
In this repository there is code for extract the data from pdf to excel for Invoices. You can follow this Repository as per you need.
# PDF to Excel Data Extractor

This project provides a graphical user interface (GUI) for extracting data from PDF files and saving it to an Excel file. The data is extracted based on predefined patterns and is stored in both an Excel sheet and a text file.

## Features
- Select multiple PDF files to process.
- Extract predefined data fields from the PDFs.
- Save extracted data in Excel and text files.
- Track the processing progress via a progress bar.

## Requirements

To run this program, you need to have Python installed on your system, along with the following libraries:

### Required Libraries
1. `pdfplumber` - To extract text from PDF files.
2. `openpyxl` - For reading and writing Excel files.
3. `re` - Regular expressions for pattern matching.
4. `os` - To handle file and directory operations.
5. `customtkinter` - Custom themed Tkinter for building the GUI.
6. `tkinter` - Standard Python library for creating the GUI.
7. `threading` - To handle background processing of PDF files.

### Install the Required Libraries

You can install all the required libraries using `pip` by running the following command:

```bash
pip install pdfplumber openpyxl customtkinter
