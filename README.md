# invoice-automation

## Overview

 This is a python script designed to automate the creation of invoices. It uses a Tkinter graphical user interface (GUI) where users can input invoice details and generate a formatted invoice in PDF format from a doc template.

## Requirements

- Python 3.x
- Tkinter (usually included with Python installations)
- `python-docx` library for handling word documents
- LibreOffice for converting Word documents into PDF

## Installation

1. **Install Python 3.x** if it's not already installed. You can download it from [python.org](https://www.python.org/downloads/).

2. **Install necessary Python packages**:
   ```bash
   pip install python-docx

3. **Install LibreOffice** It is required to convert the Word document into PDF. You can download it from libreoffice.org

## How It Works

1. **Template**
- The script generates invoices by using the template document `invoicetemplate.docx`.
- There are placeholders within the template that are formatted with square brackets (e.g `[Client]`, `[Invoice Number]`).

2. **User Interface**
- The script presents an interface where the user can enter the details of the invoice, including client information, service description, as well as the hourly rate and hours worked.

3. **Replacing Placeholders**
- When the "Create Invoice" button is clicked, the application reads the `invoicetemplate.docx` file.
- It then replaces the placeholders in the template with the given values from the user from the interface.
- The replacement is performed for placeholders in paragraphs and in tables.

4. **Saving and Converting**
- The modified template document which is now filled with user-provided details is saved as `filled.docx`.
- The `filled.docx` file is then converted into a PDF using LibreOffice through the command line, where LibreOffice processes the `.docx` file and generates a `.pdf` file.
- A file dialog prompts the user to select the location and name of the newly generated PDF file.

5. **Completion**
- On completion, a success message is displayed to inform the user that the invoice has been created and saved successfully.
