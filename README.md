# Invoice Generator Tool

## Overview
The Invoice Generator is a Python application that automates invoice creation and organization. It uses a GUI to manage customer data and generates Word invoices with a structured directory system.

## Features
- Generate Word invoices with formatted placeholders.
- SQLite database for customer management.
- Organizes invoices by year, month, and week.
- Automatic invoice numbering.

## Installation
1. Download the executable:
   - Locate the `dist` folder and find the file `gui.exe` (created using PyInstaller).

2. Place the executable in your desired directory.

## Usage
1. Double-click the `gui.exe` file to launch the application.
2. Add or select a customer.
3. Input quantities and generate the invoice.
4. Find invoices in the `output/year-month-week` folder.

## Customization
The Python scripts are modular and can be customized to fit specific requirements:
- **Invoice Layout**: Modify the `invoice_template.docx` file in the `templates` directory to adjust the format of the generated invoices.
- **Database Fields**: Update the `InvoiceGenerator` class in `invoice_generator.py` to include additional customer fields or logic.
- **GUI Adjustments**: Enhance the user interface by editing `gui.py`, which uses `Tkinter` for the application layout.
- **Feature Expansion**: Add functionality like PDF exports, email integration, or multi-language support by extending the existing scripts.

## Author
Developed by WhiteTrafficLight. For inquiries, contact via GitHub.



