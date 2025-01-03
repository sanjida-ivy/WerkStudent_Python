import os
import fitz  # PyMuPDF for PDF parsing
import pandas as pd
import re
import openpyxl
from tkinter import Tk, Label, Button, filedialog, StringVar
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button as TTKButton
from ttkbootstrap.widgets import Label as TTKLabel
from ttkbootstrap.widgets import Progressbar
from ttkbootstrap.constants import PRIMARY, SUCCESS
import subprocess
import win32com.client as win32  # For creating actual pivot tables
from datetime import datetime
import locale
import csv
import requests

# Add your API key for currency conversion
API_KEY = "355440549f101a8cbd2be119bf857cde"

def convert_currency(amount, from_currency, to_currency="USD"):
    """
    Convert currency using a real-time API.

    :param amount: Amount to convert.
    :param from_currency: Source currency code (e.g., "EUR").
    :param to_currency: Target currency code (e.g., "USD").
    :return: Converted amount.
    """
    try:
        # API endpoint for currency conversion
        api_url = f"https://api.exchangerate-api.com/v4/latest/{from_currency}"
        response = requests.get(api_url, params={"apikey": API_KEY})
        response.raise_for_status()
        data = response.json()

        # Get conversion rate
        rate = data["rates"].get(to_currency)
        if rate is None:
            raise ValueError(f"Conversion rate for {to_currency} not found.")

        # Convert the amount
        converted_amount = float(amount) * rate
        return round(converted_amount, 2)
    except Exception as e:
        print(f"Error during currency conversion: {e}")
        return None

def extract_data_from_pdf(file_path):
    """
    Extract specific values dynamically from a PDF file.

    :param file_path: Path to the PDF file.
    :return: Dictionary containing file name, date, and converted values.
    """
    try:
        # German to English month mapping
        german_to_english_months = {
            "Januar": "January", "Februar": "February", "März": "March",
            "April": "April", "Mai": "May", "Juni": "June",
            "Juli": "July", "August": "August", "September": "September",
            "Oktober": "October", "November": "November", "Dezember": "December"
        }

        pdf_document = fitz.open(file_path)
        for page in pdf_document:
            text = page.get_text()

            # Extract date dynamically
            date_match = re.search(
                r"(?:Date|Invoice date|Invoice Date|Datum|Dated):?\s*(\d{1,2}\.\s*\w+\s*\d{4}|\w{3,9}\s+\d{1,2},\s+\d{4})",
                text, re.IGNORECASE)
            raw_date = date_match.group(1).strip() if date_match else "Unknown"

            # Initialize date to "Unknown"
            date = "Unknown"

            if raw_date != "Unknown":
                try:
                    # Check if the date is in German format
                    if re.search(r"\d{1,2}\.\s*\w+\s*\d{4}", raw_date):
                        for german_month, english_month in german_to_english_months.items():
                            if german_month in raw_date:
                                raw_date = raw_date.replace(german_month, english_month)
                                break
                        date = datetime.strptime(raw_date, "%d. %B %Y").strftime("%Y-%m-%d")
                    elif re.search(r"\w{3,9}\s+\d{1,2},\s+\d{4}", raw_date):
                        # Manual parsing for English dates like "Nov 26, 2016"
                        match = re.match(r"(\w+)\s+(\d{1,2}),\s+(\d{4})", raw_date)
                        if match:
                            month_name, day, year = match.groups()
                            month_number = datetime.strptime(month_name, "%b").month
                            date = f"{year}-{month_number:02d}-{int(day):02d}"
                        else:
                            print(f"Failed to extract components from: {raw_date}")
                except Exception as e:
                    print(f"Error parsing date '{raw_date}': {e}")

            # Extract and convert values
            gross_amount_match = re.search(r"Gross Amount incl\. VAT\s*([\d,]+\s*[,]\d{2})\s*\u20ac", text, re.IGNORECASE)
            total_amount_match = re.search(r"Total\s*USD\s*\$\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

            converted_value = "Unknown"
            if gross_amount_match:
                raw_value = gross_amount_match.group(1).replace(".", "").replace(",", ".")
                euro_value = float(raw_value)
                converted_value = convert_currency(euro_value, "EUR", "USD")
            elif total_amount_match:
                raw_value = total_amount_match.group(1).replace(",", "")
                converted_value = float(raw_value)

            # Return extracted values with converted value only
            return {
                "file_name": os.path.basename(file_path),
                "date": date,
                "value": converted_value
            }
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None
    

    """
    Extract specific values (Gross Amount incl. VAT and Total) dynamically from a PDF file.

    :param file_path: Path to the PDF file.
    :return: Dictionary containing file name, date, and extracted values.
    """
    try:
        # German to English month mapping
        german_to_english_months = {
            "Januar": "January", "Februar": "February", "März": "March",
            "April": "April", "Mai": "May", "Juni": "June",
            "Juli": "July", "August": "August", "September": "September",
            "Oktober": "October", "November": "November", "Dezember": "December"
        }

        pdf_document = fitz.open(file_path)
        for page in pdf_document:
            text = page.get_text()

            # Extract date dynamically
            date_match = re.search(
                r"(?:Date|Invoice date|Invoice Date|Datum|Dated):?\s*(\d{1,2}\.\s*\w+\s*\d{4}|\w{3,9}\s+\d{1,2},\s+\d{4})",
                text, re.IGNORECASE)
            raw_date = date_match.group(1).strip() if date_match else "Unknown"

            # Debugging: Log the extracted raw date
            print(f"Raw date extracted: {raw_date}")

            # Initialize date to "Unknown"
            date = "Unknown"

            if raw_date != "Unknown":
                try:
                    # Check if the date is in German format
                    if re.search(r"\d{1,2}\.\s*\w+\s*\d{4}", raw_date):
                        for german_month, english_month in german_to_english_months.items():
                            if german_month in raw_date:
                                raw_date = raw_date.replace(german_month, english_month)
                                break
                        date = datetime.strptime(raw_date, "%d. %B %Y").strftime("%Y-%m-%d")
                    elif re.search(r"\w{3,9}\s+\d{1,2},\s+\d{4}", raw_date):
                        # Manual parsing for English dates like "Nov 26, 2016"
                        match = re.match(r"(\w+)\s+(\d{1,2}),\s+(\d{4})", raw_date)
                        if match:
                            month_name, day, year = match.groups()
                            month_number = datetime.strptime(month_name, "%b").month
                            date = f"{year}-{month_number:02d}-{int(day):02d}"
                        else:
                            print(f"Failed to extract components from: {raw_date}")
                except Exception as e:
                    print(f"Error parsing date '{raw_date}': {e}")

            # Extract "Gross Amount incl. VAT" and "Total" values
            gross_amount_match = re.search(r"Gross Amount incl\. VAT\s*([\d,]+\s*[,]\d{2})\s*\u20ac", text, re.IGNORECASE)
            total_amount_match = re.search(r"Total\s*USD\s*\$\s*([\d,]+\.\d{2})", text, re.IGNORECASE)

            value = "Unknown"
            if gross_amount_match:
                value = f"{gross_amount_match.group(1)} \u20ac"
            elif total_amount_match:
                value = f"USD $ {total_amount_match.group(1)}"

            # Return extracted values
            return {
                "file_name": os.path.basename(file_path),
                "date": date,
                "value": value,
            }
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None

def export_to_excel_with_pivot(data):
    """
    Export extracted data to an Excel file with a customized pivot table and a chart.

    :param data: List of dictionaries containing file name, date, and values.
    """
    output_file = os.path.abspath("output_xlsx.xlsx")
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False, sheet_name="Sheet1")

    try:
        # Open Excel for processing
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False  # Set to True for debugging or visual inspection

        wb = excel.Workbooks.Open(output_file)
        ws_raw = wb.Worksheets("Sheet1")

        # Format the "value" column as Currency (USD)
        last_row = ws_raw.Cells(ws_raw.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
        value_column = ws_raw.Columns("C:C")
        value_column.NumberFormat = "$#,##0.00"

        # Add a new worksheet for the pivot table
        ws_pivot = wb.Sheets.Add()
        ws_pivot.Name = "Sheet2"

        # Define the pivot table range
        pivot_source_range = ws_raw.Range(f"A1:C{last_row}")
        pivot_table_destination = ws_pivot.Range("A3")

        # Create pivot cache and table
        pivot_cache = wb.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=pivot_source_range,
            Version=6  # xlPivotTableVersion15
        )

        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=pivot_table_destination, TableName="ModifiedPivotTable"
        )

        # Configure pivot table fields
        # Add "date" as a Row field
        date_field = pivot_table.PivotFields("date")
        date_field.Orientation = 1  # xlRowField
        date_field.Position = 1  # Set as the first row field

        # Add "file_name" as a Column field
        file_name_field = pivot_table.PivotFields("file_name")
        file_name_field.Orientation = 2  # xlColumnField
        file_name_field.Position = 1  # Set as the first column field

        # Add "value" as a Data field
        value_field = pivot_table.PivotFields("value")
        value_field.Orientation = 4  # xlDataField
        value_field.Function = -4157  # xlSum
        value_field.NumberFormat = "$#,##0.00"

        # Add a Report Filter for "file_name"
        report_filter = pivot_table.PivotFields("file_name")
        report_filter.Orientation = 3  # xlPageField
        report_filter.Position = 1

        # Apply custom styling and auto-fit columns
        ws_pivot.Columns.AutoFit()

        # Add a chart based on the pivot table
        chart = ws_pivot.Shapes.AddChart2(
            201,  # xlColumnClustered
            ws_pivot.Cells(1, 6).Top,  # Chart top position
            ws_pivot.Cells(1, 6).Left,  # Chart left position
            500, 300  # Chart width and height
        )
        chart.Chart.SetSourceData(ws_pivot.Range(pivot_table_destination.Address))
        chart.Chart.HasTitle = True
        chart.Chart.ChartTitle.Text = "Pivot Table Summary"

        # Save and close the workbook
        wb.Save()
        print("Modified pivot table and chart created successfully.")
    except Exception as e:
        print(f"Error occurred while modifying the pivot table: {e}")
    finally:
        if 'wb' in locals():
            wb.Close(SaveChanges=True)
        if 'excel' in locals():
            excel.Quit()

    print(f"Data exported to {output_file}.")
    return output_file

def export_to_csv(data, output_file='output.csv'):
    """
    Export extracted data to a CSV file with a semicolon delimiter.

    :param data: List of dictionaries containing extracted data.
    :param output_file: Name of the output CSV file.
    """
    try:
        # Define the header based on the keys of the first dictionary
        headers = data[0].keys()

        # Open the CSV file in write mode
        with open(output_file, 'w', newline='', encoding='utf-8') as file:
            # Create a CSV writer object with a semicolon delimiter
            writer = csv.DictWriter(file, fieldnames=headers, delimiter=';')

            # Write the header row
            writer.writeheader()

            # Write the data rows
            writer.writerows(data)

        print(f"Data successfully exported to {output_file}.")

    except Exception as e:
        print(f"An error occurred while exporting to CSV: {e}")

def process_files(selected_files, status_var, progress_bar):
    """
    Process selected PDF files, extract data, and generate the Excel file.

    :param selected_files: List of selected PDF file paths.
    :param status_var: Tkinter StringVar to update the status label.
    :param progress_bar: ttk Progressbar to show progress.
    """
    extracted_data = []
    total_files = len(selected_files)
    progress_step = 100 / total_files if total_files > 0 else 0

    progress_bar['value'] = 0
    for i, file_path in enumerate(selected_files, start=1):
        data = extract_data_from_pdf(file_path)
        if data:
            extracted_data.append(data)
        progress_bar['value'] += progress_step
        progress_bar.update_idletasks()  # Force the UI to update

    if extracted_data:
        # Export to Excel with pivot table
        excel_path = export_to_excel_with_pivot(extracted_data)
        # Export to CSV
        export_to_csv(extracted_data, output_file='output_csv.csv')

        if excel_path and os.name == "nt":  # Only attempt to open if on Windows
            try:
                subprocess.run(["start", excel_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                status_var.set("Processing completed! Opening Excel file...")
            except Exception as e:
                status_var.set(f"File saved, but could not open Excel: {e}")
        else:
            status_var.set("Processing completed! File saved.")
    else:
        status_var.set("No valid data extracted from selected files.")

    progress_bar['value'] = 100

def select_files(status_var, progress_bar):
    """
    Open a file dialog to select PDF files.

    :param status_var: Tkinter StringVar to update the status label.
    :param progress_bar: ttk Progressbar to show progress.
    """
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if files:
        status_var.set(f"Selected {len(files)} file(s). Processing...")
        process_files(files, status_var, progress_bar)
    else:
        status_var.set("No files selected.")




def main():
    """
    Create a modern Tkinter GUI for selecting and processing PDF files.
    """
    style = Style(theme="flatly")  # Bootstrap-inspired theme
    root = style.master

    # Remove the default title bar
    root.overrideredirect(True)

    # Set dimensions and center the window
    window_width = 600
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = (screen_width // 2) - (window_width // 2)
    y_coordinate = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    # Add a close button manually
    close_button = TTKButton(root, text="X", bootstyle="danger", command=root.destroy)
    close_button.place(x=window_width - 40, y=10, width=30, height=30)

    status_var = StringVar()
    status_var.set("Select PDF files to process.")

    header_label = TTKLabel(root, text="PDF Data Extraction Tool", font=("Helvetica", 20, "bold"), bootstyle=PRIMARY)
    header_label.pack(pady=20)

    select_button = TTKButton(root, text="Select PDF Files", bootstyle=PRIMARY, command=lambda: select_files(status_var, progress_bar))
    select_button.pack(pady=10)

    

    progress_bar = Progressbar(root, orient="horizontal", mode="indeterminate", length=400, bootstyle=PRIMARY)
    progress_bar.pack(pady=10)

    status_label = TTKLabel(root, textvariable=status_var, wraplength=500, bootstyle=SUCCESS)
    status_label.pack(pady=20)

    root.mainloop()
if __name__ == "__main__":
    main()
