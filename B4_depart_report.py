import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Define file paths
input_path = "Sterling.xlsx"  # Input file (modify path as needed)
output_path = "depart_report.xlsx"  # Output file

# Load the 2024 sheet from the Excel file
try:
    df_sterling = pd.read_excel(input_path, sheet_name='2024', skiprows=1)

    # Define correct column names
    correct_column_names = [
        "Week", "Shipper", "Ref#", "Container#", "Qty", "ETD", "ETA", "AP", "AN", "DO",
        "Origin", "Ocean Freight", "Delivery Date", "Destination", "Carrier", "Note", "B/L #",
        "Invoice Needed", "Invoice Sent"
    ]
    df_sterling.columns = correct_column_names

    # Convert date columns to datetime format
    date_columns = ["Delivery Date", "ETD", "ETA", "AN", "DO"]
    for col in date_columns:
        df_sterling[col] = pd.to_datetime(df_sterling[col], format="%Y-%m-%d", errors='coerce')

    # Remove invalid dates (NaT values) from 'Delivery Date'
    df_sterling = df_sterling.dropna(subset=['Delivery Date'])

    # Define the date range for filtering
    start_date = pd.Timestamp("2025-04-21")
    end_date = pd.Timestamp("2025-04-25")

    # Filter the DataFrame based on the date range
    df_filtered = df_sterling[
        (df_sterling['Delivery Date'] >= start_date) & (df_sterling['Delivery Date'] <= end_date)
        ].copy()

    # Convert date columns back to MM/DD/YYYY format for Excel output
    for col in date_columns:
        df_filtered[col] = df_filtered[col].dt.strftime('%m/%d/%Y')

    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Save the filtered results into a new Excel file
    df_filtered.to_excel(output_path, index=False)

    # Apply formatting using openpyxl
    wb = load_workbook(output_path)
    ws = wb.active

    # Center align all cells & set column width
    for col in ws.columns:
        max_length = 15  # Default length
        col_letter = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2  # Add padding

    # Save the formatted file
    wb.save(output_path)