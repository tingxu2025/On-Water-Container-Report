import pandas as pd
import xlsxwriter

# Load the Excel file
file_path = r"C:\Users\tin.xu\Desktop\Report\depart_report.xlsx"  # Update path if needed
df_depart_report = pd.read_excel(file_path)

# Ensure "Container#" column exists before counting unique values
if "Container#" in df_depart_report.columns and "Destination" in df_depart_report.columns:
    # Count unique containers per destination
    df_destination_count = df_depart_report.groupby("Destination")["Container#"].nunique().reset_index()
    df_destination_count.rename(columns={"Container#": "Qty of Ctns", "Destination": "WM"}, inplace=True)

    # Calculate total unique containers across all destinations
    total_unique_ctns = df_depart_report["Container#"].dropna().nunique()
else:
    # Handle missing columns
    df_destination_count = pd.DataFrame(columns=["WM", "Qty of Ctns"])
    total_unique_ctns = 0

# Ensure all WM categories exist, even if missing from data
wm_list = ["ELF", "ELO", "ELG", "DAL", "HOU", "SAC", "NY"]
df_existing_wm = pd.DataFrame({"WM": wm_list})

# Merge to ensure missing WM locations are shown as 0
df_summary = df_existing_wm.merge(df_destination_count, on="WM", how="left").fillna(0)
df_summary["Qty of Ctns"] = df_summary["Qty of Ctns"].astype(int)

# Add numbering column
df_summary.insert(0, "No.", range(1, len(df_summary) + 1))

# Remove any pre-existing total row
df_summary = df_summary[df_summary["WM"] != "Total Ctns"]

# Add total row with the sum of unique containers
df_total = pd.DataFrame([["", "Total Ctns", total_unique_ctns]], columns=df_summary.columns)
df_summary = pd.concat([df_summary, df_total], ignore_index=True)

# Delete row 14 if it exists
df_summary = df_summary.iloc[:13]  # Keeps only first 13 rows (removes row 14)

# Define output path
summary_output_path = r"C:\Users\tin.xu\Desktop\Report\depart_summary.xlsx"

# Save to Excel with formatting
with pd.ExcelWriter(summary_output_path, engine="xlsxwriter") as writer:
    df_summary.to_excel(writer, index=False, sheet_name="Depart Summary", startrow=2)  # Start from A3

    # Get workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets["Depart Summary"]

    # Define formats
    title_format = workbook.add_format(
        {"bold": True, "font_size": 14, "align": "center", "bg_color": "#B0C4DE", "border": 1})
    header_format = workbook.add_format(
        {"bold": True, "align": "center", "bg_color": "#FFA500", "border": 1})  # Orange header
    data_format = workbook.add_format({"align": "center", "border": 1})  # Border applied only to columns A to C
    total_format = workbook.add_format({"bold": True, "align": "center", "border": 1})  # Bold for total row

    # Merge title row
    worksheet.merge_range("A1:C1", "Depart Summary --- WHG", title_format)

    # Write header row explicitly in row 2
    worksheet.write(1, 0, "No.", header_format)  # Column A
    worksheet.write(1, 1, "WM", header_format)  # Column B
    worksheet.write(1, 2, "Qty of Ctns", header_format)  # Column C

    # Set column widths
    worksheet.set_column("A:A", 13)  # First column (No.)
    worksheet.set_column("B:B", 20)  # Second column (WM)
    worksheet.set_column("C:C", 13)  # Third column (Qty of Ctns)

    # Apply formatting only to columns A to C
    for row in range(2, len(df_summary) + 2):  # Data rows
        for col in range(3):  # Only columns A to C
            value = df_summary.iloc[row - 2, col]
            worksheet.write(row, col, value, data_format)

    # Format total row (only A to C)
    worksheet.set_row(len(df_summary) + 1, None, total_format)
    for col in range(3):
        worksheet.write(len(df_summary) + 1, col, df_summary.iloc[-1, col], total_format)

    # EXPLICITLY REMOVE ALL FORMATTING PAST COLUMN C
    for col in range(3, 50):  # Remove formatting from column D onward
        worksheet.set_column(col, col, None, None)