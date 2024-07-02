import pandas as pd
from datetime import datetime, timedelta
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import re

def process_file(file_path):
    # Read the text file
    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Extract the relevant data
    data = []
    skip_lines = 0
    for line in lines:
        if skip_lines > 0:
            skip_lines -= 1
            continue
        if "Page" in line:
            skip_lines = 4
            continue
        if line.strip() and re.match(r'^\d', line):
            # Use regex to split the line into columns
            match = re.match(r'(\S+)\s+(\S+)\s+(.+?)\s+(\d{6})\s+(\d{4})\s+(\d+)\s+(\d+)\s+(\S+)\s+(.+)', line)
            if match:
                coil = match.group(1)
                type_ = match.group(2)
                part = match.group(3)
                date_str = match.group(4)
                try:
                    weight = int(match.group(7))
                    pieces = int(match.group(6))
                except ValueError:
                    continue  # Skip lines with invalid numeric values
                whse = match.group(8)
                status = match.group(9)
                data.append([coil, type_, part, date_str, pieces, weight, whse, status])

    # Convert to DataFrame
    df = pd.DataFrame(data, columns=['Coil', 'Type', 'Part', 'Date', 'Pieces', 'Weight', 'Whse', 'Status'])

    # Convert Date column to datetime
    df['Date'] = pd.to_datetime(df['Date'], format='%m%d%y', errors='coerce').dt.date

    # Filter out rows with invalid dates
    df = df.dropna(subset=['Date'])

    # Filter for inventory older than 6 months
    six_months_ago = datetime.now().date() - timedelta(days=6*30)
    filtered_df = df[df['Date'] < six_months_ago]

    # Calculate totals
    total_coils = filtered_df['Coil'].nunique()
    total_pieces = filtered_df['Pieces'].sum()
    total_weight = filtered_df['Weight'].sum()
    total_lines = len(filtered_df)

    # Save to Excel
    output_file = 'Filtered_Inventory.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Filtered Inventory')
        summary_df = pd.DataFrame({
            'Total Coils': [total_coils],
            'Total Pieces': [total_pieces],
            'Total Weight': [total_weight],
            'Total Lines': [total_lines]
        })
        summary_df.to_excel(writer, index=False, sheet_name='Summary')

    # Load the workbook to apply formatting
    wb = writer.book
    ws_filtered = wb['Filtered Inventory']
    ws_summary = wb['Summary']

    # Apply formatting to the 'Filtered Inventory' sheet
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    for cell in ws_filtered[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws_filtered.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        adjusted_width = (max_length + 2)
        ws_filtered.column_dimensions[col[0].column_letter].width = adjusted_width

    # Apply formatting to the 'Summary' sheet
    for cell in ws_summary[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws_summary.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        adjusted_width = (max_length + 2)
        ws_summary.column_dimensions[col[0].column_letter].width = adjusted_width

    wb.save(output_file)

    print(f"Filtered inventory saved to {output_file}")
    print(f"Total Coils: {total_coils}")
    print(f"Total Pieces: {total_pieces}")
    print(f"Total Weight: {total_weight}")
    print(f"Total Lines: {total_lines}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python InvGreaterThan6Months.py <path_to_text_file>")
    else:
        process_file(sys.argv[1])