# non wifi python script
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

# Define input and output folder paths
input_folder = 'data/non-wifi'
output_folder = 'data/non-wifi-output'

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Loop through each Excel file in the input folder
for filename in os.listdir(input_folder):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(input_folder, filename)
        output_file_path = os.path.join(output_folder, f"Updated_{filename}")

        # Load the workbook and read the data
        book = load_workbook(file_path)
        data = pd.read_excel(file_path, skiprows=2)

        # Remove rows without dates and filter relevant data
        data.columns = ['Date', 'Time', 'SupplyTemp', 'ReturnTemp', 'Mode', 'Request', 'State', 'Status']
        data = data[data['Date'].str.match(r'^\d{2}/\d{2}/\d{4}$', na=False)].copy()
        data.reset_index(drop=True, inplace=True)

        # Filter rows for "testing" and switching states
        testing_data = data[data['Status'].str.lower() == 'testing'].copy()
        testing_data['Switch'] = testing_data['State'].ne(testing_data['State'].shift()).astype(int)
        switch_data = testing_data[testing_data['Switch'] == 1].copy()

        # Calculate the time deviation in both decimal hours and mm:ss format
        switch_data['Datetime'] = pd.to_datetime(switch_data['Date'] + ' ' + switch_data['Time'], dayfirst=True, errors='coerce')
        switch_data.dropna(subset=['Datetime'], inplace=True)
        switch_data['Time_Diff'] = switch_data['Datetime'].diff().dt.total_seconds() / 3600  # in hours
        switch_data['Decimal_Deviation'] = switch_data['Time_Diff'] - 2  # 2 hours as baseline

        # Convert deviation in hours to mm:ss format for clarity
        switch_data['Deviation_Min_Sec'] = (switch_data['Decimal_Deviation'] * 3600).apply(
            lambda x: f"{int(abs(x) // 60)}:{int(abs(x) % 60):02}" if pd.notnull(x) else None
        )

        # Add deviation analysis to a new sheet
        with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a') as writer:
            writer.book = book
            writer.sheets = {ws.title: ws for ws in book.worksheets}
            switch_data[['Datetime', 'State', 'Decimal_Deviation', 'Deviation_Min_Sec']].to_excel(
                writer, sheet_name='Deviation Analysis', index=False
            )

        # Apply conditional formatting in the new sheet
        deviation_analysis_sheet = book['Deviation Analysis']
        
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # Yellow: for deviations over 2 minutes (0.0333 hours) but less than 5 minutes (0.0833 hours)
        deviation_analysis_sheet.conditional_formatting.add(
            'C2:C1000',
            CellIsRule(operator='between', formula=['0.0333', '0.0833'], fill=yellow_fill)
        )

        # Red: for deviations over 5 minutes (0.0833 hours)
        deviation_analysis_sheet.conditional_formatting.add(
            'C2:C1000',
            CellIsRule(operator='greaterThan', formula=['0.0833'], fill=red_fill)
        )

        # Save the workbook with updated sheet and formatting
        book.save(output_file_path)
        print(f"Processed and saved {filename} with deviation analysis and formatting.")

