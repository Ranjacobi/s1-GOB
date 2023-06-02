from tkinter import messagebox
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import glob
import os
import datetime
import random
import time

class ExclusionsProcessor:
    def __init__(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "MAIN"
    def combine_functions(self):
        self.process_csv_files()
        self.call_pivot_by_exclusion()
        self.create_pivot_by_exclusion_mode()
        self.scan_excel_file()
        self.create_exclusions_data()

    def process_csv_files(self):
        # Step 1: Importing all CSVs from the RawData folder
        raw_data_folder = os.path.expanduser("~/S1GOB/RawData")
        if not os.path.exists(raw_data_folder):
            try:
                os.makedirs(raw_data_folder)
                print("Success: Created 'RawData' folder.")
            except OSError:
                print("Error: Failed to create 'RawData' folder. Please check write permissions.")
                return

        if not os.access(raw_data_folder, os.W_OK):
            print("Error: Insufficient permissions to write to 'RawData' folder.")
            return

        csv_files = glob.glob(os.path.join(raw_data_folder, "exclusions-*.csv"))

        total_files = len(csv_files)  # Total number of CSV files
        processed_files = 0  # Counter for processed files

        print("Processing the 'exclusions' CSV file...")
        for csv_file in csv_files:
            # Read data from the CSV file
            with open(csv_file, 'r') as file:
                # Assuming the CSV file has comma-separated values
                lines = file.readlines()
                data = [line.strip().split(',') for line in lines]

            # Write data to the Excel worksheet
            for row in data:
                self.worksheet.append(row)

            processed_files += 1

            # Calculate progress percentage
            progress = (processed_files / total_files) * 100
            print(f"Progress: {progress:.2f}%")  # Print progress percentage

            # Check if a progress update is due
            if processed_files % 10 == 0:
                print("Still processing...")

            # Add a delay to simulate processing time
            time.sleep(0.1)

        print("Processing completed.")

    def pivot_by_exclusion(self, sheet_name, exclusion_column, another_exclusion_column):
        # Create a new sheet for the pivot table
        pivot_sheet = self.workbook.create_sheet(title=sheet_name)

        # Convert the sheet data to a DataFrame
        data = list(self.worksheet.values)
        if not data:
            print(f"No data found in sheet 'MAIN'.")
            return

        columns = data.pop(0)
        df = pd.DataFrame(data, columns=columns)

        # Reset index to handle potential duplicate entries
        df.reset_index(inplace=True)

        # Pivot the data by the exclusion column and count the occurrences of 'scopePath'
        pivot_table = df.groupby([exclusion_column, another_exclusion_column]).size().reset_index(name='count')

        # Sort the pivot table by 'count' column in descending order
        pivot_table = pivot_table.sort_values('count', ascending=False)

        # Write the pivot table to the sheet
        for row in dataframe_to_rows(pivot_table, index=False, header=True):
            pivot_sheet.append(row)

    def call_pivot_by_exclusion(self):
        exclusion_column = "description"
        another_exclusion_column = "scopePath"
        sheet_name = "Pivot by Exclusion"
        self.pivot_by_exclusion(sheet_name, exclusion_column, another_exclusion_column)

    def create_pivot_by_exclusion_mode(self):
        # Create a new sheet for the pivot table
        pivot_sheet = self.workbook.create_sheet(title="Pivot by Exclusion Mode")

        # Convert the sheet data to a DataFrame
        data = list(self.worksheet.values)
        if not data:
            print(f"No data found in sheet 'MAIN'.")
            return

        columns = data.pop(0)
        df = pd.DataFrame(data, columns=columns)

        # Reset index to handle potential duplicate entries
        df.reset_index(inplace=True)

        # Pivot the data by the 'mode' column and count the occurrences of 'scopePath'
        pivot_table = df.groupby(['mode', 'scopePath']).size().reset_index(name='count')

        # Sort the pivot table by 'count' column in descending order
        pivot_table = pivot_table.sort_values('count', ascending=False)

        # Write the pivot table to the sheet
        for row in dataframe_to_rows(pivot_table, index=False, header=True):
            pivot_sheet.append(row)

    def scan_excel_file(self):
        # Specify the file path
        excel_file = os.path.expanduser("~/GOB/not.xlsx")

        # Check if the file exists
        if not os.path.isfile(excel_file):
            print("Excel file does not exist.")
            return

        # Read the Excel file
        df = pd.read_excel(excel_file, engine='openpyxl')

        # Create a new sheet called "Not Recommended" if it doesn't already exist
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
            writer.book = Workbook()
            writer.sheets = {ws.title: ws for ws in writer.book.worksheets}  # Set the sheets dictionary

            if 'Not Recommended' not in writer.sheets:
                writer.book.create_sheet('Not Recommended')

            # Write the DataFrame to the "Not Recommended" sheet
            df.to_excel(writer, sheet_name='Not Recommended', index=False)

            # Count the keywords and create a pivot table
            total_rows = len(df)
            processed_rows = 0
            keyword = 'keyword'  # Modify this with the actual keyword you want to search

            for i, row in df.iterrows():
                count = sum(row.astype(str).str.contains(keyword, case=False))
                df.at[i, 'Count'] = count

                processed_rows += 1

                # Calculate progress percentage
                progress = (processed_rows / total_rows) * 100
                print(f"Progress: {progress:.2f}%")  # Print progress percentage
                print("Still searching...")

            pivot_table = df.pivot_table(index='Count', aggfunc='size')

            # Write the pivot table to the "Not Recommended" sheet
            pivot_table.to_excel(writer, sheet_name='Not Recommended', startrow=total_rows + 2)

        print("Scanning and counting completed.")

        # Prompt the user if they want to open the modified file
        response = input("Do you want to open the modified file? (yes/no): ")
        if response.lower() == 'yes':
            # Open the file using the default application
            os.system(f'open "{excel_file}"')

    def create_exclusions_data(self):
        # Iterate over each sheet in the workbook
        sheet_names = self.workbook.sheetnames
        total_sheets = len(sheet_names)
        processed_sheets = 0

        print("Processing sheets...")
        for sheet_name in sheet_names:
            worksheet = self.workbook[sheet_name]

            # Set color for the first row
            fill = PatternFill(start_color='4916ad', end_color='4916ad', fill_type='solid')
            for cell in worksheet[1]:
                cell.fill = fill
                cell.font = Font(color='FFFFFF')

            # Apply borders to the table
            rows = list(worksheet.rows)
            border_style = Side(border_style='thin', color='000000')
            border = Border(top=border_style, bottom=border_style, left=border_style, right=border_style)
            for row in rows:
                for cell in row:
                    cell.border = border

            # Apply autofilter
            worksheet.auto_filter.ref = worksheet.dimensions

            # Autofit columns
            for column in worksheet.columns:
                max_length = 0
                column = list(column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                column_name = column[0].column_letter
                worksheet.column_dimensions[column_name].width = adjusted_width

            processed_sheets += 1

            # Calculate progress percentage
            progress = (processed_sheets / total_sheets) * 100
            print(f"Progress: {progress:.2f}%")  # Print progress percentage
            print("Still processing...")

        # Save the modified Excel file
        filename = datetime.datetime.now().strftime("Exclusions_%d%m%Y_{}.xlsx".format(random.randint(0, 99)))
        s1gob_folder = os.path.expanduser("~/S1GOB")
        if not os.path.exists(s1gob_folder):
            os.makedirs(s1gob_folder)
        file_path = os.path.join(s1gob_folder, filename)
        self.workbook.save(file_path)

        # Wait for 2 seconds
        time.sleep(2)

        print("\n\nYour New Exclusions Excel File is ready:", file_path)

        # Prompt the user if they want to open the new file
        response = messagebox.askquestion("File Created",
                                          f"Do you want to open the new file?\n\nFile name: {file_path}")
        if response == 'yes':
            # Open the file using the default application
            os.system(f'open "{file_path}"')

        print("Success: Excel Exclusions file saved.")