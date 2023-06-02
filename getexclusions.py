import csv
from tkinter import messagebox

from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import glob
import datetime
import random
import time
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class ExclusionsProcessor:
    def __init__(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "MAIN"

    def combine_functions(self):
        self.process_csv_files()
        self.call_pivot_by_exclusion()
        self.create_pivot_by_exclusion_mode()
        self.search_keywords_and_create_pivot()
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

    def search_keywords_and_create_pivot(self):
        # Create a new sheet for "Not Recommended"
        not_recommended_sheet = self.workbook.create_sheet(title="Not Recommended")

        # Load the "not.xlsx" file from the directory
        not_file_path = os.path.expanduser("~/S1GOB/not.xlsx")
        if not os.path.exists(not_file_path):
            print("Error: File 'not.xlsx' not found.")
            return

        # Load the "not.xlsx" workbook
        not_workbook = load_workbook(not_file_path)

        # Get the active sheet of the "not.xlsx" workbook
        not_worksheet = not_workbook.active

        # Define the folder path containing the CSV files
        raw_data_folder = os.path.expanduser("~/S1GOB/RawData")

        # Get the CSV file paths
        csv_files = glob.glob(os.path.join(raw_data_folder, "exclusions-*.csv"))

        print("CSV Files:")
        for csv_file in csv_files:
            print(csv_file)

        # Get the column index of the "description" column in CSV files
        description_column_index = None

        # Iterate over each CSV file to find the "description" column index
        for csv_file in csv_files:
            with open(csv_file, 'r') as file:
                reader = csv.reader(file)
                header_row = next(reader, None)  # Read the header row
                if header_row and "description" in header_row:
                    description_column_index = header_row.index("description")
                    break

        # If "description" column index is found, proceed with comparison
        if description_column_index is not None:
            # Create a dictionary to store the counts
            count_dict = {}

            # Iterate over each cell in the "not.xlsx" worksheet
            for row in not_worksheet.iter_rows():
                for cell in row:
                    # Get the cell value
                    cell_value = cell.value

                    # Count the occurrences of cell value in the CSV files
                    count = 0
                    for csv_file in csv_files:
                        with open(csv_file, 'r') as file:
                            reader = csv.reader(file)
                            for csv_row in reader:
                                if cell_value.lower() in csv_row[description_column_index].lower():
                                    count += 1

                    # Store the count in the dictionary if it's greater than zero
                    if count > 0:
                        count_dict[cell_value] = count

            # Sort the count dictionary in descending order by values
            sorted_counts = sorted(count_dict.items(), key=lambda x: x[1], reverse=True)

            # Write the sorted counts to the "Not Recommended" sheet
            for cell_value, count in sorted_counts:
                not_recommended_sheet.append([cell_value, count])

        # Print the contents of the "Not Recommended" sheet for testing
        print("Not Recommended Sheet:")
        for row in not_recommended_sheet.iter_rows(values_only=True):
            print(row)

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
