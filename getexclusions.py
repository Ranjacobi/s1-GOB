def search_keywords_and_create_pivot(self):
    # Create a new sheet for "Not Recommended"
    not_recommended_sheet = self.workbook.create_sheet(title="Not Recommended")

    # Load the "not.xlsx" file from the directory
    not_file_path = os.path.expanduser("~/S1GOB/not.xlsx")
    if not os.path.exists(not_file_path):
        print("Error: File 'not.xlsx' not found.")
        return

    # Load the "not.xlsx" workbook
    not_workbook = pd.read_excel(not_file_path)

    # Get the "MAIN" sheet from the workbook
    main_sheet = self.workbook["MAIN"]

    # Get the column index of the "description" column in the "MAIN" sheet
    description_column_index = None
    header_row = main_sheet[1]  # Assuming the header row is the first row
    for index, cell in enumerate(header_row, start=1):
        if cell.value == "description":
            description_column_index = index
            break

    # If "description" column index is found, proceed with comparison
    if description_column_index is not None:
        # Create a dictionary to store the counts and corresponding cell addresses
        count_dict = {}
        cell_address_dict = {}

        # Iterate over each cell in the "not.xlsx" worksheet
        for row in not_workbook.iterrows():
            for cell_value in row[1].values:
                if pd.notnull(cell_value):
                    # Count the occurrences of cell value in the "MAIN" sheet
                    count = 0
                    cell_addresses = []
                    for row_idx, row in enumerate(main_sheet.iter_rows(min_row=2, values_only=True), start=2):
                        # Assuming data starts from row 2
                        if row[description_column_index - 1] and cell_value.lower() in row[
                            description_column_index - 1].lower():
                            count += 1
                            cell_address = main_sheet.cell(row=row_idx, column=description_column_index).coordinate
                            cell_addresses.append(cell_address)

                    # Store the count and cell addresses in the dictionaries if count is greater than zero
                    if count > 0:
                        count_dict[cell_value] = count
                        cell_address_dict[cell_value] = cell_addresses

        # Convert count_dict and cell_address_dict to DataFrames
        count_df = pd.DataFrame.from_dict(count_dict, orient='index', columns=['Count'])
        count_df.index.name = 'Keyword'
        count_df.reset_index(inplace=True)

        address_df = pd.DataFrame.from_dict(cell_address_dict, orient='index')
        address_df.index.name = 'Keyword'
        address_df.columns = [f'Cell Address {i + 1}' for i in range(address_df.shape[1])]
        address_df.reset_index(inplace=True)

        # Sort the count DataFrame in descending order by counts
        sorted_count_df = count_df.sort_values(by='Count', ascending=False)

        # Merge count_df and address_df on 'Keyword' column
        merged_df = pd.merge(sorted_count_df, address_df, on='Keyword')

        # Write the merged DataFrame to the "Not Recommended" sheet
        not_recommended_sheet.cell(row=1, column=1, value="Keyword")
        not_recommended_sheet.cell(row=1, column=2, value="Count")

        for i, column_name in enumerate(merged_df.columns[2:], start=3):
            not_recommended_sheet.cell(row=1, column=i, value=f"Found In {i - 2}")

        for index, row in merged_df.iterrows():
            keyword = row['Keyword']
            count = row['Count']
            cell_addresses = row.drop(['Keyword', 'Count']).dropna().values.tolist()

            # Write the keyword and count to the corresponding columns
            not_recommended_sheet.cell(row=index + 2, column=1, value=keyword)
            not_recommended_sheet.cell(row=index + 2, column=2, value=count)

            # Write the cell addresses as hyperlinks to the corresponding columns
            for i, cell_address in enumerate(cell_addresses, start=1):
                column = get_column_letter(i + 2)
                cell = not_recommended_sheet.cell(row=index + 2, column=2 + i, value=cell_address)
                hyperlink = f'#MAIN!{cell_address}'
                cell.hyperlink = hyperlink

    # Print the contents of the "Not Recommended" sheet for testing
    print("Not Recommended Sheet:")
    for row in not_recommended_sheet.iter_rows(values_only=True):
        print(row)
