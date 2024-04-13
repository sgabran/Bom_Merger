# ver = '2024-4-12'

import pandas as pd

import filename_methods
import filename_methods as fm
import openpyxl
import math
from constants import *


class SessionProcessXLSX:
    def __init__(self, user_entry, session_log, textbox):
        self.user_entry = user_entry
        self.session_log = session_log
        self.textbox = textbox

        self.rows_to_peak_pd = []
        self.data = []
        self.data_pd = pd.DataFrame()
        self.data_parsed = []
        self.rows_to_peak = []
        self.data_n_rows = 0
        self.data_pd_n_columns = 0

        self.peak_data()
        self.process_data()

    def load_and_peak(self):
        # Read file
        message = "Reading File\n"
        message_colour = "brown"
        self.session_log.write_textbox(message, message_colour)

        # Construct file name
        file_fullname = fm.FileNameMethods.build_file_name_full(self.user_entry.file_location, self.user_entry.file_name, self.user_entry.file_suffix)

        self.data = self.read_excel_file(file_fullname)

        # Get dimensions of data file
        self.data_n_rows = len(self.data)

        self.data_pd = pd.DataFrame(self.data)

        # Peak data
        # Display first rows of file on the textbox
        for index, row in self.data.iterrows():
            self.rows_to_peak.append(row)

        self.rows_to_peak_pd = pd.DataFrame(self.rows_to_peak)
        self.textbox_pandas_update_2(self.rows_to_peak_pd.iloc[0:self.user_entry.n_rows_to_peak])

    def peak_data(self):
        self.load_and_peak()

        message = "Total Data File Rows: " + str(self.data_n_rows) + '\n'
        message_colour = "blue"
        self.session_log.write_textbox(message, message_colour)

        message = "Total Data File Columns: " + str(self.data_pd_n_columns) + '\n'
        message_colour = "blue"
        self.session_log.write_textbox(message, message_colour)

    def process_data(self):
        self.load_and_peak()
        print(">>> component_index", self.user_entry.component_index)
        self.detect_repeated_values(self.data_pd, self.user_entry.component_index)
        print(">>> component_index", self.user_entry.component_index)
        repeated_values = self.sum_column_3_for_repeated_values(self.data_pd, self.user_entry.component_index, self.user_entry.quantity_index)
        merged_data = self.retrieve_rows_for_keys(self.data_pd, repeated_values, self.user_entry.component_index, self.user_entry.quantity_index)

        output_filename = self.user_entry.file_name + FILE_NAME_SAVE
        output_filename_full = filename_methods.FileNameMethods.build_file_name_full(self.user_entry.file_location, output_filename, self.user_entry.file_suffix)
        self.save_dataframe_to_excel(merged_data, output_filename_full)

    ############################################################################################

    def textbox_update(self, data):
        self.textbox.configure(state='normal')
        self.textbox.delete('1.0', 'end')
        self.textbox.insert('end', data)
        self.textbox.configure(state='disabled')

    def textbox_pandas_update(self, df):
        self.textbox.configure(state='normal')
        self.textbox.delete('1.0', 'end')
        for index, row in df.iterrows():
            row_text = "\t".join(str(val) for val in row) + "\n"
            self.textbox.insert('end', row_text)
        self.textbox.configure(state='disabled')

    def textbox_pandas_update_2(self, df):
        self.textbox.configure(state='normal')
        self.textbox.delete('1.0', 'end')
        # Calculate the width of each column
        column_widths = [max(len(str(value)) for value in df[col]) for col in df.columns]

        if max(column_widths) < 100:
            width = 100
            self.textbox.config(width=width)
        else:
            width = math.ceil(max(column_widths) * 3)
            self.textbox.config(width=width)

        # Print column headers
        header = " | ".join(f"{col:{width}}" for col, width in zip(df.columns, column_widths)) + "\n"
        self.textbox.insert('end', header)

        # Print horizontal line
        line = "-" * (sum(column_widths) + len(df.columns) * 3 - 1) + "\n"
        # line = "-" * (width + len(df.columns) * 3 - 1) + "\n"
        self.textbox.insert('end', line)

        # Print rows with vertical lines
        for index, row in df.iterrows():
            row_text = " | ".join(f"{row[col]:{width}}" for col, width in zip(df.columns, column_widths)) + "\n"
            self.textbox.insert('end', row_text)

    @staticmethod
    def read_excel_file(file_path):
        try:
            # Load the workbook
            workbook = openpyxl.load_workbook(file_path)
            # Select the active sheet
            sheet = workbook.active
            # Read all rows into a list
            data = []
            for row in sheet.iter_rows(values_only=True):
                data.append(row)
            df = pd.DataFrame(data)
            return df
        except Exception as e:
            print("Error reading Excel file:", e)
            return None

    @staticmethod
    def detect_repeated_values(df, column_index):
        # Get the first column of the DataFrame
        first_column = df.iloc[:, column_index]
        print(">>> first_column\n", first_column)
        # Find repeated values
        repeated_values = first_column[first_column.duplicated()].unique().tolist()
        print(">>> repeated_values\n", repeated_values)
        return repeated_values

    @staticmethod
    def sum_column_3_for_repeated_values(df, column_index, quantity_index):
        # Check if the specified column index is within the range of column indices
        if column_index < 0 or column_index >= len(df.columns):
            raise ValueError("Column index is out of bounds.")
        # Create a dictionary to store the sums of values in column 3 for each group of repeated values in column 0
        sums_dict = {}
        # Iterate through the specified column of the DataFrame to find repeated values
        for index, value in df.iloc[:, column_index].items():
            if value in sums_dict:
                sums_dict[value] += df.iloc[index, quantity_index]
            else:
                sums_dict[value] = df.iloc[index, quantity_index]

        return sums_dict

    @staticmethod
    def retrieve_rows_for_keys(df, dictionary, column_index, quantity_index):
        # Create an empty list to store the retrieved rows
        retrieved_rows = []

        # Iterate through the keys of the dictionary
        for key in dictionary:
            # Retrieve the first occurrence of the row with the specified value in column 0
            rows = df[df.iloc[:, column_index] == key].copy()
            if not rows.empty:
                # Update the value in column column[quantity_index] with the corresponding value from the dictionary
                rows.iloc[0, quantity_index] = dictionary[key]
                # Selects the first row of the DataFrame
                retrieved_rows.append(rows.iloc[0])

        # Create a DataFrame from the retrieved rows
        data = pd.DataFrame(retrieved_rows)

        return data

    @staticmethod
    def save_dataframe_to_excel(df, file_path):
        df.to_excel(file_path, index=False, header=False)
