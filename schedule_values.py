import tkinter as tk
from tkinter import messagebox, Label, Entry, Button, StringVar
import xlsxwriter
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import NamedStyle, Font, PatternFill, Border, Side, Alignment, numbers

class ScheduleBuilder:
    def __init__(self, filename):
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet()
        self.setup_formats()
        self.setup_headers_and_columns(f'{filename}.xlsx')

    def setup_formats(self):
        self.header_format = self.workbook.add_format({
            'bold': True, 'font_size': 12, 'bg_color': '#C6EFCE', 'border': 1
        })
        self.category_format = self.workbook.add_format({
            'bold': True, 'font_size': 11, 'border': 1
        })
        self.currency_format = self.workbook.add_format({
            'num_format': '$#,##0.00', 'font_size': 11, 'border': 1
        })
        self.total_format = self.workbook.add_format({
            'bold': True, 'bg_color': '#FFEB9C', 'num_format': '$#,##0.00', 'border': 1
        })
        self.negative_format = self.workbook.add_format({
            'font_color': 'red', 'border': 1
        })
        self.default_format = self.workbook.add_format({
            'border': 1
        })
    def write_dataframe_to_sheet(self, df):
        df = df.fillna(0).replace([float('inf'), float('-inf')], 999999)
        # Write headers
        for col_num, value in enumerate(df.columns.values):
            self.worksheet.write(0, col_num, value, self.header_format)
        
        # Write data rows
        for row_num, row in enumerate(df.iterrows()):
            for col_num, value in enumerate(row[1]):
                # Determine the appropriate format
                if isinstance(value, str):
                    cell_format = self.default_format
                elif isinstance(value, (int, float)):
                    if value < 0:
                        cell_format = self.negative_format
                    else:
                        cell_format = self.currency_format if 'cost' in df.columns[col_num].lower() or 'price' in df.columns[col_num].lower() else self.default_format
                else:
                    cell_format = self.default_format

                # Write the data
                self.worksheet.write(row_num + 1, col_num, value, cell_format)

        
    def setup_headers_and_columns(self, name):
        headers = [
            f'{name}', 'Scheduled Value', 'Work Billed from Previous Period', 'Work Billed for this Period',
            'Total billed & Stored to date', '% billed', 'Balance to finish'
        ]
        categories = [
            'General Conditions:', 'Excavation', 'Concrete', 'Masonry and siding',
            'Framing', 'Softit', 'Guttering', 'Roofing', 'Garage doors', 'Windows & exterior doors',
            'Plumbing Install includes water heater', 'Plumbing Fixtures', 'Electric Install',
            'Electric Fixtures', 'HVAC', 'Insulation', 'Drywall', 'Millwork & Trim includes',
            'Cabinets & Vanities', 'Finish Paint', 'Tile, hardwood, carpet', 'Fireplace',
            'Countertops', 'Appliances', 'Golf cart garage', 'Basement', 'IT', 'Overhead/profit', 'Other',
            'Total'
        ]

        # Write header row
        for col, header in enumerate(headers):
            self.worksheet.write(0, col, header, self.header_format)

        # Write category column and initial formatting for rows
        total_row_index = len(categories)  # Index of the 'Total' row
        for row, category in enumerate(categories, start=1):
            self.worksheet.write(row, 0, category, self.category_format)
            for col in range(1, len(headers)):
                if category == 'Total':
                    if headers[col] in {'Scheduled Value', 'Work Billed from Previous Period', 'Work Billed for this Period', 'Total billed & Stored to date', 'Balance to finish'}:
                        # Sum the columns from the second row to the row before the total row
                        cell_range = f'B{2}:B{total_row_index}'
                        self.worksheet.write_formula(row, col, f'=SUM({cell_range.replace("B", chr(65 + col))})', self.total_format)
                else:
                    self.worksheet.write(row, col, '', self.default_format)

        # Set column widths
        self.worksheet.set_column('B:G', 20)


    def set_values(self, values):
        # Only set values for the "Scheduled Value" column
        for row, value in enumerate(values, start=1):  # Start after the header and category name
            self.worksheet.write(row, 1, value, self.currency_format)

        # Apply conditional formatting to "Scheduled Value" column for negative values
        self.worksheet.conditional_format('B2:B31', {
            'type': 'cell', 'criteria': '<', 'value': 0, 'format': self.negative_format
        })

    def close_workbook(self):
        self.workbook.close()


class BudgetApp:
    def __init__(self, master, name):
        self.master = master
        self.master.title("Set Budget Values")
        self.filename = f'{name}'
        self.builder = ScheduleBuilder(self.filename)

        # Prepare the list of categories for inputs from the 'Scheduled Value' column
        # The 'Total' category is not included since it will be a formula in the Excel sheet
        self.categories = [
            'General Conditions:', 'Excavation', 'Concrete', 'Masonry and siding',
            'Framing', 'Softit', 'Guttering', 'Roofing', 'Garage doors', 'Windows & exterior doors',
            'Plumbing Install includes water heater', 'Plumbing Fixtures', 'Electric Install',
            'Electric Fixtures', 'HVAC', 'Insulation', 'Drywall', 'Millwork & Trim includes',
            'Cabinets & Vanities', 'Finish Paint', 'Tile, hardwood, carpet', 'Fireplace',
            'Countertops', 'Appliances', 'Golf cart garage', 'Basement', 'IT', 'Overhead/profit', 'Other'
        ]
        self.entries = []

        # Create entry widgets for each category
        for idx, category in enumerate(self.categories):
            Label(self.master, text=category).grid(row=idx, column=0, sticky='w')
            var = tk.Variable(self.master)
            entry = Entry(self.master, textvariable=var)
            entry.grid(row=idx, column=1)
            self.entries.append(var)

        # Button to submit the budget values
        submit_button = Button(self.master, text="Submit Budget", command=self.submit_budget)
        submit_button.grid(row=len(self.categories) + 1, column=0, columnspan=2)

    

    def submit_budget(self):
        # Collect budget values from entry widgets
        budget_values = []
        try:
            for entry in self.entries:
                # Default to 0.0 if the entry is empty
                value = float(entry.get()) if entry.get() else 0.0
                budget_values.append(value)
            
            # Set values in the Excel sheet
            self.builder.set_values(budget_values)
            
            # Save and close the workbook
            self.builder.close_workbook()
            
            # Confirmation message
            messagebox.showinfo("Success", "Budget values have been successfully submitted.")
        except ValueError:
            messagebox.showerror("Input Error", "Please ensure all entries are numeric.")


import tkinter as tk
from tkinter import Button, messagebox, filedialog
from tkinter.ttk import Frame

class Controller:
    def __init__(self):
        
        self.master = tk.Tk()
        self.master.title("Budget Controller")

        # Label for the name entry
        self.name_label = Label(self.master, text="Enter House Name and Number:")
        self.name_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        # Entry widget for entering the name
        self.name_entry = Entry(self.master, width=25)
        self.name_entry.grid(row=0, column=1, padx=10, pady=10)

        # Button to generate a new budget
        self.generate_budget_button = Button(self.master, text="Generate Budget", command=self.generate_budget)
        self.generate_budget_button.grid(row=1, column=0, padx=10, pady=10)

        # Button to add draws to an existing budget
        self.add_draws_button = Button(self.master, text="Add Draws", command=self.add_draws)
        self.add_draws_button.grid(row=1, column=1, padx=10, pady=10)

    def generate_budget(self):
        name = self.name_entry.get()
        # Placeholder for the actual budget app initialization
        budget_window = tk.Toplevel(self.master)
        BudgetApp(budget_window, name)

    def add_draws(self):
        name = self.name_entry.get()
        # Open a file dialog to select the first existing budget file
        filename1 = filedialog.askopenfilename(
            title="Select the Schedule of Values File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename1:
            # Ask for the second budget file if the first one is selected
            filename2 = filedialog.askopenfilename(
                title="What draw file will you add to it?",
                filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
            )
            if filename2:
                # If two files were selected, proceed with processing both files
                messagebox.showinfo("Files Selected", f"First file: {filename1}\nSecond file: {filename2}")
                self.process_files(filename1, filename2, name)
            else:
                messagebox.showinfo("File Selection", "Second file selection was cancelled.")
        else:
            messagebox.showinfo("File Selection", "First file selection was cancelled.")

    def process_files(self, file1, file2, name):
        ss = ScheduleBuilder(file1)
        pd.set_option('display.max_rows', 500)  # or None to display all rows
        pd.set_option('display.max_columns', 500)  # or None to display all columns
        # Read the data from both files
        df1 = pd.read_excel(file1)
        tempdf2 = pd.read_excel(file2)
        tempdf2.columns = tempdf2.columns.str.strip().str.lower()
        start = self.find_start_row(tempdf2)
        if start is not None:
            df2 = pd.read_excel(file2, skiprows=start + 1)
        else:
            print("Data start row not found")

        oname = name
        df1.drop(29, inplace=True)
        df1.fillna(0, inplace=True)
        df2.fillna(0, inplace=True)
        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()
        name = name.strip().lower()
        print(df2)
        df1['work billed for this period'] = pd.to_numeric(df1['work billed for this period'], errors='coerce')
        df1['work billed from previous period'] = pd.to_numeric(df1['work billed from previous period'], errors='coerce')
        df1['total billed & stored to date'] = pd.to_numeric(df1['total billed & stored to date'], errors='coerce')
        df1['scheduled value'] = pd.to_numeric(df1['scheduled value'], errors='coerce')
        df1['% billed'] = pd.to_numeric(df1['% billed'], errors='coerce')
        df1['balance to finish'] = pd.to_numeric(df1['balance to finish'], errors='coerce')
        df2['amount'] = pd.to_numeric(df2['amount'], errors='coerce')
        
        

        # Check if necessary columns exist
        required_columns = ['code', 'amount']
        missing_columns = [col for col in required_columns if col not in df2.columns]
        if missing_columns:
            messagebox.showerror("Error", f"Missing columns in the file: {', '.join(missing_columns)}")
            return  # Stop processing if columns are missing
        
        # Move current 'Billed this Period' values to 'Billed the Previous Period' and clear 'this period'
        df1['work billed from previous period'] = df1['work billed for this period']
        for i in df1['work billed for this period']:
            i = 0

        # Add dataframes
        for i, j in zip(df2.loc[:, 'code'], df2.loc[:, 'amount']):
            for k in df1[f'{name}.xlsx']:
                if str(i).lower() in str(k).lower():
                    df1.loc[df1[f'{name}.xlsx'].astype(str).str.lower() == str(k).lower(), 'work billed for this period'] = j

                    
        # Update the 'Total Billed' by adding the new 'Billed this Period' to it
        df1['total billed & stored to date'] += df1['work billed for this period'].astype(float)

        # Calculate '% Billed' as the ratio of 'Total Billed' to 'Scheduled Value'
        # Protect against division by zero by using np.where
        
        df1['% billed'] = np.where(df1['scheduled value'] != 0, 
                                   df1['total billed & stored to date'] / df1['scheduled value'] * 100, 
                                   0)

        # Calculate 'Balance to Finish' as 'Scheduled Value' minus 'Total Billed'
        df1['balance to finish'] = df1['scheduled value'] - df1['total billed & stored to date']

         
        totalbudget = 0
        totalprev = 0
        totalthis = 0
        totalbill = 0
        for i, j, k, l in zip(df1['scheduled value'], df1['work billed from previous period'], df1['work billed for this period'], df1['total billed & stored to date']):
            totalbudget += i
            totalprev += j
            totalthis += k
            totalbill += l
        df1.loc[11, 'work billed from previous period'] = round(totalprev, 2)
        df1.loc[11, 'work billed for this period'] = round(totalthis, 2)
        df1.loc[11, 'total billed & stored to date'] = round(totalbill, 2)
        df1.loc[11, '% billed'] = round(100 * (totalbill / totalbudget))
        df1.loc[11, 'balance to finish'] = round(totalbudget - totalbill, 2)
        df1.rename(columns = {name : oname}, inplace=True)
        ss.write_dataframe_to_sheet(df1)
        ss.close_workbook()

        messagebox.showinfo("Process Complete", "Files have been processed and data has been updated.")

    def find_start_row(self, df):
        for index, row in df.iterrows():
            # Check for a specific header name or pattern; adjust condition based on your needs
            if row.str.contains('Code').any() and row.str.contains('Amount').any():
                return index
        return None
