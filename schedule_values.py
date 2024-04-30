import tkinter as tk
from tkinter import messagebox, Label, Entry, Button, StringVar
import xlsxwriter
import pandas as pd
import numpy as np
import openpyxl

class ScheduleBuilder:
    def __init__(self, filename):
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet()
        self.setup_formats()
        self.setup_headers_and_columns('Ramsey Run 9')

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
        for row, category in enumerate(categories, start=1):
            self.worksheet.write(row, 0, category, self.category_format)
            for col in range(1, len(headers)):
                if category == 'Total' and col == 1:  # Total row, only under 'Scheduled Value'
                    self.worksheet.write(row, col, '=SUM(B2:B{0})'.format(row), self.total_format)
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
    def __init__(self, master, filename="schedule.xlsx"):
        self.master = master
        self.master.title("Set Budget Values")
        self.filename = filename
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
            var = StringVar()
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

        # Button to generate a new budget
        self.generate_budget_button = Button(master=self.master, text="Generate Budget", command=self.generate_budget)
        self.generate_budget_button.grid(row=0, column=0, padx=10, pady=10)

        # Button to add draws to an existing budget
        self.add_draws_button = Button(master=self.master, text="Add Draws", command=self.add_draws)
        self.add_draws_button.grid(row=0, column=1, padx=10, pady=10)

    def generate_budget(self):
        # Placeholder for the actual budget app initialization
        budget_window = tk.Toplevel(self.master)
        BudgetApp(budget_window)

    def add_draws(self):
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
                self.process_files(filename1, filename2, 'Ramsey Run 9')
            else:
                messagebox.showinfo("File Selection", "Second file selection was cancelled.")
        else:
            messagebox.showinfo("File Selection", "First file selection was cancelled.")

    def process_files(self, file1, file2, name):
        # Read the data from both files
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        print(df1.head())
        print(df2.head())
        df1['Work Billed for this Period'] = pd.to_numeric(df1['Work Billed for this Period'], errors='coerce')
        df1['Work Billed from Previous Period'] = pd.to_numeric(df1['Work Billed from Previous Period'], errors='coerce')
        df1['Total billed & Stored to date'] = pd.to_numeric(df1['Total billed & Stored to date'], errors='coerce')
        df1['Scheduled Value'] = pd.to_numeric(df1['Scheduled Value'], errors='coerce')
        df1['% billed'] = pd.to_numeric(df1['% billed'], errors='coerce')
        df1['Balance to finish'] = pd.to_numeric(df1['Balance to finish'], errors='coerce')
        df2['Amount'] = pd.to_numeric(df2['Amount'], errors='coerce')
        # Standardize column names by stripping whitespace and converting to a consistent case
        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()
        name = name.strip().lower()

        # Check if necessary columns exist
        required_columns = ['code', 'amount']
        missing_columns = [col for col in required_columns if col not in df2.columns]
        if missing_columns:
            messagebox.showerror("Error", f"Missing columns in the file: {', '.join(missing_columns)}")
            return  # Stop processing if columns are missing

        # Ensure the necessary columns exist in df1 to avoid KeyError
        for column in ['workbilledforthisperiod', 'workbilledfrompreviousperiod', 'Total Billed', 'Scheduled Value']:
            if column not in df1.columns:
                df1[column] = 0
        print(df1.head())
        # Move current 'Billed this Period' values to 'Billed the Previous Period'
        df1['workbilledfrompreviousperiod'] = df1['workbilledforthisperiod']

        # Merge df1 with df2 based on 'Code' to update 'Billed this Period'
        # This merge assumes df2's Amount is the new data for 'Billed this Period'
        df1 = df1.merge(df2, left_on=f'{name}', right_on='code', how='left')

        # Update 'Billed this Period' with new Amount from df2, and fill NaN with 0 if no match was found
        df1['workbilledforthisperiod'] = df1[f'{name}'].fillna(0)

        # Update the 'Total Billed' by adding the new 'Billed this Period' to it
        df1['Total Billed'] += df1['workbilledforthisperiod']

        # Calculate '% Billed' as the ratio of 'Total Billed' to 'Scheduled Value'
        # Protect against division by zero by using np.where
        import numpy as np
        df1['% billed'] = np.where(df1['Scheduled Value'] != 0, 
                                   df1['Total billed & Stored to date'] / df1['Scheduled Value'] * 100, 
                                   0)

        # Calculate 'Balance to Finish' as 'Scheduled Value' minus 'Total Billed'
        df1['Balance to finish'] = df1['Scheduled Value'] - df1['Total billed & Stored to date']

        # Drop the extra 'Amount' column brought in by the merge
        df1.drop(columns=[f'{name}'], inplace=True)

        # Save the updated DataFrame back to the first file
        df1.to_excel(file1, index=False)

        messagebox.showinfo("Process Complete", "Files have been processed and data has been updated.")



