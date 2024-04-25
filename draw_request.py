import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xlsxwriter
import numpy as np
from datetime import datetime

##CONTROLLER HELPER FUNCTIONS
def __find_files__(self, direct):

    #Calculate time frame for invoices
    end_date = datetime.datetime.now()
    start_date = end_date - datetime.timedelta(days=7)

    #List of all files in provided directory
    all_files = os.listdir(direct)

    #List of excel files
    excel_files = []
    
    #Filter files based on their type and modification time within the specified time frame
    for file in all_files:
        file_path = os.path.join(direct, file)
        if file.endswith('.xlsx') and os.path.getmtime(file_path) >= start_date.timestamp():
            excel_files.append(file_path)

    # Read all Excel files into a list of DataFrames
    excel_list = [pd.read_excel(file) for file in excel_files]
    excel_df = pd.concat(excel_list, ignore_index=True)
    return excel_df

##MODEL CLASS
class Draw_Request:
    def __init__(self):
        self.workbook = xlsxwriter.Workbook('sample.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.subtitle_format = self.workbook.add_format({'bold': True, 'font_size': 12})
        self.info_format = self.workbook.add_format({'font_size': 12})

    def __build_request__(self, excel_df):
        title_data = {'Customer': ['Ramsey Run'], 'Draw Number': [0],
                      'Property': [excel_df.iloc[0, excel_df.columns.get_loc('House Number')]],
                      'Date': [datetime.now()]}
        title_df = pd.DataFrame(title_data)
        self.format_title(title_df)

        invoice_df = excel_df.loc[:, ['Requesting Party', 'Category',
                                      'Shipment Quantity', 'Unit Price', 'Date Issued']]
        self.format_invoice_data(invoice_df)

    def format_title(self, title_df):

        title_format = self.workbook.add_format({'bold': True, 'font_size': 14})
        self.worksheet.merge_range('B1:E1', 'Construction Draw Request Form', title_format)
        self.worksheet.write('B3', 'Customer:', self.subtitle_format)
        self.worksheet.write('C3', title_df['Customer'][0], self.info_format)
        self.worksheet.write('D3', 'Draw #:', self.subtitle_format)
        self.worksheet.write('E3', str(title_df['Draw Number'][0]), self.info_format)
        self.worksheet.write('B4', 'Property:', self.subtitle_format)
        self.worksheet.write('C4', title_df['Property'][0], self.info_format)
        self.worksheet.write('D4', 'Date:', self.subtitle_format)
        self.worksheet.write('E4', title_df['Date'][0].strftime('%m/%d/%Y'), self.info_format)
        self.worksheet.set_column('B:E', 20)


    def format_invoice_data(self, invoice_df):
        #First calculate total costs for all invoices
        invoice_df['$ Amount $'] = (invoice_df['Shipment Quantity'] * invoice_df['Unit Price']).round(2)
        
        #Add in blank column for 'Check #' column in excel file
        invoice_df['Check #'] = np.nan

        #Format columns
        self.worksheet.write('A7', 'Category', self.subtitle_format)
        for i, category in enumerate(invoice_df['Category'], start=7):
            self.worksheet.write(f'A{i+1}', category)
        
        self.worksheet.write('B7', 'Requesting Party', self.subtitle_format)
        for i, category in enumerate(invoice_df['Requesting Party'], start=7):
            self.worksheet.write(f'B{i+1}', category)
        
        self.worksheet.write('C7', 'Date Issued', self.subtitle_format)
        date_format = self.workbook.add_format({'num_format': 'mm/dd/yyyy'})
        for i, date_issued in enumerate(invoice_df['Date Issued'], start=7):
            self.worksheet.write_datetime(f'C{i+1}', date_issued, date_format)

        self.worksheet.write('D7', 'Check #', self.subtitle_format) 
        for i in range(len(invoice_df['Check #'])):
            self.worksheet.write_blank(f'D{i+8}', None, self.subtitle_format) 

        self.worksheet.write('E7', '$ Amount $', self.subtitle_format) 
        for i, category in enumerate(invoice_df['$ Amount $'], start=7):
            self.worksheet.write(f'E{i+1}', category)

        self.workbook.close()

##VIEWER CLASS
class Request_Viewer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("File Selection")

    def __select__direct__(self):
        file_dir = filedialog.askdirectory()
        if file_dir:
            print("File directory: ", file_dir)

    def show_window(self):
        #Format window
        window = tk.Toplevel(self.root)
        window.title("Select Files")
        label = tk.Label(window, text = "Please select a directory for Draw Request")
        label.pack(padx = 20, pady = 20)
        
        #Enter button
        enter_button = tk.Button(window, text = "Select Directory", command = self.__select__direct__)
        enter_button.pack(pady = 10)
        
        #Cancel button
        cancel_button = tk.Button(window, text="Cancel", command=lambda: self.close_window(window))
        cancel_button.pack(pady = 5)
    
    def close_window(self, window):
        window.destory()

#TESTING
def main():
    view = Request_Viewer()
    root = view.root
    start = tk.Button(root, text="TESTING", command = view.show_window)
    start.pack(padx=20, pady=10)
    view.root.mainloop()

if __name__ == "__main__":
    main()
