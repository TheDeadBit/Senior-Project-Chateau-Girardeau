import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xlsxwriter
import numpy as np
import datetime

##MODEL CLASS
class Draw_Request:
    def __init__(self, output_direct, file_name, drawreq_num):
        self.direct = output_direct
        #Create excel workbook in directory
        self.workbook = xlsxwriter.Workbook(os.path.join(self.direct, file_name))
        self.worksheet = self.workbook.add_worksheet()
        self.subtitle_format = self.workbook.add_format({'bold': True, 'font_size': 12})
        self.info_format = self.workbook.add_format({'font_size': 12})
        self.draw_num = drawreq_num

    def __build_request__(self, excel_df):
        title_data = {'Customer': ['Ramsey Run'], 'Draw Number': [self.draw_num],
                      'Property': [excel_df.iloc[0, excel_df.columns.get_loc('House Number')]],
                      'Date': [datetime.datetime.now()]}
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
        for i, amount in enumerate(invoice_df['$ Amount $'], start=7):
            self.worksheet.write(f'E{i+1}', amount)

        self.workbook.close()

##VIEWER CLASS
class Request_Viewer(object):
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Generate Draw Request")
        self.root.geometry("800x500")
        self.entry = None
        self.elements = []
    
    def main_window(self) -> None:
        # Select directory
        label = tk.Label(self.root, text="Please select directory for invoice files:")
        self.elements.append(label)
        label.pack()

        # Enter button
        enter = tk.Button(self.root, text="Select Directory", command=self.select_invoice)
        self.elements.append(enter)
        enter.pack()

        # Cancel button
        cancel = tk.Button(self.root, text="Cancel", command=self.close_window)
        self.elements.append(cancel)
        cancel.pack()

        self.root.mainloop()

    def exit(self) -> None:
        #Exit
        label = tk.Label(self.root, text = "Draw Request generated!")
        self.elements.append(label)
        label.pack()

        button = tk.Button(self.root, text="Return to Main Menu", command=self.close_window)
        self.elements.append(button)
        button.pack()
    
    def get_filename_window(self) -> None:
        #Ask for filename
        label = tk.Label(self.root, text = "Save as?")
        self.elements.append(label)
        label.pack()

        #Create entry point for input
        self.entry = tk.Entry(self.root)
        self.elements.append(self.entry)
        self.entry.pack()

        button = tk.Button(self.root, text = "Enter", command = self.get_filename)
        self.elements.append(button)
        button.pack()

    def select_invoice(self) -> None:
        #Receives invoice files
        self.invoice_direct = filedialog.askdirectory()
        self.clear_gui()

        #Next part of the program
        #Directory that draw request will be saved to
        label = tk.Label(self.root, text="Save Draw Request to?")
        self.elements.append(label)
        label.pack()
        
        button = tk.Button(self.root, text = "Select Folder", command = self.select_drawreq)
        self.elements.append(button)
        button.pack()

    def select_drawreq(self) -> None:
        self.drawreq_direct = filedialog.askdirectory()
        self.clear_gui()

        #Ask for draw req number
        drawnum_label = tk.Label(self.root, text="Draw Request Number?")
        self.elements.append(drawnum_label)
        drawnum_label.pack()
        
        #Create entry point for input
        self.entry = tk.Entry(self.root)
        self.elements.append(self.entry)
        self.entry.pack()

        #Enter button
        enter = tk.Button(self.root, text = "Enter", command = self.get_draw_num)
        self.elements.append(enter)
        enter.pack()

    def close_window(self):
        self.root.destroy()
    
    #Clears GUI Grid
    def clear_gui(self) -> None:
        for element in self.elements:
            element.destroy()
    
    #Close window (exit)
    def close_window(self) -> None:
        self.root.destroy()
    
    #Gets draw number
    def get_draw_num(self) -> None:
        self.drawreq_num = self.entry.get()
        self.clear_gui()
        self.entry = None
        self.get_filename_window()

    #Gets filename
    def get_filename(self) -> None:
        self.filename = self.entry.get()
        self.clear_gui()
        self.exit()

##CONTROLLER CLASS
class Control_Request(object):
    def __init__(self) -> None:
        pass
    
    def driver(self) -> None:
        view = Request_Viewer()
        view.main_window()

        #Read excel files to pd dataframe          
        excel_files = self.__find_files__(view.invoice_direct)

        #Analyze data using Draw_Request() model class
        drawreq_name = str(view.filename) + '.xlsx'
        model = Draw_Request(view.drawreq_direct, drawreq_name, view.drawreq_num)
        model.__build_request__(excel_files)


    def __find_files__(self, direct) -> None:
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

Control_Request().driver()