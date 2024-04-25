from datetime import datetime
import pandas as pd
import xlsxwriter
import numpy as np


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

        
        