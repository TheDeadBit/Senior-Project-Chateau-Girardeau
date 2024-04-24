from datetime import datetime
import pandas as pd
import xlsxwriter
import numpy as np
from control_request import Request_Controller


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
                                      'Shipment Quantity', 'Unit Price']]
        #invoice_df = self.format_invoice_data(invoice_df)
        #return invoice_df

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


    #def format_invoice_data(self, invoice_df):
        #First calculate total costs for all invoices
        #invoice_df['$ Amount $'] = invoice_df['Shipment Quantity'] * invoice_df['Unit Price']
        
        #Add in blank column for 'Check #' column in excel file
        #invoice_df['Check #'] = np.nan

        #Format columns
        #self.worksheet.write('A7', 'Category', self.subtitle_format)  # Write the title with the subtitle format
        #for i, category in enumerate(invoice_df['Category'], start=7):
            #self.worksheet.write(f'A{i+1}', category)

        
        


        
def main():

    # Create an instance of Request_Control
    control = Request_Controller()
    build = Draw_Request()

    #Path
    direct= r'C:\Users\ehelt\OneDrive\Documents\Draw Request Test'

    #Test
    excel_df = control.__find_files__(direct)
    build.__build_request__(excel_df)

if __name__ == "__main__":
    main()