import datetime
import os

import pandas as pd

class Request_Controller:

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