import tkinter as tk
import pandas as pd
from platform import system
from tkinter import messagebox
from tkinter import simpledialog
from utils.create_elem import GridManager
from os.path import abspath
from os import walk



class CreateInvoice(object):
    def __init__(self) -> None:
        self.invoice_data = {
            "House Number": [],
            "Requesting party": [],
            "Date issued": [],
            "Category": [],
            "Shipment Quantity": [],
            "Unit Price": [],
            "PDF file location": []
        }
        # the path for excel file which is to be edited
        self.file_path = ""

        # determine if the environment is Linux or Windows
        # setup the full path for the excel files directory
        if system() == "Windows":
            self.excel_files_dir = f"{abspath(path='excel files')}\\"
        elif system() == "Linux":
            self.excel_files_dir = f"{abspath(path='excel files')}/"
    

    # basically empty the dictionary
    # keep the keys but clear the lists
    def _reset_dict(self) -> None:
        self.invoice_data = {key: [] for key in self.invoice_data}
    

    # get the data from the user and fill the dictionary
    def _fill_invoice_data(self) -> None:
        self.invoice_data["House Number"].append(self.house_number.get())
        self.invoice_data["Requesting party"].append(self.requesting_party.get())
        self.invoice_data["Date issued"].append(self.date_issued.get())
        self.invoice_data["Category"].append(self.category.get())
        self.invoice_data["Shipment Quantity"].append(self.shipment_quantity.get())
        self.invoice_data["Unit Price"].append(self.unit_price.get())


        if not self.pdf_file_location.get() == "":
            self.invoice_data["PDF file location"].append(self.pdf_file_location())
        else:
            self.invoice_data["PDF file location"].append("Not specified")

    # if user is editng a invoice
    # show status of the current invoice
    def show_info(self) -> None:
        # read the excel file which is to be edited
        excel_invoice = pd.read_excel(io=self.file_path)

        message = []


        for column_name in excel_invoice:
            message.append(f"{column_name}: {excel_invoice[column_name][0]}\n")

        messagebox.showinfo(title="Invoice current status", message="".join(message))


    # create a new invoice
    def create_invoice(self) -> None:
        # get the data from the user    
        self._fill_invoice_data()

        # create new data frame
        data_frame = pd.DataFrame(data=self.invoice_data)

        # check if we are trying to modify existing file
        # if no then generate path for the new file
        # is yes then use the file path
        if not self.file_path:
            path = self.excel_files_dir + self.invoice_data['House Number'][0] + ".xlsx"
        else:
            path = self.file_path

        
        # open the file for editing
        with pd.ExcelWriter(path=path, engine="xlsxwriter") as writer:
            # write the data frame to the excel file in sheet named Sheet 1
            data_frame.to_excel(writer, index=False, sheet_name="Sheet 1")

            # fix the length of the columns
            for column in data_frame:
                column_length = len(column)
                col_index = data_frame.columns.get_loc(column)
                writer.sheets["Sheet 1"].set_column(col_index, col_index, column_length + 15)
            
        
        # clear the invoice data 
        self._reset_dict()
        # destroy the invoice window
        self.invoice_window.destroy()
        # display message
        messagebox.showinfo(title="Status", message="Invoice saved!")


    # edit invoice function
    def edit_invoice(self) -> None:
        # get the house number from the user
        house_number = simpledialog.askstring(title="House Number ", prompt="Please Enter the House Number of invoice: ")

        # if the user clicked cancel or has not provided house number
        # exit the function
        if house_number == None:
            return
        
        # bool flag to determine if file is found or not
        found_file = False
        # list for all valid house numbers 
        house_numbers = []
        
        for _, _, files in walk(self.excel_files_dir):
            for file in files:
                # check if the file is excel
                if file.endswith(".xlsx"):
                    # get the house number
                    file_house_number = file.removesuffix(".xlsx").split("_")[1]

                    # compare the file house number to the wanted house number
                    # if they match, change flag to true and setup the file path var
                    # if they do not match, add the file house number to the house_numbers
                    if file_house_number == house_number:
                        found_file = True
                        self.file_path = self.excel_files_dir + file
                        break
                    else:
                        house_numbers.append(file_house_number)

            break

        # if the wanted file is found
        # diplay message, and open an invoice form with the data contained in the file
        # if the file is not found, display message
        if found_file:
            messagebox.showinfo(title="Status", message="Invoice found! Place new values in the new window.")
            self.invoice_form(title=f"Edit Invoice {house_number}", status = True)
        else:
            messagebox.showerror(title="Status", message="Invoice not found!\nCheck in excel files directory!")


    # inform voice
    def invoice_form(self, title = "Create Invoice", status = False) -> None:
        # create invoice window
        self.invoice_window = tk.Tk()
        self.invoice_window.title(string=title)

        # create label frame
        label_frame = tk.Frame(master=self.invoice_window)
        # create grid manager, which will manage the frame
        # the max columns will be 2, and the master is the label frame
        grid_manager = GridManager(max_cols=2, master=label_frame)

        # create tk varibles
        self.house_number = tk.Variable(master=label_frame)
        self.requesting_party = tk.Variable(master=label_frame)
        self.date_issued = tk.Variable(master=label_frame)
        self.category = tk.Variable(master=label_frame)
        self.shipment_quantity = tk.Variable(master=label_frame)
        self.unit_price = tk.Variable(master=label_frame)
        self.pdf_file_location = tk.Variable(label_frame)
        
        # if the user creates invoice
        # add house number label
        if title == "Create Invoice":
            grid_manager.create_label(text="House Number: ")
            grid_manager.create_entry(text_variable=self.house_number)


        # creat new labels and entrys
        
        grid_manager.create_label(text="Requesting party: ")
        grid_manager.create_entry(text_variable=self.requesting_party)

        grid_manager.create_label(text="Date issued: ")
        grid_manager.create_entry(text_variable=self.date_issued)

        grid_manager.create_label(text="Category: ")
        grid_manager.create_entry(text_variable=self.category)

        grid_manager.create_label(text="Shipment Quantity: ")
        grid_manager.create_entry(text_variable=self.shipment_quantity)

        grid_manager.create_label(text="Unit Price: ")
        grid_manager.create_entry(text_variable=self.unit_price)

        grid_manager.create_label(text="PDF file location")
        grid_manager.create_entry(text_variable=self.pdf_file_location)


        grid_manager.create_button(text="Submit", command=self.create_invoice)

        # if the user edits invoice, display current status
        if status:
            grid_manager.create_button(text="Status", command=self.show_info)

        # fill empty spaces
        label_frame.pack(fill='x')



CreateInvoice()
