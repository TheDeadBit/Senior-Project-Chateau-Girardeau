import tkinter as tk
import pandas as pd
from platform import system
from tkinter import messagebox
from utils.create_elem import GridManager
from os.path import abspath



class GUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(string="PyInvoicer")
        self.root.geometry(newGeometry="800x500")
        self.invoice_data = {
            "Lot number": [],
            "Requesting party": [],
            "Date issued": [],
            "Category": [],
            "Shipment Quantity": [],
            "Unit Price": [],
            "PDF file location": []
        }

        tk.Button(master=self.root, text="Create invoice", command=self.invoice_form).pack()
        tk.Button(master=self.root, text="Edit Invoice", command=self.edit_invoice).pack()
        tk.Button(master=self.root, text="Exit", command=self.exit_main).pack()

        self.root.mainloop()
    

    def invoice_form(self) -> None:
        self.create_invoice_window = tk.Tk()
        self.create_invoice_window.title(string="Invoice")

        label_frame = tk.Frame(master=self.create_invoice_window)
        grid_manager = GridManager(max_cols=2, master=label_frame)

        self.lot_number = tk.Variable(master=label_frame)
        self.requesting_party = tk.Variable(master=label_frame)
        self.date_issued = tk.Variable(master=label_frame)
        self.category = tk.Variable(master=label_frame)
        self.shipment_quantity = tk.Variable(master=label_frame)
        self.unit_price = tk.Variable(master=label_frame)
        self.pdf_file_location = tk.Variable(label_frame)
        
        grid_manager.create_label(text="Lot number: ")
        grid_manager.create_entry(text_variable=self.lot_number)

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

        label_frame.pack(fill='x')
        print()
    

    def create_invoice(self) -> None:    
        self.invoice_data["Lot number"].append(self.lot_number.get())
        self.invoice_data["Requesting party"].append(self.requesting_party.get())
        self.invoice_data["Date issued"].append(self.date_issued.get())
        self.invoice_data["Category"].append(self.category.get())
        self.invoice_data["Shipment Quantity"].append(self.shipment_quantity.get())
        self.invoice_data["Unit Price"].append(self.unit_price.get())


        if not self.pdf_file_location.get() == "":
            self.invoice_data["PDF file location"].append(self.pdf_file_location())
        else:
            self.invoice_data["PDF file location"].append("Not specified")


        data_frame = pd.DataFrame(data=self.invoice_data)
        path = ""

        if system() == "Windows":
            path = f"{abspath(path='excel files')}\\invoice_{self.invoice_data['Lot number'][0]}.xlsx"
        elif system() == "Linux":
            path = f"{abspath(path='excel files')}/invoice_{self.invoice_data['Lot number'][0]}.xlsx"

        with pd.ExcelWriter(path=path, engine="xlsxwriter") as writer:
            data_frame.to_excel(writer, index=False, sheet_name="Sheet 1")

            for column in data_frame:
                column_length = len(column)
                col_index = data_frame.columns.get_loc(column)
                writer.sheets["Sheet 1"].set_column(col_index, col_index, column_length + 15)
            
            
        self.create_invoice_window.destroy()
        messagebox.showinfo(title="Status", message="Invoice saved!")

        
        


    def edit_invoice(self) -> None:
        # ask the user for invoice lot number
        # look inside excel files directory for the correct invoice
        # read the data
        # load the data inside window
        pass


    def exit_main(self) -> None:
        self.root.destroy()

GUI()
