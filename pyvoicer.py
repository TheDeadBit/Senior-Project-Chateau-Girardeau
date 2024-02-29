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

        tk.Button(master=self.root, text="Create invoice", command=self.invoice).pack()
        tk.Button(master=self.root, text="Exit", command=self.exit_main).pack()

        self.root.mainloop()
    

    def invoice(self) -> None:
        def submit() -> None:
            data = {
                "Lot number": [lot_number.get()],
                "Requesting party": [requesting_party.get()],
                "Date issued": [date_issued.get()],
                "Category": [category.get()],
                "Shipment Quantity": [shipment_quantity.get()],
                "Unit price": [unit_price.get()],
            }
            
            pdf = pdf_file_location.get()

            if not pdf == "":
                data["PDF file location"] = [pdf]

            data_frame = pd.DataFrame(data=data)
            path = ""

            if system() == "Windows":
                path = f"{abspath(path='excel files')}\\invoice_{lot_number.get()}.xlsx"
            elif system() == "Linux":
                path = f"{abspath(path='excel files')}/invoice_{lot_number.get()}.xlsx"

            with pd.ExcelWriter(path=path, engine="xlsxwriter") as writer:
                data_frame.to_excel(writer, index=False, sheet_name="Sheet 1")

                for column in data_frame:
                    column_length = len(column)
                    col_index = data_frame.columns.get_loc(column)
                    writer.sheets["Sheet 1"].set_column(col_index, col_index, column_length + 15)
            
            
            create_invoice_window.destroy()
            messagebox.showinfo(title="Status", message="Invoice saved!")

        
        create_invoice_window = tk.Tk()
        create_invoice_window.title(string="Invoice")

        label_frame = tk.Frame(master=create_invoice_window)
        grid_manager = GridManager(max_cols=2, master=label_frame)

        lot_number = tk.Variable(master=label_frame)
        requesting_party = tk.Variable(master=label_frame)
        date_issued = tk.Variable(master=label_frame)
        category = tk.Variable(master=label_frame)
        shipment_quantity = tk.Variable(master=label_frame)
        unit_price = tk.Variable(master=label_frame)
        pdf_file_location = tk.Variable(label_frame)
        
        grid_manager.create_label(text="Lot number: ")
        grid_manager.create_entry(textvariable=lot_number)

        grid_manager.create_label(text="Requesting party: ")
        grid_manager.create_entry(textvariable=requesting_party)

        grid_manager.create_label(text="Date issued: ")
        grid_manager.create_entry(textvariable=date_issued)

        grid_manager.create_label(text="Category: ")
        grid_manager.create_entry(textvariable=category)

        grid_manager.create_label(text="Shipment Quantity: ")
        grid_manager.create_entry(textvariable=shipment_quantity)

        grid_manager.create_label(text="Unit Price: ")
        grid_manager.create_entry(textvariable=unit_price)

        grid_manager.create_label(text="PDF file location")
        grid_manager.create_entry(textvariable=pdf_file_location)

        grid_manager.create_button(text="Submit", command=submit)

        label_frame.pack(fill='x')


    def exit_main(self) -> None:
        self.root.destroy()

GUI()
