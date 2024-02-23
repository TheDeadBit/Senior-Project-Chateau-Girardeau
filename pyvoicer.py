import tkinter as tk
from tkinter import messagebox
from utils.create_elem import GridManager



class GUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(string="PyInvoicer")
        self.root.geometry(newGeometry="800x500")

        tk.Button(master=self.root, text="Create invoice", command=self.invoice).pack()

        self.root.mainloop()
    

    def invoice(self):
        def submit() -> None:
            # needs to be added validating and actual data saving
            messagebox.showinfo(title="Status", message="Invoice saved!")

        
        self.create_invoice_window = tk.Tk()
        self.create_invoice_window.title(string="Invoice")

        label_frame = tk.Frame(master=self.create_invoice_window)
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



GUI()
