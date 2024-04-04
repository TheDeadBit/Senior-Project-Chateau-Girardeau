import tkinter as tk
import pandas as pd
from platform import system
from tkinter import messagebox
from tkinter import simpledialog
from utils.create_elem import GridManager
from os.path import abspath
from os import walk



class GUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(string="PyInvoicer")
        self.root.geometry(newGeometry="800x500")
        self.invoice_data = {
            "House Number": [],
            "Requesting party": [],
            "Date issued": [],
            "Category": [],
            "Shipment Quantity": [],
            "Unit Price": [],
            "PDF file location": []
        }
        self.file_path = ""

        if system() == "Windows":
            self.excel_files_dir = f"{abspath(path='excel files')}\\"
        elif system() == "Linux":
            self.excel_files_dir = f"{abspath(path='excel files')}/"

        tk.Button(master=self.root, text="Create invoice", command=self.invoice_form).pack()
        tk.Button(master=self.root, text="Edit Invoice", command=self.edit_invoice).pack()
        tk.Button(master=self.root, text="Exit", command=self.exit_main).pack()

        self.root.mainloop()
    

    def _reset_dict(self) -> None:
        self.invoice_data = {key: [] for key in self.invoice_data}
    

    def _fill_invoice_data(self) -> None:
        self.invoice_data["House Number"].append(self.lot_number.get())
        self.invoice_data["Requesting party"].append(self.requesting_party.get())
        self.invoice_data["Date issued"].append(self.date_issued.get())
        self.invoice_data["Category"].append(self.category.get())
        self.invoice_data["Shipment Quantity"].append(self.shipment_quantity.get())
        self.invoice_data["Unit Price"].append(self.unit_price.get())


        if not self.pdf_file_location.get() == "":
            self.invoice_data["PDF file location"].append(self.pdf_file_location())
        else:
            self.invoice_data["PDF file location"].append("Not specified")


    def show_info(self) -> None:
        excel_invoice = pd.read_excel(io=self.file_path)

        message = []

        for column_name in excel_invoice:
            message.append(f"{column_name}: {excel_invoice[column_name][0]}\n")

        messagebox.showinfo(title="Invoice current status", message="".join(message))


    def create_invoice(self) -> None:    
        self._fill_invoice_data()

        data_frame = pd.DataFrame(data=self.invoice_data)
        if not self.file_path:
            path = self.excel_files_dir + self.invoice_data['House Number'][0] + ".xlsx"
        else:
            path = self.file_path


        with pd.ExcelWriter(path=path, engine="xlsxwriter") as writer:
            data_frame.to_excel(writer, index=False, sheet_name="Sheet 1")

            for column in data_frame:
                column_length = len(column)
                col_index = data_frame.columns.get_loc(column)
                writer.sheets["Sheet 1"].set_column(col_index, col_index, column_length + 15)
            
        
        self._reset_dict()
        self.create_invoice_window.destroy()
        messagebox.showinfo(title="Status", message="Invoice saved!")


    def edit_invoice(self) -> None:
        lot_number = simpledialog.askstring(title="House Number ", prompt="Please Enter the House Number of invoice: ")
        found_file = False
        lot_numbers = []
        
        for _, _, files in walk(self.excel_files_dir):
            for file in files:
                if file.endswith(".xlsx"):
                    file_lot_number = file.removesuffix(".xlsx").split("_")[1]

                    if file_lot_number == lot_number:
                        found_file = True
                        self.file_path = self.excel_files_dir + file
                        break
                    else:
                        lot_numbers.append(file_lot_number)

            break
        
        if found_file:
            messagebox.showinfo(title="Status", message="Invoice found! Place new values in the new window.")
            self.invoice_form(title=f"Edit Invoice {lot_number}", status = True)
        else:
            messagebox.showerror(title="Status", message="Invoice not found!\nCheck in excel files directory!")


    def invoice_form(self, title = "Invoice", status = False) -> None:
        self.create_invoice_window = tk.Tk()
        self.create_invoice_window.title(string=title)

        label_frame = tk.Frame(master=self.create_invoice_window)
        grid_manager = GridManager(max_cols=2, master=label_frame)

        self.lot_number = tk.Variable(master=label_frame)
        self.requesting_party = tk.Variable(master=label_frame)
        self.date_issued = tk.Variable(master=label_frame)
        self.category = tk.Variable(master=label_frame)
        self.shipment_quantity = tk.Variable(master=label_frame)
        self.unit_price = tk.Variable(master=label_frame)
        self.pdf_file_location = tk.Variable(label_frame)
        
        if title == "Invoice":
            grid_manager.create_label(text="House Number: ")
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

        if status:
            grid_manager.create_button(text="Status", command=self.show_info)

        label_frame.pack(fill='x')


    def exit_main(self) -> None:
        self.root.destroy()


GUI()
