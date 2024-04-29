import tkinter as tk
from invoice_manager import ManageInvoice


class MainMenu(object):
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(string="Main menu")
        self.root.geometry(newGeometry="800x500")

        tk.Button(master=self.root, text="Create invoice", command=ManageInvoice().invoice_form).pack()
        tk.Button(master=self.root, text="Edit Invoice", command=ManageInvoice().edit_invoice).pack()
        tk.Button(master=self.root, text="Generate Draw Request").pack()
        tk.Button(master=self.root, text="Generate Schedule of Values").pack()
        tk.Button(master=self.root, text="Exit", command=self.exit_button).pack()

        self.root.mainloop()
    
    def exit_button(self):
        self.root.destroy()


MainMenu()
