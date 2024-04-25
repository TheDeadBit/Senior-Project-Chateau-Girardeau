import tkinter as tk
from pyvoicer import GUI


class MainMenu(object):
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(string="Main menu")
        self.root.geometry(newGeometry="800x500")

        tk.Button(master=self.root, text="Create invoice").pack()
        tk.Button(master=self.root, text="Edit Invoice").pack()
        tk.Button(master=self.root, text="Draw Request").pack()
        tk.Button(master=self.root, text="Schedule Values").pack()
        tk.Button(master=self.root, text="Exit", command=self.exit_button).pack()

        self.root.mainloop()
    
    def exit_button(self):
        self.root.destroy()


MainMenu()
