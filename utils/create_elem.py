import tkinter as tk



class GridManager:
    def __init__(self, max_cols: int, master: tk.Frame) -> None:
        for col in range(max_cols):
            master.columnconfigure(index=col, weight=1)

        self.row = 0
        self.col = 0
        self.max_cols = max_cols
        self.cnf = {
            "row": self.row,
            "column": self.col,
            "sticky": tk.W+tk.E
        }
        self.master = master

    def update_cnf(self, new_row: int, new_column: int) -> None:
        self.cnf.update({
            "row": new_row,
            "column": new_column
        })


    def check_row_col(self) -> None:
        if self.col == self.max_cols:
            self.row += 1
            self.col = 0
            self.update_cnf(new_row=self.row, new_column=self.col)
    

    def create_entry(self, textvariable: tk.Variable | None) -> None:
        self.check_row_col()
        tk.Entry(master=self.master, textvariable=textvariable).grid(cnf=self.cnf)
        self.col += 1
        self.update_cnf(new_row=self.row, new_column=self.col)
        
    
    def create_label(self, text: str | None) -> None:
        self.check_row_col()
        tk.Label(master=self.master, text=text).grid(cnf=self.cnf)
        self.col += 1
        self.update_cnf(new_row=self.row, new_column=self.col)
    

    def create_button(self, text: str, command) -> None:
        self.check_row_col()
        tk.Button(master=self.master, text=text, command=command).grid(cnf=self.cnf)
        self.col += 1
        self.update_cnf(new_row=self.row, new_column=self.col)
