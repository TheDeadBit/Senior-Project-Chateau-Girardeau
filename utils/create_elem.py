import tkinter as tk



class GridManager(object):
    # constructor which will set up the label frame (master) by max_columns
    def __init__(self, max_cols: int, master: tk.Frame) -> None:
        
        # make all columns weight to 1
        for col in range(max_cols):
            master.columnconfigure(index=col, weight=1)

        # set up row, col and max_cols
        self.row = 0
        self.col = 0
        self.max_cols = max_cols
        
        # current configuration for cell at row,col
        self.cnf = {
            "row": self.row,
            "column": self.col,
            "sticky": tk.W+tk.E
        }

        self.master = master

    # update the config for cell at row,col, with new row and col
    def update_cnf(self, new_row: int, new_column: int) -> None:
        self.cnf.update({
            "row": new_row,
            "column": new_column
        })

    # if the max col is reached
    # reset the col to 0 and increase the row
    # basically add new row and start from the beginning (0 col)
    def check_row_col(self) -> None:
        if self.col == self.max_cols:
            self.row += 1
            self.col = 0
            self.update_cnf(new_row=self.row, new_column=self.col)
    

    # add new cell where text will be entered
    def create_entry(self, text_variable: tk.Variable) -> None:
        # check if new row is needed
        self.check_row_col()
        # create new cell with current configuration
        tk.Entry(master=self.master, textvariable=text_variable).grid(cnf=self.cnf)
        # increment the column and update the configuration
        self.col += 1
        self.update_cnf(new_row=self.row, new_column=self.col)
        
    # add new cell where a lebel will be held
    def create_label(self, text: str) -> None:
        # check if new row is needed
        self.check_row_col()
        # create new cell with current configuration
        tk.Label(master=self.master, text=text).grid(cnf=self.cnf)
        # increment the column and update the configuration
        self.col += 1
        self.update_cnf(new_row=self.row, new_column=self.col)
    
    # add new cell where a button will be held
    def create_button(self, text: str, command) -> None:
        # check if new row is needed
        self.check_row_col()
        # create new cell with current configuration
        tk.Button(master=self.master, text=text, command=command).grid(cnf=self.cnf)
        # increment the column and update the configuration
        self.col += 1
        self.update_cnf(new_row=self.row, new_column=self.col)
