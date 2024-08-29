import pandas as pd
import os
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import matplotlib.pyplot as plt

#### Main Class ####
class App:
    def __init__(self, master):
        self.master = master
        self.main_page()

    ### Functions for creating the GUI and its pages ###

    ## Main page function
    def main_page(self):
        # Destroying all widgets from the previous page (when returned into from another page)
        for i in self.master.winfo_children():
            i.destroy()

        # title and geometry
        self.master.title("Excel Master")
        self.master.geometry("280x250")
        
        # greeting message and menu buttons
        tk.Label(self.master, text="Welcome to Excel Master!", font=("Times","24"), fg="#5ea832").grid(row=0, column=0, padx=10, pady=10)
        tk.Label(self.master, text="Choose what you want to do:").grid(row=1, column=0, padx=10)

        # Excel cleaner site button
        self.clean_button = tk.Button(self.master, text="Excel Cleaner", command=self.clean_page, activeforeground="blue")
        self.clean_button.grid(row=2, column=0, padx=10, pady=10)

        # Excel pivot creator site button
        self.pivot_button = tk.Button(self.master, text="Excel Pivot Creator", command = None, activeforeground="blue")
        self.pivot_button.grid(row=3, column=0, padx=10, pady=10)

        # Excel plotter site button
        self.plotter_button = tk.Button(self.master, text="Excel Plotter", command = None, activeforeground="blue")
        self.plotter_button.grid(row=4, column=0, padx=10, pady=10)

        # Trademark
        tk.Label(self.master, text="© Made by Peter Peško, 2024", font=("Times","12"), fg="#d4d4d4").grid(row=5, column=0, padx=10, sticky="w")

    ## Cleaner page function
    def clean_page(self):
        # Destroying all widgets from the main page
        for i in self.master.winfo_children():
            i.destroy()
        self.master.geometry("520x240")
        
        # Back button
        back_button = tk.Button(self.master, text="<<<", command=self.main_page, cursor="hand2", activeforeground="blue")
        back_button.grid(row=0, column=0, padx=10, sticky="w")

        # Cleaner page title
        tk.Label(self.master, text = "Excel Cleaner", font=("Times","20"), fg="#5ea832").grid(row=1, column=1, padx=10, sticky="w")

        # Choosing an Excel file with button
        self.filelabel = tk.Label(self.master, text="Choose Excel file:", font=("Times","15"), fg="#5ea832")
        self.filelabel.grid(row=2, column=0, padx=5, sticky="w")
        self.file = tk.Entry(self.master, state="disabled", bg="#74e8f2")
        self.file.grid(row=2, column=1, sticky="w")
        file_button = tk.Button(self.master, text="Choose .xlsx file", command=self.choose_file, activeforeground="blue")
        file_button.grid(row=2, column=2, sticky="w")

        # Specifying if there is an index in the excell file
        tk.Label(self.master, text="Does your file have index?", font=("Times","15"), fg="#5ea832").grid(row=3, column=0, padx=5, sticky="w")
        self.index_check = tk.IntVar()
        self.index = tk.Checkbutton(self.master, variable=self.index_check, text="(no)", onvalue=1, offvalue=0, activeforeground="blue", command=self.toggle_index)
        self.index.grid(row=3, column=1, sticky="w")
        self.index_col_label = tk.Label(self.master, text="Column (0,...):", font=("Times","15"), fg="#5ea832")
        self.index_col = tk.Entry(self.master, width=5)
    
        # Choosing preferred output file in a dropdown menu
        tk.Label(self.master, text="Preferred output:", font=("Times","15"), fg="#5ea832").grid(row=4, column=0, padx=5,sticky="w")
        self.n = tk.StringVar()
        self.output_choice = ttk.Combobox(self.master, state="readonly", values=[".xlsx", ".csv", ".sqlite"], textvariable=self.n, width=10)
        self.output_choice.bind("<<ComboboxSelected>>", self.toggle_sqlite_choice)
        self.output_choice.grid(row=4, column=1, sticky="w")
        self.sqllabel = tk.Label(self.master, text="Enter table name:", font=("Times","15"), fg="#5ea832")
        self.tablename = tk.Entry(self.master, width=10)
        self.output_choice.current(0)
        
        # Cleaning options
        tk.Label(self.master, text="Cleaning options:", font=("Times","15"), fg="#5ea832").grid(row=5, column=0, padx=5,sticky="w")
        self.clean_option1 = tk.IntVar()
        self.clean_option2 = tk.IntVar()
        checkbox1 = tk.Checkbutton(self.master, text="Drop Duplicates", onvalue=1, offvalue=0, variable=self.clean_option1)
        checkbox2 = tk.Checkbutton(self.master, text="Drop empty rows", onvalue=1, offvalue=0, variable=self.clean_option2, command=self.toggle_how)
        checkbox1.grid(row=5, column=1, sticky="w")
        checkbox2.grid(row=6, column=1, sticky="w")
        
        # DropNA How section
        self.howlabel = tk.Label(self.master, text="How to filter?", font=("Times","15"), fg="#5ea832")
        self.how = tk.StringVar()
        self.checkbox2_how = ttk.Combobox(self.master, state="readonly", values=["any", "all"], textvariable=self.how, width=5)
        self.checkbox2_how.current(0)

        # Clean+Save button
        clean_button = tk.Button(self.master, text="Clean!", command=self.clean, activeforeground="blue")
        clean_button.grid(row=7, column=0, columnspan=3, padx=10, pady=10)



    ### Functions for the cleaner page ###

    ## Toggle yes for checkbox fucntion
    def toggle_index(self):
        if self.index_check.get():
            self.master.geometry("625x240")
            self.index.config(text="(yes)")
            self.index_col_label.grid(row=3, column=2, sticky="w")
            self.index_col.grid(row=3, column=3, sticky="w")
        elif not self.index_check.get() and (self.output_choice.get() == ".sqlite" or self.clean_option2.get()):
            self.master.geometry("625x240")
            self.index.config(text="(no)")
            self.index_col_label.grid_forget()
            self.index_col.grid_forget()
        else:
            self.master.geometry("520x240")
            self.index.config(text="(no)")
            self.index_col_label.grid_forget()
            self.index_col.grid_forget()

    ## Toggle SQLite choice function
    def toggle_sqlite_choice(self, event):
        if self.output_choice.get() == ".sqlite":
            self.master.geometry("625x240")
            self.sqllabel.grid(row=4, column=2, sticky="w")
            self.tablename.grid(row=4, column=3,sticky="w")
        elif self.output_choice.get() != ".sqlite" and (self.clean_option2.get() or self.index_check.get()):
            self.master.geometry("625x240")
            self.sqllabel.grid_forget()
            self.tablename.grid_forget()
        else:
            self.master.geometry("520x240")
            self.sqllabel.grid_forget()
            self.tablename.grid_forget()


    ## Toggle how section function
    def toggle_how(self):
        if self.clean_option2.get():
            self.master.geometry("625x240")
            self.howlabel.grid(row=6, column=2, sticky="w")
            self.checkbox2_how.grid(row=6, column=3, sticky="w")
        elif not self.clean_option2.get() and (self.output_choice.get() == ".sqlite" or self.index_check.get()):
            self.master.geometry("625x240")
            self.howlabel.grid_forget()
            self.checkbox2_how.grid_forget()
        else:
            self.master.geometry("520x240")
            self.howlabel.grid_forget()
            self.checkbox2_how.grid_forget()

    ## Choosing an Excel file function
    def choose_file(self):
        file_choice = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_choice:
            self.file.config(state="normal")
            self.file.delete(0, tk.END)
            self.file.insert(0, file_choice)
            self.file.config(state="disabled")

    ## SQLite handling function
    def sqlite_convert(self, filename, tablename,df):
        ###creating sqlite database
        if tablename == "":
            table_name = "my_table"
        else:
            table_name = tablename
        try:
            connection = sqlite3.connect(filename)
        except Exception as e:
            messagebox.showerror('Error', f"Error occured while creating the SQLite database: {e}")
            return
        ###converting pandas df to sqlite table
        try:
            df.to_sql(table_name, connection, if_exists='replace', index=False)
            messagebox.showinfo("Success", f"Excel file converted to an SQLite database with a table \"{table_name}\" successfully!")
        except Exception as e:
            messagebox.showerror('Error', f"Error occured while converting the excel file to SQLite table: {e}")
            return
        finally:
            connection.commit()
            connection.close()

    ## Cleaning+Saving function
    def clean(self):
        if not all([self.file.get(), self.output_choice.get()]):
            if not self.file.get():
                self.filelabel.config(text="* Choose Excel file:")
                self.filelabel.config(fg="red")
            messagebox.showwarning("Warning", "Please choose a file!")
        else:
            self.filelabel.config(text="Choose Excel file:")
            self.filelabel.config(fg="white")

            # Throwing an error if the directory doesn't exitst
            if not os.path.exists(self.file.get()):
                messagebox.showerror("Error", "The directory or file does not exist!")
                return

            # Reading the original file and creating a new file with _cleaned appended and the chosen output type
            original_file_path = self.file.get()
            directory, filename = os.path.split(original_file_path)
            file_name_without_ext, file_ext = os.path.splitext(filename)
            cleaned_filename = f"{file_name_without_ext}_cleaned{self.output_choice.get()}"
            cleaned_file_path = os.path.join(directory, cleaned_filename)

            # Setting up pandas dataframe
            if self.index_check.get() == 0:
                df = pd.read_excel(original_file_path)
            else:
                df = pd.read_excel(original_file_path, index_col=int(self.index_col.get()))
            
            # Cleaning
            if self.clean_option1.get():
                df.drop_duplicates(inplace=True)
                df.reset_index(drop=True, inplace=True)
            if self.clean_option2.get():
                df.dropna(how=self.how.get(), inplace=True)
                df.reset_index(drop=True, inplace=True)

            # Saving file
            if self.output_choice.get() == ".xlsx":
                df.to_excel(cleaned_file_path, index=False)
                messagebox.showinfo("Success", f"Excel file cleaned and saved as \"{cleaned_filename}\" successfully!")
            elif self.output_choice.get() == ".csv":
                df.to_csv(cleaned_file_path, index=False)
                messagebox.showinfo("Success", f"Excel file converted to a CSV file and saved as \"{cleaned_filename}\" successfully!")
            elif self.output_choice.get() == ".sqlite":
                self.sqlite_convert(cleaned_file_path, self.tablename.get(), df)



#### GUI ####
gui = tk.Tk()
#### App ####
app = App(gui)
#### App Main Loop ####
app.master.mainloop()
