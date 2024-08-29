import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class App:
    def __init__(self, master):
        self.master = master
        self.main_page()

    def main_page(self):
        for i in self.master.winfo_children():
            i.destroy()
        self.master.title("Excel Master")
        self.master.geometry("280x300")
        tk.Label(self.master, text="Welcome to Excel Master!", font=("Times","24"), fg="#5ea832").grid(row=0, column=0, padx=10, pady=10)
        tk.Label(self.master, text="Choose what you want to do:").grid(row=1, column=0, padx=10)
        self.clean_button = tk.Button(self.master, text="Excel Cleaner", command=self.clean_page, activeforeground="blue")
        self.clean_button.grid(row=2, column=0, padx=10, pady=10)

    # Cleaner page function
    def clean_page(self):
        for i in self.master.winfo_children():
            i.destroy()
        self.master.geometry("500x210")
        
        # Back button
        back_button = tk.Button(self.master, text="<<<", command=self.main_page, cursor="hand2", activeforeground="blue")
        back_button.grid(row=0, column=0, padx=10, sticky="w")

        # Cleaner page title
        tk.Label(self.master, text = "Excel Cleaner", font=("Times","24"), fg="#5ea832").grid(row=1, column=1, padx=10)

        # Choosing a file with button
        self.filelabel = tk.Label(self.master, text="Choose xlsx file:", font=("Times","15"), fg="#5ea832")
        self.filelabel.grid(row=2, column=0, padx=5, sticky="w")
        self.file = tk.Entry(self.master, state="disabled", bg="#74e8f2")
        self.file.grid(row=2, column=1, sticky="w")
        file_button = tk.Button(self.master, text="Choose .xlsx file", command=self.choose_file, activeforeground="blue")
        file_button.grid(row=2, column=2, sticky="w")
    
        # Choosing preferred output file in a dropdown menu
        tk.Label(self.master, text="Preferred output:", font=("Times","15"), fg="#5ea832").grid(row=3, column=0, padx=5,sticky="w")
        self.n = tk.StringVar()
        self.output_choice = ttk.Combobox(self.master, state="readonly", values=[".xlsx", ".csv"], textvariable=self.n)
        self.output_choice.grid(row=3, column=1, sticky="w")
        self.output_choice.current(0)

        # Cleaning options
        tk.Label(self.master, text="Cleaning options:", font=("Times","15"), fg="#5ea832").grid(row=4, column=0, padx=5,sticky="w")
        self.clean_option1 = tk.IntVar()
        self.clean_option2 = tk.IntVar()
        checkbox1 = tk.Checkbutton(self.master, text="Drop Duplicates", onvalue=1, offvalue=0, variable=self.clean_option1)
        checkbox2 = tk.Checkbutton(self.master, text="Drop empty rows", onvalue=1, offvalue=0, variable=self.clean_option2, command=self.toggle_how)
        checkbox1.grid(row=4, column=1, sticky="w")
        checkbox2.grid(row=5, column=1, sticky="w")
        
        self.howlabel = tk.Label(self.master, text="how?", width=5)
        self.how = tk.StringVar()
        self.checkbox2_how = ttk.Combobox(self.master, state="readonly", values=["any", "all"], textvariable=self.how, width=5)
        self.checkbox2_how.current(0)
        # Clean button
        clean_button = tk.Button(self.master, text="Clean!", command=self.clean, activeforeground="blue")
        clean_button.grid(row=6, column=1, columnspan=1, padx=10)
        














    def toggle_how(self):
        if self.clean_option2.get():
            self.master.geometry("550x210")
            self.howlabel.grid(row=5, column=2, sticky="w")
            self.checkbox2_how.grid(row=5, column=3, sticky="w")
        else:
            self.master.geometry("500x210")
            self.howlabel.grid_forget()
            self.checkbox2_how.grid_forget()


    def choose_file(self):
        file_choice = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_choice:
            self.file.config(state="normal")
            self.file.delete(0, tk.END)
            self.file.insert(0, file_choice)
            self.file.config(state="disabled")

    def clean(self):
        if not all([self.file.get(), self.output_choice.get()]):
            if not self.file.get():
                self.filelabel.config(text="* Choose xlsx file:")
                self.filelabel.config(fg="red")
            messagebox.showwarning("Warning", "Please choose a file!")
        else:
            self.filelabel.config(text="Choose xlsx file:")
            self.filelabel.config(fg="white")

            # Reading the original file
            original_file_path = self.file.get()
            directory, filename = os.path.split(original_file_path)
            file_name_without_ext, file_ext = os.path.splitext(filename)
            cleaned_filename = f"{file_name_without_ext}_cleaned{self.output_choice.get()}"
            cleaned_file_path = os.path.join(directory, cleaned_filename)

            df = pd.read_excel(original_file_path)

            ##########nemám ošetrený index
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
            elif self.output_choice.get() == ".csv":
                df.to_csv(cleaned_file_path, index=False)
            messagebox.showinfo("Success", f"File saved as \"{cleaned_filename}\" to ({cleaned_file_path})")

gui = tk.Tk()
app = App(gui)
app.master.mainloop()
