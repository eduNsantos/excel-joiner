import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import *
import pandas as p
import re


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.columns = []
        self.column_types = {
            'General': 'Geral',
            'Numeric': 'Número',
            'Date': 'Data'
        }
        self.field_options = ('Somar', 'Único', 'Deixar lado a lado')
        self.selected_files = 'Nenhum arquivo selecionado'
        self.request_excel_files()

    def create_widgets(self):
        self.create_headers()
        self.create_select_files_button()

        for excel in self.excels:
            columns = excel.columns.ravel()

            i = 0
            row = 3
            while i < len(columns):
                tk.Label(self.master, text=columns[i]).grid(row=row, column=1, sticky="w", pady=5, padx=20)

                default_column_type = self.verify_value_type(excel[columns[i]][0])

                column_type = Combobox(self.master, values=[self.column_types[column] for column in self.column_types])
                column_type.set(self.column_types.get(default_column_type))
                column_type.grid(row=row, column=2, sticky="w", padx=20)

                field_option = Combobox(self.master, values=self.field_options)
                field_option.grid(row=row, column=3, sticky="w", padx=20)

                if default_column_type == "Numeric":
                    field_option.current(0)
                elif default_column_type == "General":
                    field_option.current(1)
                else:
                    field_option.current(2)
                i += 1
                row += 1

    def request_excel_files(self):
        files = []
        while len(files) == 0:
            file_dialog = tk.filedialog
            files = file_dialog.askopenfilename(multiple=True, title="Selecione os arquivos excel", filetypes=[("Excel files", '.xlsx', '.xls')])
        for widget in self.master.winfo_children():
            widget.destroy()
        self.file_names = [file[file.rindex('/')+1:] + ' | ' for file in files]
        self.excels = [p.read_excel(file) for file in files]
        self.create_widgets()


    def create_select_files_button(self):
        self.lblSelectedFiles = tk.Label(self.master, text = self.selected_files)
        self.lblSelectedFiles.grid(padx=20, row=1, column=0, columnspan=3, sticky="w")

        button = tk.Button(self.master, text="Selecione os arquivos para fazer a leitura")
        button['command'] = self.request_excel_files
        button.grid(row=0, column=0, sticky="w", columnspan=3, padx=20)

    def create_headers(self):
        tk.Label(self.master, text="Campo",font=('Helvetica', 12)).grid(row=2, column=1, sticky="w", padx=20)
        tk.Label(self.master, text="Tipo do campo", font=('Helvetica', 12)).grid(row=2, column=2, sticky="w", padx=20)
        tk.Label(self.master, text="Opção", font=('Helvetica', 12)).grid(row=2, column=3, sticky="w", padx=20)

    def verify_value_type(self, value):
        if re.match("[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1]) (2[0-3]|[01][0-9]):[0-5][0-9]", str(value)):
            default_column_type = 'Date'
        elif re.match("[0-9]", str(value)):
            default_column_type = 'Numeric'
        else:
            default_column_type = 'General'
        return default_column_type


root = tk.Tk()
app = Application(master=root)
app.mainloop()