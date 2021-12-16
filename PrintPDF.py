###
#
#
#     /$$$$$                     /$$
#    |__  $$                    | $$
#       | $$  /$$$$$$   /$$$$$$$| $$$$$$$  /$$   /$$  /$$$$$$
#       | $$ /$$__  $$ /$$_____/| $$__  $$| $$  | $$ |____  $$
#  /$$  | $$| $$  \ $$|  $$$$$$ | $$  \ $$| $$  | $$  /$$$$$$$
# | $$  | $$| $$  | $$ \____  $$| $$  | $$| $$  | $$ /$$__  $$
# |  $$$$$$/|  $$$$$$/ /$$$$$$$/| $$  | $$|  $$$$$$/|  $$$$$$$
#  \______/  \______/ |_______/ |__/  |__/ \______/  \_______/
#
#
#
#  /$$      /$$
# | $$  /$ | $$
# | $$ /$$$| $$  /$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$$
# | $$/$$ $$ $$ |____  $$ /$$__  $$ /$$__  $$ /$$__  $$| $$__  $$
# | $$$$_  $$$$  /$$$$$$$| $$  \__/| $$  \__/| $$$$$$$$| $$  \ $$
# | $$$/ \  $$$ /$$__  $$| $$      | $$      | $$_____/| $$  | $$
# | $$/   \  $$|  $$$$$$$| $$      | $$      |  $$$$$$$| $$  | $$
# |__/     \__/ \_______/|__/      |__/       \_______/|__/  |__/
#
# for any questions, email me
# warrenj@rushenterprises.com
#
#

import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
import win32com.client
import pywintypes
import threading
from threading import Thread
import sys

printer_list = []

objFSO = win32com.client.Dispatch("Scripting.FilesystemObject")

class Window(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.master = master
        self.init_window()

    def init_window(self):
        self.master.title("Mass Print PDF")
        self.pack(fill = 'both', expand = 1)

        self.pdf_file_path = tk.StringVar()
        self.printer_file = tk.StringVar()

        pdfButton = ttk.Button(self, text='Get PDF', command=self.get_pdf)
        pdfButton.grid(column=0, row=3, padx=5, pady=5)

        self.printerButton = ttk.Button(self, text='Get List of Printers', command=self.get_printers)

        self.executeButton = ttk.Button(self, text='Execute', command=self.show_pb)

        exitButton = ttk.Button(self, text='Exit', command=self.close_window)
        exitButton.grid(column=1, row=4, padx=5, pady=5)

        pdf_file_path = ttk.Entry(self, textvariable = self.pdf_file_path, width=30)
        pdf_file_path.grid(column=1, row=0, padx=5, pady=0)
        pdfLabel = ttk.Label(self, text="PDF File")
        pdfLabel.grid(column=0, row=0, padx=5, pady=0)

        printer_file_path = ttk.Entry(self, textvariable = self.printer_file, width=30)
        printer_file_path.grid(column=1, row=2, padx=5, pady=0)
        printerLabel = ttk.Label(self, text="Printer File")
        printerLabel.grid(column=0, row=2)

        self.progress_bar= ttk.Progressbar(self, orient='horizontal', mode="determinate", length=100, value=0)

    def close_window(self):
        form.destroy()

    def get_pdf_browser(self):
        self.filename = fd.askopenfilename(filetypes=[("PDF", "*.pdf")])
        return self.filename

    def get_pdf(self):
        file = self.get_pdf_browser()
        self.pdf_file_path.set(file)
        self.pdf_file = file
        self.printerButton.grid(column=1, row=3, padx=5, pady=5)

    def get_printer_file(self):
        self.filename = fd.askopenfilename(filetypes=[("Text", "*.txt")])
        return self.filename

    def get_printers(self):
        file = self.get_printer_file()
        self.printer_file.set(file)
        self.x = 0
        with open(file) as file:
            for line in file:
                printer_list.append(line.strip())
                self.x += 1
        self.executeButton.grid(column=0, row=4, padx=5, pady=5)

    def show_pb(self):
        self.progress_bar.grid(column=1, row=5, padx=5, pady=5)
        threading.Thread(target=self.print_the_pdf).start()

    def print_the_pdf(self):
        n = 100/self.x
        for p in printer_list:
            try:
                objFSO.CopyFile(self.pdf_file, p)
                self.progress_bar.step(n)
                print(p)
            except pywintypes.com_error:
                print("Error, could not print to ", p)
                self.progress_bar.step(n)
        tk.messagebox.showinfo(title=None, message='PDF file has been printed')

form = tk.Tk()
form.geometry('300x150')
app = Window(form)

form.mainloop()
