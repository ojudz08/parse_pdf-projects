import os, sys
from tkinter.filedialog import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pypdf import PdfWriter, PdfReader
import re


def open_pdf():
    """Open file and returns filepath"""
    open_file = askopenfilenames(initialdir = init_dir)
    filename = os.path.basename(open_file[0])

    open_file_entry.delete(0, tk.END)
    open_file_entry.insert(tk.END, filename)


def save_as(filename):
    """Save the parsed pdf to desired target folder"""
    file_types = [('pdf', '*.pdf'), ('All Files', '*.*')]
    save_file_as = asksaveasfilename(initialdir = init_dir, initialfile = filename, filetypes = file_types, defaultextension = '.pdf', confirmoverwrite = True)
    if save_file_as is None:
        return None
    else:
        return save_file_as


def splitter():
    filename = str(open_file_entry.get())
    filepath = os.path.join(init_dir, "reports", filename)
 
    pdf = PdfReader(filepath)
    pdf_writer = PdfWriter()

    temp = page_entry.get()
    comma, hyphen = re.compile(","), re.compile("-")

    if any(True for c in temp if c in ","):
        page = [int(x) for x in comma.split(temp) if x]
    elif any(True for c in temp if c in "-"):
        page = list(range(int(hyphen.split(temp)[0]), int(hyphen.split(temp)[1]) + 1))

    try:
        for pg in page:
            pdf_writer.add_page(pdf.get_page(pg - 1))
        
        output_file = save_as(filename)

        with open(output_file, 'wb') as output:
            pdf_writer.write(output)
    except:
        messagebox.showerror("Error", "Invalid page number")


if __name__ == '__main__':
    init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))


    window = tk.Tk()
    window.geometry('550x100')
    window.title("PDF Splitter")
    window.resizable(0, 0)

    pages = tk.StringVar()
    filename = tk.StringVar()

    window.columnconfigure(0, weight=1)
    window.columnconfigure(1, weight=3)

    open_file_label = ttk.Label(window, text="Filename: ")
    open_file_label.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

    open_file_entry = ttk.Entry(window, width=60, textvariable=filename)
    open_file_entry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

    open_button = ttk.Button(window, text="Open File", command=open_pdf)
    open_button.grid(column=2, row=0, sticky=tk.E, padx=5, pady=5)


    page_label = ttk.Label(window, text="Page/s to split: ")
    page_label.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)

    page_entry = ttk.Entry(window, width=60, textvariable=pages)
    page_entry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

    page_entry.delete(0, tk.END)

    submit_button = ttk.Button(window, text="Submit", command=splitter)
    submit_button.grid(column=2, row=1, sticky=tk.E, padx=5, pady=5)

    window.mainloop()
    
    