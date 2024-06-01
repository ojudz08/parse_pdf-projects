"""
    Author: Ojelle Rogero
    Created on: June 02, 2024
    About:
        Converts the Weekly Market Recap section of GSAM Market Monitor
        Parse each asset type and save as a data table in separate sheet
"""

import os, sys
from pypdf import PdfReader
import tkinter as tk
from tkinter.filedialog import *


class pdfParse():

    def __init__(self) -> None:
        self.init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))

    def open_file(self):
        open_file = askopenfilenames(initialdir=self.init_dir)
        filepath = open_file[0]   # read 1 file only
        return filepath

    def read_page(self):
        pdf_file = self.open_file()
        read_pdf = PdfReader(pdf_file)
        page = read_pdf.pages[0].extract_text()

        return page
    
    def page_reader(self):
        page = reader.pages[0]
        #lay = page.extract_text(extraction_mode="layout", layout_mode_scale_weight=1.0)
        return page
    
    def visitor_vody(text, cm, tm, font_dict, font_size):
        parts = []
        y = cm[5]
        if y > 50 and y < 720:
            parts.append(text)
        
        pass




if __name__ == '__main__':
    test = pdfParse().read_page()

    print(test)