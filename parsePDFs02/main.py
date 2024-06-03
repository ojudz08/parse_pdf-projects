"""
    Author: Ojelle Rogero
    Created on: June 02, 2024
    About:
        Converts the Weekly Market Recap section of GSAM Market Monitor
        Parse each asset type and save as a data table in separate sheet
"""

import os, sys
from pypdf import PdfReader
from tkinter.filedialog import *

import pymupdf
from pprint import pprint




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
        page = read_pdf.pages[0]
        test = page.extract_text()   # resolved_objectts


        return test
    

    def test_pymudpdf(self):
        pdf_file = self.open_file()
        pdf = pymupdf.open(pdf_file)

        pg_cnt = pdf.page_count
        
        cnt = 0
        for pg in range(0, 11):
            page = pdf.load_page(pg)

            extract_pg_txt = page.get_text()
            search_pg = page.search_for("Common Stocks")
            if search_pg: 
                what_page = pg + 1
                break
            
            # create a dictionary
            # key Natixis Sustainable Future 2025 Fund -- March 31, 2024, February 29, 2024 and January 31, 2024
            # value - start to end page of each month end

            #bnd = page.bound()
            #pdf.get_page_text(0) -- texts of page 1, same as pdf.load_page(0).get_text()
            #pdf.load

            #if page.search_for("Common Stocks"):  
            #    search_pg = page.search_for("Common Stocks") # returns bounding box
            #    result = [pg, search_pg]

                
        return search_pg
      



if __name__ == '__main__':
    data = pdfParse().test_pymudpdf()
    


    print(data)
    