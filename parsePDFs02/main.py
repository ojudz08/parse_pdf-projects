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
            
            # from month_end_pages, search for Common Stocks and parse from there
                
        return search_pg
    
    def month_end_pages(self):
        pdf_file = self.open_file()
        pdf = pymupdf.open(pdf_file)

        month_end = ["January 31, 2024", "February 29, 2024", "March 31, 2024"]
        month_pg = {}
        pg_cnt = pdf.page_count     

        for month in month_end:
            month_page = []

            for pg in range(0, pg_cnt):
                if pdf.load_page(pg).search_for(month): month_page.append(pg)
            
            month_pg[month] = [min(month_page), max(month_page)]

        return month_pg
      



if __name__ == '__main__':
    data = pdfParse().month_end_pages()

    print(data)
    