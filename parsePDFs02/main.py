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
        open_file = askopenfilenames(initialdir = self.init_dir)
        filepath = open_file[0]   # read 1 file only
        return filepath
        
    def month_end_pages(self, filepath):
        doc = pymupdf.open(filepath)

        month_end = ["January 31, 2024", "February 29, 2024", "March 31, 2024"]
        month_pg = {}
        pg_cnt = doc.page_count     

        for month in month_end:
            month_page = []
            for pg in range(0, pg_cnt):
                if doc.load_page(pg).search_for(month): month_page.append(pg)
            
            month_pg[month] = [min(month_page), max(month_page)]
        return month_pg

    
    def asset_rect(self, filepath, search, pg):
        doc = pymupdf.open(filepath)

        for i in range(pg[0], pg[1] + 1):
            page = doc.load_page(i)
            
            if page.search_for(search):
                search_word = page.search_for(search)
                search_page = i
                break
        return search_page, search_word


    def common_stocks(self):
        file_open = self.open_file()
        filepath = file_open
        doc = pymupdf.open(filepath)

        search =  "Common Stocks"
        pg_range = [0, 10]
        get_page_word = self.asset_rect(filepath, search, pg_range)

        pg, rect = get_page_word[0], get_page_word[1]

        page = doc.load_page(pg)
        cm_result = page.get_textbox(rect[0])
        
        # from the Common Stocks rect, parse the table from start to end
        # page.rect, page.bound

        return rect
    

    def bonds_and_notes(self):
        filepath = file_open
        doc = pymupdf.open(filepath)

        search =  "Bonds and Notes"
        pg_range = [0, 10]
        get_page_word = self.asset_rect(filepath, search, pg_range)

        pg, rect = get_page_word[0], get_page_word[1]
        pass

    def exchange_traded_funds(self):
        pass

    def mutual_funds(self):
        pass

    def affiliated_mutual_funds(self):
        pass

    def short_term_investments(self):
        pass
      



if __name__ == '__main__':
    data = pdfParse().common_stocks()

    print(data)
    