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

import tabula
import pymupdf
from pprint import pprint
import pandas as pd




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


    def read_area(self, filepath, pg, rectarea):
        df = tabula.read_pdf(filepath, pages=pg, area=[rectarea], pandas_options={'header': None}, output_format='dataframe', stream=True)[0]
        return df


    def common_stocks(self):
        file_open = self.open_file()
        filepath = file_open
        doc = pymupdf.open(filepath)

        search =  "Common Stocks"
        pgrng = [0, 10]
        get_page_word = self.asset_rect(filepath, search, pgrng)

        pag, rect = get_page_word[0], get_page_word[1]

        page = doc.load_page(pag)
        cm_result = page.get_textbox(rect[0])
        
        common_stocks_df = pd.DataFrame()
        strt, end = pgrng[0] + 1, 9
        for pg in range(strt, 9):
            if pg == 1: top, bot = 130, 725            
            elif pg > 1 and pg < 8: top, bot = 117, 725
            elif pg == 8: top, bot = 117, 220

            rect_area = {"Shares": [top, 60, bot, 91], "Security Description": [top, 94, bot, 400], "Market Value ($)": [top, 420, bot, 487], f"% of Fund": [top, 490, bot, 540]}

            data, colrename = pd.DataFrame(), []
            for key, value in rect_area.items():
                df = self.read_area(filepath, pg, value)
                data = pd.concat([data, df], axis=1, ignore_index=True)
                colrename.append(key)
        
            common_stocks_df = pd.concat([common_stocks_df, data], axis=0, ignore_index=True)

        common_stocks_df.columns = colrename

        return common_stocks_df
    

    def bonds_and_notes(self):
        file_open = self.open_file()
        filepath = file_open
        doc = pymupdf.open(filepath)

        search =  "Bonds and Notes"
        pg_range = [0, 10]
        get_page_word = self.asset_rect(filepath, search, pg_range)

        pg, rect = get_page_word[0], get_page_word[1]

        return pg, rect

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
    