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

    def __init__(self, filepath) -> None:
        self.filepath = filepath

    def month_end_pages(self, period):
        doc = pymupdf.open(self.filepath)
        pg_cnt = doc.page_count     

        month_page = []
        for pg in range(0, pg_cnt):
            if doc.load_page(pg).search_for(period): month_page.append(pg)
            
            pg_range = [min(month_page), max(month_page)]            
        return pg_range

    
    def asset_rect(self, search, pgrng):
        doc = pymupdf.open(self.filepath)

        for i in range(pgrng[0], pgrng[1] + 1):
            page = doc.load_page(i)
            
            if page.search_for(search):
                search_word = page.search_for(search)[0]
                search_page = i
                break
                
        return search_page, search_word


    def read_area(self, pg, rectarea):
        df = tabula.read_pdf(self.filepath, pages=pg, area=[rectarea], pandas_options={'header': None}, output_format='dataframe', stream=True)[0]
        return df
    

    def asset_pg(self, period_pg):
        assets = ["Common Stocks", "Bonds and Notes", "Exchange-Traded Funds", "Mutual Funds", "Affiliated Mutual Funds", "Short-Term Investments"]

        assets_pg = {}
        for i in range(0, len(assets)):
            asset_i = self.asset_rect(assets[i], period_pg)
            
            if i <= 3:
                asset_next = self.asset_rect(assets[i + 1], period_pg)
                pg_i, pg_nxt = asset_i[0], asset_next[0]
                rect_i, rect_next = round(asset_i[1][1], 2), round(asset_next[1][1], 2)
                
                if pg_nxt > pg_i and rect_next > rect_i:
                    assets_pg[assets[i]] = [asset_i[0], asset_next[0]]
                elif pg_nxt > pg_i and rect_next < 188:
                    assets_pg[assets[i]] = [asset_i[0], asset_i[0]]
                elif pg_nxt == pg_i:
                    assets_pg[assets[i]] = [asset_i[0], asset_i[0]]
            else:
                assets_pg[assets[i]] = [asset_i[0], asset_i[0]]
            
        return assets_pg
    
    
    def common_stocks(self, search_asset, period_pg):
        asset_pages = self.asset_pg(period_pg)
        
        for key, val in asset_pages.items():
            if key == search_asset: pgrng = val

        common_stocks_df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2
        for pg in range(strt, end):
            if pg == 1: top, bot = 130, 725            
            elif pg > 1 and pg < end - 1: top, bot = 117, 725
            elif pg == end - 1: top, bot = 117, 220

            rect_area = {"Shares": [top, 60, bot, 91], "Security Description": [top, 94, bot, 400], "Market Value ($)": [top, 420, bot, 487], f"% of Fund": [top, 490, bot, 540]}

            data, colrename = pd.DataFrame(), []
            for key, value in rect_area.items():
                df = self.read_area(pg, value)
                data = pd.concat([data, df], axis=1, ignore_index=True)
                colrename.append(key)
        
            common_stocks_df = pd.concat([common_stocks_df, data], axis=0, ignore_index=True)

        common_stocks_df.columns = colrename
        return common_stocks_df
    

    def get_data(self):
        period = "March 31, 2024"
        period_pg = self.month_end_pages(period)

        common_stocks_df = self.common_stocks("Common Stocks", period_pg)

        return common_stocks_df




if __name__ == '__main__':
    #month_end = ["January 31, 2024", "February 29, 2024", "March 31, 2024"]
    init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
    open_file = askopenfilenames(initialdir = init_dir)
    filepath = open_file[0]    
     
    data = pdfParse(filepath).get_data()
    print(data)
    