"""
    Author: Ojelle Rogero
    Created on: June 02, 2024
    About:
        Parses the different assets in Natixis Sustainable Future Fund for the latest period
        Save data table as an excel file in each separate sheet
"""

import os, sys
from tkinter.filedialog import *
import tabula
import pymupdf
import pandas as pd




class pdfParse():

    def __init__(self, period) -> None:
        self.init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
        tempfile = self.open_pdf(period)
        self.filepath = tempfile
        self.period = period


    def save_as(self):
        """Save the parsed pdf to desired target folder"""
        file_types = [('Excel Files', '*.xlsx'), ('All Files', '*.*')]
        outputfilename = os.path.basename(self.filepath)[:-4]
        save_file_as = asksaveasfilename(initialdir = self.init_dir, initialfile = outputfilename, filetypes = file_types, defaultextension = '.xlsx', confirmoverwrite = True)
        if save_file_as is None:
            return None
        else:
            return save_file_as
    

    def open_pdf(self):
        """Open file and returns filepath"""
        open_file = askopenfilenames(initialdir = self.init_dir)
        filepath = open_file[0]
        return filepath


    def month_end_pages(self, period):
        """Get page range from the specified period.
           Returns the page range"""
        doc = pymupdf.open(self.filepath)
        pg_cnt = doc.page_count     

        month_page = []
        for pg in range(0, pg_cnt):
            if doc.load_page(pg).search_for(period): month_page.append(pg)
            
            pg_range = [min(month_page), max(month_page)]            
        return pg_range

    
    def asset_rect(self, search, pgrng):
        """Get the coordinates of the asset for a specified page range.
           Returns the page where the asset is found and bounding box or rectangle"""
        doc = pymupdf.open(self.filepath)

        for i in range(pgrng[0], pgrng[1] + 1):
            page = doc.load_page(i)
            
            if page.search_for(search):
                search_word = page.search_for(search)[0]
                search_page = i
                break
                
        return search_page, search_word


    def read_area(self, pg, rectarea):
        """Parse the area of the specified bounding box.
           Returns dataframe df"""
        df = tabula.read_pdf(self.filepath, pages = pg, area = [rectarea], pandas_options = {'header': None}, output_format = 'dataframe', stream = True)[0]
        return df
    

    def asset_pg(self, period_pg):
        """Get the pages for each asset for the specified period page range.
           Returns a dictionary of the assets and page => {"Asset": [start, end]}
        """
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
    
    
    def concat_cols(self, rect_area, pg):
        """Concatenates the asset column into one dataframe"""
        data, colrename = pd.DataFrame(), []
        for key, value in rect_area.items():
            df = self.read_area(pg, value)
            data = pd.concat([data, df], axis=1, ignore_index=True)
            colrename.append(key)
        
        return data, colrename


    def get_rect_area(self, asset, top, bot):
        """Rectange Area for each asset"""

        if asset in ('Common Stocks', 'Exchange-Traded Funds', 'Mutual Funds', 'Affiliated Mutual Funds'):
            rect_area = {"Shares": [top, 60, bot, 91],
                         "Security Description": [top, 94, bot, 330],
                         "Market Value ($)": [top, 425, bot, 487],
                         f"% of Fund": [top, 490, bot, 540]}

        elif asset in ('Bonds and Notes', 'Short-Term Investments'):
            rect_area = {"Principal Amount": [top, 60, bot, 91],
                         "Security Description": [top, 94, bot, 330],
                         "Interest Rate": [top, 332, bot, 370],
                         "Maturity Date": [top, 373, bot, 420],
                         "Market Value ($)": [top, 425, bot, 487],
                         f"% of Fund": [top, 490, bot, 540]}

        return rect_area
    

    def common_stocks(self, asset, page_range):
        """Parse Common Stocks"""
        pgrng = page_range

        df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2

        for pg in range(strt, end):
            if pg == 1: top, bot = 130, 725            
            elif pg > 1 and pg < end - 1: top, bot = 117, 725
            elif pg == end - 1: top, bot = 117, 220

            rect_area = self.get_rect_area(asset, top, bot)

            temp = self.concat_cols(rect_area, pg)
            data, colrename = temp[0], temp[1]
            df = pd.concat([df, data], axis=0, ignore_index=True)

        df.columns = colrename
        return df
    

    def bonds_and_notes(self, asset, page_range):
        """Parse Bonds and Notes"""
        pgrng = page_range

        df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2

        for pg in range(strt, end):
            if pg == strt: top, bot = 250, 725       # add a function to derive top 250     
            elif pg > strt and pg < end - 1: top, bot = 117, 725
            elif pg == end - 1: top, bot = 117, 630  # add a function to derive bot 
            
            rect_area = self.get_rect_area(asset, top, bot)

            temp = self.concat_cols(rect_area, pg)
            data, colrename = temp[0], temp[1]
            df = pd.concat([df, data], axis=0, ignore_index=True)

        df.columns = colrename
        return df


    def exchange_traded_funds(self, asset, page_range):
        """Parse Exchange-Traded Funds"""
        pgrng = page_range

        df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2

        for pg in range(strt, end):
            if pg == strt: top, bot = 660, 725

            rect_area = self.get_rect_area(asset, top, bot)
            
            temp = self.concat_cols(rect_area, pg)
            data, colrename = temp[0], temp[1]
            df = pd.concat([df, data], axis=0, ignore_index=True)

        df.columns = colrename
        return df


    def mutual_funds(self, asset, page_range):
        """Parse Mutual Funds"""
        pgrng = page_range

        df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2

        for pg in range(strt, end):
            if pg == strt: top, bot = 134, 185   # add function to derive both

            rect_area = self.get_rect_area(asset, top, bot)
            
            temp = self.concat_cols(rect_area, pg)
            data, colrename = temp[0], temp[1]
            df = pd.concat([df, data], axis=0, ignore_index=True)

        df.columns = colrename
        return df
    

    def affiliated_mutual_funds(self, asset, page_range):
        """Parse Affiliated Mutual Funds"""
        pgrng = page_range

        df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2

        for pg in range(strt, end):
            if pg == strt: top, bot = 199, 300   # add function to derive both

            rect_area = self.get_rect_area(asset, top, bot)
            
            temp = self.concat_cols(rect_area, pg)
            data, colrename = temp[0], temp[1]
            df = pd.concat([df, data], axis=0, ignore_index=True)

        df.columns = colrename
        return df
    

    def short_term_investments(self, asset, page_range):
        """Short Term Investments"""
        pgrng = page_range

        df = pd.DataFrame()
        strt, end = pgrng[0] + 1, pgrng[1] + 2

        for pg in range(strt, end):
            if pg == strt: top, bot = 322, 390   # add function to derive both

            rect_area = self.get_rect_area(asset, top, bot)
            
            temp = self.concat_cols(rect_area, pg)
            data, colrename = temp[0], temp[1]
            df = pd.concat([df, data], axis=0, ignore_index=True)

        df.columns = colrename
        return df
    

    def get_data(self):
        """Get the data for each asset and save into an excel file"""
        period_pg = self.month_end_pages(self.period)

        asset_pages = self.asset_pg(period_pg)
        for key, val in asset_pages.items():
            if key == "Common Stocks":
                asset1 = self.common_stocks(key, val)
            elif key == "Bonds and Notes":
                asset2 = self.bonds_and_notes(key, val)
            elif key == "Exchange-Traded Funds":
                asset3 = self.exchange_traded_funds(key, val)
            elif key == "Mutual Funds":
                asset4 = self.mutual_funds(key, val)
            elif key == "Affiliated Mutual Funds":
                asset5 = self.affiliated_mutual_funds(key, val)
            elif key == "Short-Term Investments":
                asset6 = self.short_term_investments(key, val)

        output_file = self.save_as()

        with pd.ExcelWriter(output_file) as writer:
            asset1.to_excel(writer, sheet_name="Common Stocks", index=False)
            asset2.to_excel(writer, sheet_name="Bonds and Notes", index=False)
            asset3.to_excel(writer, sheet_name="Exchange-Traded Funds", index=False)
            asset4.to_excel(writer, sheet_name="Mutual Funds", index=False)
            asset5.to_excel(writer, sheet_name="Affiliated Mutual Funds", index=False)
            asset6.to_excel(writer, sheet_name="Short-Term Investments", index=False)
            
        return "Done!"




if __name__ == '__main__':
    #month_end = ["January 31, 2024", "February 29, 2024", "March 31, 2024"]
    
    period = "March 31, 2024"
    data = pdfParse(period).get_data()
    print(data)
    